[CmdletBinding()]
param(
    [switch]$Install,
    [Parameter(ValueFromRemainingArguments=$true)][string[]]$DrawingFiles
)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$libraryPath = (Split-Path -Parent $PSScriptRoot).TrimEnd('\')
$visio = $null
$createdApplication = $false
$openedDrawing = $false
Add-Type -AssemblyName System.Windows.Forms

function Assert-Library {
    if ($libraryPath -eq [IO.Path]::GetPathRoot($libraryPath).TrimEnd('\')) {
        throw 'Use a separate GraFiS folder, not a drive root.'
    }
    $report = Get-Content -LiteralPath (Join-Path $libraryPath 'binary-rewrite-report.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    foreach ($entry in $report) {
        $path = Join-Path $libraryPath $entry.file
        if ((Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash.ToLowerInvariant() -ne $entry.output_sha256) {
            throw "Library file differs from the verified package: $($entry.file)"
        }
    }
}

function Get-DrawingWindow($Document) {
    foreach ($window in $visio.Windows) {
        if ($window.Type -eq 1 -and $window.Document.ID -eq $Document.ID) { return $window }
    }
    throw 'Drawing window not found.'
}

function Add-LibraryTrust {
    if ($libraryPath.StartsWith('\\')) { throw 'Install the library in a local folder.' }
    $majorVersion = ([string]$visio.Version).Split('.')[0]
    $version = $majorVersion + '.0'
    $locations = "HKCU:\Software\Microsoft\Office\$version\Visio\Security\Trusted Locations"
    $folder = $libraryPath + '\'
    New-Item -Path $locations -Force | Out-Null
    $matching = $null
    foreach ($key in Get-ChildItem -LiteralPath $locations) {
        $value = Get-ItemProperty -LiteralPath $key.PSPath
        if ($null -ne $value.PSObject.Properties['Path']) {
            if ($value.Path.TrimEnd('\').Equals($libraryPath, [StringComparison]::OrdinalIgnoreCase)) {
                $matching = $key.PSPath; break
            }
        }
    }
    if ($null -eq $matching) {
        $index = 0
        while (Test-Path -LiteralPath (Join-Path $locations "Location$index")) { $index++ }
        $matching = Join-Path $locations "Location$index"
        New-Item -Path $matching | Out-Null
        New-ItemProperty -LiteralPath $matching -Name Path -Value $folder -PropertyType String | Out-Null
        New-ItemProperty -LiteralPath $matching -Name AllowSubFolders -Value 0 -PropertyType DWord | Out-Null
        New-ItemProperty -LiteralPath $matching -Name Description -Value 'GraFiS verified library folder' -PropertyType String | Out-Null
    }
    $report = Get-Content -LiteralPath (Join-Path $libraryPath 'binary-rewrite-report.json') -Raw -Encoding UTF8 | ConvertFrom-Json
    foreach ($entry in $report) { Unblock-File -LiteralPath (Join-Path $libraryPath $entry.file) }
}

try {
    Assert-Library
    try { $visio = [Runtime.InteropServices.Marshal]::GetActiveObject('Visio.Application') }
    catch { $visio = New-Object -ComObject Visio.Application; $createdApplication = $true }
    if ($Install) {
        Add-LibraryTrust
        $shell = New-Object -ComObject WScript.Shell
        $shortcut = $shell.CreateShortcut((Join-Path ([Environment]::GetFolderPath('Desktop')) 'GraFiS.lnk'))
        $shortcut.TargetPath = Join-Path $libraryPath 'Start-GraFiS.cmd'
        $shortcut.WorkingDirectory = $libraryPath
        $shortcut.Description = 'Open drawings with the GraFiS library in the Shapes pane'
        $shortcut.Save()
        [Windows.Forms.MessageBox]::Show('Установка завершена. Открывайте схемы через ярлык GraFiS на рабочем столе. Папка библиотеки добавлена в доверенные расположения Visio. Корпоративные политики могут ограничивать этот параметр.', 'GraFiS') | Out-Null
    } else {
        if (-not $DrawingFiles -or $DrawingFiles.Count -eq 0) {
            $dialog = New-Object Windows.Forms.OpenFileDialog
            $dialog.Title = 'Открыть схемы с библиотекой GraFiS'
            $dialog.Filter = 'Схемы Visio|*.vsd;*.vsdx;*.vsdm;*.vdx'
            $dialog.Multiselect = $true
            if ($dialog.ShowDialog() -ne [Windows.Forms.DialogResult]::OK) { return }
            $DrawingFiles = $dialog.FileNames
        }
        $manifest = Get-Content -LiteralPath (Join-Path $libraryPath 'manifest.json') -Raw -Encoding UTF8 | ConvertFrom-Json
        $visio.Visible = $true
        foreach ($file in $DrawingFiles) {
            $drawingPath = (Resolve-Path -LiteralPath $file).Path
            if ([IO.Path]::GetExtension($drawingPath).ToLowerInvariant() -notin @('.vsd','.vsdx','.vsdm','.vdx')) {
                throw "Select a drawing, not a stencil: $drawingPath"
            }
            $drawing = $null
            foreach ($doc in $visio.Documents) {
                if ($doc.FullName.Equals($drawingPath, [StringComparison]::OrdinalIgnoreCase)) { $drawing = $doc; break }
            }
            # Ignore stale saved window/workspace paths; the drawing content is unchanged.
            if ($null -eq $drawing) { $drawing = $visio.Documents.OpenEx($drawingPath, [int16]256) }
            $openedDrawing = $true
            $window = Get-DrawingWindow $drawing
            $window.Activate()
            $central = $visio.Documents.OpenEx((Join-Path $libraryPath 'Моделирование.vss'), [int16](2+4+8+256))
            foreach ($project in $manifest.projects) {
                $window.Activate()
                $stencil = $visio.Documents.OpenEx((Join-Path $libraryPath $project.stencil), [int16](2+4+8+256))
            }
            $window.Activate()
        }
    }
} catch {
    [Windows.Forms.MessageBox]::Show($_.Exception.Message, 'GraFiS: ошибка', [Windows.Forms.MessageBoxButtons]::OK, [Windows.Forms.MessageBoxIcon]::Error) | Out-Null
    throw
} finally {
    if ($createdApplication -and -not $openedDrawing -and $null -ne $visio) {
        try { $visio.Quit() } catch { }
    }
}
