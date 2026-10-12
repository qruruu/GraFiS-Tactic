[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$SourceDirectory,
    [Parameter(Mandatory=$true)][string]$OutputDirectory,
    [string]$PackageDirectory = (Split-Path -Parent $PSScriptRoot)
)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

function Get-CodeHash([string]$Code) {
    $lines = $Code -split '\r?\n' | ForEach-Object { $_.TrimEnd() }
    $canonical = (($lines | Where-Object { $_ -notmatch '^Attribute ' }) -join "`n").Trim()
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($canonical)))).Replace('-','').ToLowerInvariant() }
    finally { $sha.Dispose() }
}
function Get-SourceCode([string]$Path, [string]$EncodingName) {
    $text = [IO.File]::ReadAllText($Path, [Text.Encoding]::GetEncoding($EncodingName))
    $lines = $text -split '\r?\n'
    $start = 0
    for ($i=0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match '^Attribute VB_Name') { $start = $i; break }
    }
    return (($lines[$start..($lines.Count-1)] | Where-Object { $_ -notmatch '^Attribute ' }) -join "`r`n").Trim()
}

$source = (Resolve-Path -LiteralPath $SourceDirectory).Path.TrimEnd('\')
$output = [IO.Path]::GetFullPath($OutputDirectory).TrimEnd('\')
if (-not (Test-Path -LiteralPath $source -PathType Container)) { throw 'SourceDirectory must be a library folder.' }
if ($source -eq [IO.Path]::GetPathRoot($source).TrimEnd('\')) { throw 'Select the GraFiS folder, not a drive root.' }
if ($output.Equals($source, [StringComparison]::OrdinalIgnoreCase) -or $output.StartsWith($source + '\', [StringComparison]::OrdinalIgnoreCase)) {
    throw 'OutputDirectory must be a separate new folder outside SourceDirectory.'
}
if (Test-Path -LiteralPath $output) { throw 'OutputDirectory already exists. Choose a new folder; nothing will be overwritten.' }
$manifest = Get-Content -LiteralPath (Join-Path $PackageDirectory 'manifest.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$selected = @($manifest.projects | Where-Object { Test-Path -LiteralPath (Join-Path $source $_.stencil) -PathType Leaf })
if ($selected.Count -eq 0) { throw 'No matching GraFiS stencils found in SourceDirectory.' }
$visio = $null
$documents = New-Object System.Collections.Generic.List[object]
$plans = New-Object System.Collections.Generic.List[object]
$completed = $false
try {
    # InvisibleApp creates an isolated automation instance; never attach to a user's active Visio.
    $visio = New-Object -ComObject Visio.InvisibleApp
    $visio.EventsEnabled = $false
    # Preflight all original stencils as read-only, hidden, with VBA macros disabled.
    foreach ($project in $selected) {
        $doc = $visio.Documents.OpenEx((Join-Path $source $project.stencil), [int16](2+64+128+256+8))
        $documents.Add($doc)
        if ($doc.VBProject.Protection -ne 0) { throw "VBA project is locked: $($project.stencil)" }
        $components = $doc.VBProject.VBComponents
        foreach ($module in $project.modules) {
            $component = $null
            foreach ($candidate in $components) {
                if ($candidate.Name -eq $module.name) { $component = $candidate; break }
            }
            if ($null -eq $component) {
                if (-not $module.is_new) { throw "Existing module missing: $($project.stencil) / $($module.name)" }
            } else {
                $cm = $component.CodeModule
                $currentCode = if ($cm.CountOfLines -gt 0) { $cm.Lines(1, $cm.CountOfLines) } else { '' }
                $currentHash = Get-CodeHash $currentCode
                if ($currentHash -ne $module.original_hash -and $currentHash -ne $module.new_hash -and $currentHash -notin $module.accepted_previous_hashes) {
                    throw "Source version differs: $($project.stencil) / $($module.name). No changes have been made."
                }
            }
            $code = Get-SourceCode (Join-Path $PackageDirectory $module.path) $module.encoding
            if ((Get-CodeHash $code) -ne $module.new_hash) { throw "Package module hash mismatch: $($module.path)" }
        }
        $doc.Close()
        $documents.Remove($doc) | Out-Null
        $plans.Add($project)
    }
    New-Item -ItemType Directory -Path $output | Out-Null
    foreach ($item in Get-ChildItem -LiteralPath $source -Force) {
        if ($item.Name -eq '.git' -or $item.Name -like '*.laccdb') { continue }
        Copy-Item -LiteralPath $item.FullName -Destination $output -Recurse
    }
    foreach ($project in $plans) {
        $target = Join-Path $output $project.stencil
        $doc = $visio.Documents.OpenEx($target, [int16](32+64+128+256+8))
        $documents.Add($doc)
        $components = $doc.VBProject.VBComponents
        foreach ($module in $project.modules) {
            $component = $null
            foreach ($candidate in $components) {
                if ($candidate.Name -eq $module.name) { $component = $candidate; break }
            }
            if ($null -eq $component) {
                $component = $components.Add([int]$module.component_type) # New standard modules/classes only; preserve existing designers.
                $component.Name = $module.name
            }
            $code = Get-SourceCode (Join-Path $PackageDirectory $module.path) $module.encoding
            $cm = $component.CodeModule
            if ($cm.CountOfLines -gt 0) { $cm.DeleteLines(1, $cm.CountOfLines) }
            $cm.AddFromString($code)
            if ((Get-CodeHash $cm.Lines(1, $cm.CountOfLines)) -ne $module.new_hash) {
                throw "Read-back verification failed: $($project.stencil) / $($module.name)"
            }
        }
        $doc.Save()
        $doc.Close()
        $documents.Remove($doc) | Out-Null
        Write-Host "Patched copy: $($project.stencil)"
    }
    $completed = $true
    Write-Host "Copy prepared: $output"
    Write-Host 'VBA has not been compiled or runtime-tested. Compile projects and run the checks in README before normal use.'
} finally {
    foreach ($doc in $documents.ToArray()) {
        try { $doc.Saved = $true; $doc.Close() } catch { }
    }
    if ($null -ne $visio) {
        try { $visio.Quit() } catch { }
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($visio)
    }
    if (-not $completed -and (Test-Path -LiteralPath $output)) {
        Write-Warning "Incomplete output copy: $output. Do not use it; original library was not modified."
    }
}
