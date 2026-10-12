VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_GFSStencilSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents host As Visio.Application
Private owner As Visio.Document
Private preparedDrawings As Object
Private busy As Boolean
Private manageRuntime As Boolean
Private initiallySaved As Boolean
Private cleaning As Boolean
Public Event DrawingReady(ByVal drawingWindow As Visio.Window)

Public Sub Start(ByVal stencil As Visio.Document)
    If stencil Is Nothing Then Exit Sub
    If stencil.Type <> visTypeStencil Then Exit Sub
    Set owner = stencil
    initiallySaved = stencil.Saved
    manageRuntime = stencil.ReadOnly
    Set preparedDrawings = CreateObject("Scripting.Dictionary")
    Set host = stencil.Application
    TryDrawing host.ActiveWindow
End Sub

Public Sub Stop()
    Set host = Nothing
    Set owner = Nothing
    Set preparedDrawings = Nothing
End Sub

' Read-only library files are runtime resources. Editable stencils and drawings
' are never marked saved here; this does not write or overwrite any file.
Public Sub MarkRuntimeClean()
    On Error GoTo Failed
    If cleaning Then Exit Sub
    If owner Is Nothing Then Exit Sub
    If Not manageRuntime Then Exit Sub
    If Not initiallySaved Then Exit Sub
    If owner.Type <> visTypeStencil Then Exit Sub
    If Not owner.ReadOnly Then Exit Sub
    cleaning = True
    owner.Saved = True
    cleaning = False
    Exit Sub
Failed:
    cleaning = False
    Debug.Print "GraFiS stencil cleanup: " & Err.Description
End Sub

Private Sub host_WindowActivated(ByVal Window As IVWindow)
    TryDrawing Window
End Sub

Private Sub host_BeforeDocumentClose(ByVal doc As IVDocument)
    On Error Resume Next
    If preparedDrawings Is Nothing Then Exit Sub
    If doc.Type = visTypeDrawing Then
        If preparedDrawings.Exists(CStr(doc.ID)) Then preparedDrawings.Remove CStr(doc.ID)
    End If
    If Not owner Is Nothing Then
        If doc Is owner Then MarkRuntimeClean
    End If
End Sub

Private Sub host_DocumentChanged(ByVal doc As IVDocument)
    If owner Is Nothing Then Exit Sub
    If doc Is owner Then MarkRuntimeClean
End Sub

Private Sub host_CellChanged(ByVal Cell As IVCell)
    On Error GoTo Failed
    If owner Is Nothing Then Exit Sub
    If Cell.Shape Is Nothing Then Exit Sub
    If Cell.Shape.Document Is owner Then MarkRuntimeClean
    Exit Sub
Failed:
    Debug.Print "GraFiS stencil cell event: " & Err.Description
End Sub

Private Sub TryDrawing(ByVal drawingWindow As Visio.Window)
    Dim drawing As Visio.Document, drawingKey As String
    Dim entered As Boolean
    If busy Then Exit Sub
    If owner Is Nothing Then Exit Sub
    If drawingWindow Is Nothing Then Exit Sub
    On Error GoTo Failed
    If drawingWindow.Type <> visDrawing Then Exit Sub
    Set drawing = drawingWindow.Document
    If drawing Is Nothing Then Exit Sub
    If drawing.Type <> visTypeDrawing Then Exit Sub
    If drawingWindow.Page Is Nothing Then Exit Sub
    busy = True
    drawingWindow.Activate
    If manageRuntime Then
        DockStencil drawingWindow, owner
        DockVisibleReadOnlyStencils drawingWindow
        EnsureLibrary drawingWindow
    End If
    drawingWindow.Activate
    drawingKey = CStr(drawing.ID)
    If Not drawing.ReadOnly Then
        If Not preparedDrawings.Exists(drawingKey) Then
            preparedDrawings.Add drawingKey, True
            entered = True
            RaiseEvent DrawingReady(drawingWindow)
        End If
    End If
Finished:
    MarkRuntimeClean
    busy = False
    Exit Sub
Failed:
    Debug.Print "GraFiS stencil session: " & Err.Number & " " & Err.Description
    If entered Then
        If preparedDrawings.Exists(drawingKey) Then preparedDrawings.Remove drawingKey
    End If
    Resume Finished
End Sub

Private Function IsDocked(ByVal drawingWindow As Visio.Window, ByVal stencil As Visio.Document) As Boolean
    Dim names As Variant, i As Long, dockedDocument As Visio.Document
    On Error GoTo NotDocked
    drawingWindow.DockedStencils names
    For i = LBound(names) To UBound(names)
        Set dockedDocument = host.Documents.Item(CStr(names(i)))
        If StrComp(dockedDocument.FullName, stencil.FullName, vbTextCompare) = 0 Then
            IsDocked = True
            Exit Function
        End If
    Next i
NotDocked:
End Function

Private Sub DockStencil(ByVal drawingWindow As Visio.Window, ByVal stencil As Visio.Document)
    Dim opened As Visio.Document, topWindow As Visio.Window
    On Error GoTo Failed
    If stencil Is Nothing Then Exit Sub
    If stencil.Type <> visTypeStencil Then Exit Sub
    If Not stencil.ReadOnly Then Exit Sub
    If Len(stencil.FullName) = 0 Then Exit Sub
    If Not IsDocked(drawingWindow, stencil) Then
        drawingWindow.Activate
        Set opened = host.Documents.OpenEx(stencil.FullName, visOpenRO + visOpenDocked + visOpenNoWorkspace + visOpenDontList)
    End If
    ' Hide only an MDI stencil window after its docked copy was verified.
    ' Window type 2 is the built-in MDI stencil window; docked panes are type 7.
    If IsDocked(drawingWindow, stencil) Then
        For Each topWindow In host.Windows
            If topWindow.Type = 2 Then
                If Not topWindow.Document Is Nothing Then
                    If topWindow.Document Is stencil Then topWindow.Visible = False
                End If
            End If
        Next topWindow
    End If
    Exit Sub
Failed:
    Debug.Print "GraFiS docking: " & stencil.Name & " " & Err.Description
End Sub

Private Sub DockVisibleReadOnlyStencils(ByVal drawingWindow As Visio.Window)
    Dim topWindow As Visio.Window, stencil As Visio.Document
    Dim candidates As Collection, seen As Object
    Set candidates = New Collection
    Set seen = CreateObject("Scripting.Dictionary")
    On Error GoTo Failed
    ' Take a snapshot before OpenEx changes the window collection.
    For Each topWindow In host.Windows
        If topWindow.Type = 2 Then
            If topWindow.Visible Then
                Set stencil = topWindow.Document
                If Not stencil Is Nothing Then
                    If stencil.Type = visTypeStencil Then
                        If stencil.ReadOnly Then
                            If Not seen.Exists(CStr(stencil.ID)) Then
                                seen.Add CStr(stencil.ID), True
                                candidates.Add stencil
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next topWindow
    For Each stencil In candidates
        DockStencil drawingWindow, stencil
    Next stencil
    Exit Sub
Failed:
    Debug.Print "GraFiS stencil window scan: " & Err.Description
End Sub

Private Function LibraryDocumentLoaded(ByVal fileName As String) As Boolean
    Dim doc As Visio.Document
    For Each doc In host.Documents
        If StrComp(doc.FullName, fileName, vbTextCompare) = 0 Then
            LibraryDocumentLoaded = True
            Exit Function
        End If
    Next doc
End Function

Private Sub EnsureLibrary(ByVal drawingWindow As Visio.Window)
    Dim files As Variant, item As Variant, fileName As String
    Dim central As String, opened As Visio.Document
    If Not manageRuntime Then Exit Sub
    central = owner.Path & "Моделирование.vss"
    If StrComp(owner.Name, "Моделирование.vss", vbTextCompare) <> 0 Then
        If Len(Dir$(central)) > 0 Then
            If Not LibraryDocumentLoaded(central) Then
                drawingWindow.Activate
                Set opened = host.Documents.OpenEx(central, visOpenRO + visOpenDocked + visOpenNoWorkspace + visOpenDontList)
                drawingWindow.Activate
            End If
        End If
        Exit Sub
    End If
    files = Split("Авиация.vss|Автотехника.vss|Водоснабжение НРС.vss|Водоснабжение.vss|ГДЗС.vss|Железная дорога.vss|ИИ.vss|Конструкции.vss|" & _
        "Линии НРС.vss|Линии.vss|Моделирование.vss|Отчеты.vss|Очаг.vss|ПТВ НРС.vss|ПТВ.vss|План на местности.vss|" & _
        "Пожарная техника НРС.vss|Пожарная техника.vss|Поэтажный план.vss|Прочее.vss|Резервуары.vss|Робототехнические средства.vss|Свойства и экспорт.vss|Связь и освещение.vss|" & _
        "Совмещенный график.vss|Специальные функции.vss|Техника прочее НРС.vss|Техника прочее.vss|Управление СиС.vss|Учения.vss|Формулы.vss|Эвакуация.vss", "|")
    For Each item In files
        fileName = owner.Path & CStr(item)
        If Len(Dir$(fileName)) > 0 Then
            If Not LibraryDocumentLoaded(fileName) Then
                drawingWindow.Activate
                Set opened = host.Documents.OpenEx(fileName, visOpenRO + visOpenDocked + visOpenNoWorkspace + visOpenDontList)
                drawingWindow.Activate
            End If
        End If
    Next item
End Sub
