Attribute VB_Name = "GFS_StencilSessionTests"
Option Explicit
' Import this test and c_GFSStencilSession.cls into a disposable blank drawing.
' The test creates its own temporary stencil file and two unsaved drawings.
Public Sub RunStencilSessionTests()
    Dim sourceStencil As Visio.Document, runtimeStencil As Visio.Document
    Dim drawingA As Visio.Document, drawingB As Visio.Document
    Dim editableStencil As Visio.Document
    Dim session As c_GFSStencilSession
    Dim editableSession As c_GFSStencilSession
    Dim wasEventsEnabled As Boolean, tempFile As String
    Dim fso As Object, errNumber As Long, errDescription As String
    wasEventsEnabled = Application.EventsEnabled
    On Error GoTo Failed
    Set fso = CreateObject("Scripting.FileSystemObject")
    tempFile = fso.BuildPath(fso.GetSpecialFolder(2), fso.GetTempName & ".vss")
    Set sourceStencil = Application.Documents.AddEx("vss", visMSDefault, visAddHidden)
    sourceStencil.SaveAs tempFile
    sourceStencil.Close
    Set sourceStencil = Nothing
    Set drawingA = Application.Documents.Add("")
    drawingA.Pages(1).DrawRectangle 0, 0, 1, 1
    drawingA.Saved = False
    Set runtimeStencil = Application.Documents.OpenEx(tempFile, visOpenRO)
    Application.EventsEnabled = True
    DrawingWindow(drawingA).Activate
    Set session = New c_GFSStencilSession
    session.Start runtimeStencil
    Require DockedWith(DrawingWindow(drawingA), runtimeStencil), "Stencil must be docked in drawing A"
    Require NoSeparateStencilWindow(runtimeStencil), "Runtime stencil must not have a visible separate window"
    Require Not drawingA.Saved, "Unsaved drawing A must not be marked saved"
    runtimeStencil.Saved = False
    session.MarkRuntimeClean
    Require runtimeStencil.Saved, "Read-only runtime stencil must not request saving"
    Require Not drawingA.Saved, "Cleaning library state must leave A unsaved"
    Set drawingB = Application.Documents.Add("")
    DrawingWindow(drawingB).Activate
    Require DockedWith(DrawingWindow(drawingB), runtimeStencil), "Same stencil must be docked in drawing B"
    session.Stop
    Set session = Nothing
    runtimeStencil.Close
    Set runtimeStencil = Nothing
    Set editableStencil = Application.Documents.OpenEx(tempFile, visOpenRW + visOpenHidden)
    Set editableSession = New c_GFSStencilSession
    editableSession.Start editableStencil
    editableStencil.Saved = False
    editableSession.MarkRuntimeClean
    Require Not editableStencil.Saved, "Intentional edits to a writable stencil must remain unsaved"
    Debug.Print "PASS: docking in two drawings; runtime-only prompts; drawing and writable stencil edits preserved"
Finished:
    On Error Resume Next
    If Not session Is Nothing Then session.Stop
    If Not editableSession Is Nothing Then editableSession.Stop
    If Not editableStencil Is Nothing Then
        editableStencil.Saved = True
        editableStencil.Close
    End If
    If Not runtimeStencil Is Nothing Then
        runtimeStencil.Saved = True
        runtimeStencil.Close
    End If
    If Not sourceStencil Is Nothing Then
        sourceStencil.Saved = True
        sourceStencil.Close
    End If
    If Not drawingB Is Nothing Then
        drawingB.Saved = True
        drawingB.Close
    End If
    If Not drawingA Is Nothing Then
        drawingA.Saved = True
        drawingA.Close
    End If
    If Len(tempFile) > 0 Then
        If Not fso Is Nothing Then
            If fso.FileExists(tempFile) Then fso.DeleteFile tempFile
        End If
    End If
    Application.EventsEnabled = wasEventsEnabled
    On Error GoTo 0
    If errNumber <> 0 Then Err.Raise errNumber, "RunStencilSessionTests", errDescription
    Exit Sub
Failed:
    errNumber = Err.Number
    errDescription = Err.Description
    Resume Finished
End Sub

Private Function DrawingWindow(ByVal drawing As Visio.Document) As Visio.Window
    Dim w As Visio.Window
    For Each w In Application.Windows
        If w.Type = visDrawing Then
            If w.Document Is drawing Then
                Set DrawingWindow = w
                Exit Function
            End If
        End If
    Next w
    Err.Raise vbObjectError + 2700, , "Drawing window not found"
End Function

Private Function DockedWith(ByVal w As Visio.Window, ByVal stencil As Visio.Document) As Boolean
    Dim names As Variant, i As Long, doc As Visio.Document
    On Error GoTo NotDocked
    w.DockedStencils names
    For i = LBound(names) To UBound(names)
        Set doc = Application.Documents.Item(CStr(names(i)))
        If doc Is stencil Then
            DockedWith = True
            Exit Function
        End If
    Next i
NotDocked:
End Function

Private Function NoSeparateStencilWindow(ByVal stencil As Visio.Document) As Boolean
    Dim w As Visio.Window
    NoSeparateStencilWindow = True
    For Each w In Application.Windows
        If w.Type = 2 Then
            If w.Document Is stencil Then
                If w.Visible Then
                    NoSeparateStencilWindow = False
                    Exit Function
                End If
            End If
        End If
    Next w
End Function

Private Sub Require(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then Err.Raise vbObjectError + 2701, , message
End Sub
