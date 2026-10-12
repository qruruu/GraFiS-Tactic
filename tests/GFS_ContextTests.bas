Attribute VB_Name = "GFS_ContextTests"
Option Explicit

' Import this module and m_GFSSafety.bas into a disposable blank drawing.
' Run RunContextTests. It creates and closes only its own two unsaved drawings.
Public Sub RunContextTests()
    Dim docA As Visio.Document, docB As Visio.Document
    Dim shapeA As Visio.Shape, shapeB As Visio.Shape, resolved As Visio.Shape
    Dim keyA As String, keyB As String, keyOther As String
    Dim wasEventsEnabled As Boolean
    Dim testError As Long, testDescription As String
    Dim timeA As Date, timeB As Date
    wasEventsEnabled = Application.EventsEnabled
    On Error GoTo Failed
    Application.EventsEnabled = False
    Set docA = Application.Documents.Add("")
    Set shapeA = docA.Pages(1).DrawRectangle(0, 0, 1, 1)
    Set docB = Application.Documents.Add("")
    Set shapeB = docB.Pages(1).DrawRectangle(0, 0, 2, 2)
    AssertTrue shapeA.ID = shapeB.ID, "Fixture must contain equal IDs in two documents"
    Set resolved = GFSResolveShape(shapeA.ID, shapeA)
    AssertTrue (resolved Is shapeA), "Event object context must win over active document B"
    Set resolved = GFSResolveShape(shapeB.ID, shapeB)
    AssertTrue (resolved Is shapeB), "Explicit B context must resolve B"
    resolved.Cells("Width").FormulaU = "3 in"
    AssertTrue Abs(shapeA.Cells("Width").ResultIU - 1#) < 0.000001, "Updating B must leave A unchanged"
    timeA = GFSCurrentTime(shapeA)
    timeB = GFSCurrentTime(shapeB)
    AssertTrue docA.DocumentSheet.CellExistsU("User.CurrentTime", 0) <> 0, "Time cells must exist in A"
    AssertTrue docB.DocumentSheet.CellExistsU("User.CurrentTime", 0) <> 0, "Second document B must get its own time cells"
    docA.DocumentSheet.CellsU("User.CurrentTime").FormulaU = "DATETIME(46000)"
    AssertTrue GFSCurrentTime(shapeB) = timeB, "Setting A time must leave B unchanged"
    keyA = GFSBeginCellEvent(shapeA.Cells("Width"))
    AssertTrue Len(keyA) > 0, "First A event must enter"
    AssertTrue Len(GFSBeginCellEvent(shapeA.Cells("Width"))) = 0, "Recursive same-cell event must be blocked"
    keyOther = GFSBeginCellEvent(shapeA.Cells("Height"))
    AssertTrue Len(keyOther) > 0, "Dependent event for a different cell must still enter"
    keyB = GFSBeginCellEvent(shapeB.Cells("Width"))
    AssertTrue Len(keyB) > 0, "Same shape ID in another document must have a separate guard"
    GFSEndCellEvent keyA
    keyA = GFSBeginCellEvent(shapeA.Cells("Width"))
    AssertTrue Len(keyA) > 0, "Cleanup must permit the next event"
    Debug.Print "PASS: source objects, duplicate IDs, document isolation, per-cell recursion and cleanup"
Finished:
    On Error Resume Next
    GFSEndCellEvent keyA
    GFSEndCellEvent keyB
    GFSEndCellEvent keyOther
    If Not docB Is Nothing Then
        docB.Saved = True
        docB.Close
    End If
    If Not docA Is Nothing Then
        docA.Saved = True
        docA.Close
    End If
    Application.EventsEnabled = wasEventsEnabled
    On Error GoTo 0
    If testError <> 0 Then Err.Raise testError, "RunContextTests", testDescription
    Exit Sub
Failed:
    testError = Err.Number
    testDescription = Err.Description
    Resume Finished
End Sub

Private Sub AssertTrue(ByVal condition As Boolean, ByVal message As String)
    If Not condition Then Err.Raise vbObjectError + 2600, "GFS_ContextTests", message
End Sub
