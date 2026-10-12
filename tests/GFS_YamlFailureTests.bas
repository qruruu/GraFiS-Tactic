Attribute VB_Name = "GFS_YamlFailureTests"
Option Explicit
' Import into a COPY of the AI stencil. Select a disposable drawing before running.
Public Sub RunYamlFailureTests()
    Dim pg As Visio.Page
    Dim originalCount As Long, errorNumber As Long
    Dim savedEvents As Boolean, state As Integer
    Dim result As String
    Set pg = Application.ActivePage
    If pg Is Nothing Then Err.Raise vbObjectError + 2601, , "Select a disposable drawing first"
    originalCount = pg.Shapes.Count
    savedEvents = Application.EventsEnabled
    On Error GoTo Failed
    For state = 0 To 1
        Application.EventsEnabled = CBool(state)
        On Error Resume Next
        result = ApplyModelYaml("- id: regression_missing_master" & vbCrLf & _
            "  visio_master_name: __GFS_REGRESSION_NONEXISTENT_MASTER__" & vbCrLf)
        errorNumber = Err.Number
        Err.Clear
        On Error GoTo Failed
        If errorNumber = 0 Then Err.Raise vbObjectError + 2602, , "Missing master must report failure"
        If Application.EventsEnabled <> CBool(state) Then Err.Raise vbObjectError + 2603, , "Event state was not restored"
        If pg.Shapes.Count <> originalCount Then Err.Raise vbObjectError + 2604, , "Failed model left added shapes"
    Next state
    Application.EventsEnabled = savedEvents
    Debug.Print "PASS: missing master reports failure and restores both original event states"
    Exit Sub
Failed:
    Dim savedError As Long, savedDescription As String
    savedError = Err.Number
    savedDescription = Err.Description
    Application.EventsEnabled = savedEvents
    Err.Raise savedError, "RunYamlFailureTests", savedDescription
End Sub
