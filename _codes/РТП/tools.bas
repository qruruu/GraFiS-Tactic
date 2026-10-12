Attribute VB_Name = "tools"


Public Function GetScaleAt200() As Double
'Возвращает коэффициент приведения размера текущей страницы относительно масштаба 1:200
Dim v_Minor As Double
Dim v_Major As Double

    v_Minor = Application.ActivePage.PageSheet.Cells("PageScale").Result(visNumber)
    v_Major = Application.ActivePage.PageSheet.Cells("DrawingScale").Result(visNumber)
    GetScaleAt200 = (v_Major / v_Minor) / 200
End Function


'--------------------------------Сохранение лога ошибки-------------------------------------
Public Sub SaveLog(ByRef error As ErrObject, ByVal eroorPosition As String, Optional ByVal addition As String)
    Dim logHandle As Integer
    Dim logOpened As Boolean
    Dim logPath As String, errString As String
    Dim savedNumber As Long, savedDescription As String, savedSource As String
    savedNumber = error.Number
    savedDescription = error.Description
    savedSource = error.Source
    On Error GoTo LogFailed
    errString = Now & " | " & Environ$("OS") & " | Visio " & Application.Version & _
        " | " & ThisDocument.FullName & " | " & eroorPosition & " | " & savedNumber & _
        " | " & savedDescription & " | " & savedSource & " | " & addition
    logPath = ThisDocument.Path
    If Len(logPath) > 0 Then
        If Right$(logPath, 1) <> "\" Then logPath = logPath & "\"
    Else
        logPath = Environ$("TEMP") & "\"
    End If
    logHandle = FreeFile
    Open logPath & "Log.txt" For Append As #logHandle
    logOpened = True
    Print #logHandle, errString
LogFinished:
    On Error Resume Next
    If logOpened Then Close #logHandle
    Exit Sub
LogFailed:
    Debug.Print "GraFiS log unavailable: " & errString
    Resume LogFinished
End Sub

