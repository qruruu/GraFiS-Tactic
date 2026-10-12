Attribute VB_Name = "m_tools"
Option Explicit



Public Function GetReadyString(ByVal val As Variant, ByVal prefix As String, ByVal postfix As String, Optional ignore As Variant = 0, Optional ifEmpty As String = "") As String
'Возвращает сформированную строку в виде prefix & val & postfix. В случае если val=ignore, возвращает ifEmpty
'В случае невозможности преобразования val к строке возвращает ifEmpty
    
    On Error GoTo ex
    
    If val = ignore Then
        GetReadyString = ifEmpty
    Else
        GetReadyString = prefix & str(val) & postfix
    End If
Exit Function
ex:
    GetReadyString = ifEmpty
End Function
Public Function GetReadyStringA(ByVal elemID As String, ByVal prefix As String, ByVal postfix As String, Optional ignore As Variant = 0, Optional ifEmpty As String = "") As String
'Возвращает сформированную строку в виде prefix & val & postfix. В случае если val=ignore, возвращает ifEmpty
'В случае невозможности преобразования val к строке возвращает ifEmpty
'Самостоятельно образщается к A, где elemID - код данных получаемых анализатором моделей
Dim val As Variant
    
    On Error GoTo ex
    
    val = A.Result(elemID)
    If val = ignore Then
        GetReadyStringA = ifEmpty
    Else
        GetReadyStringA = prefix & str(val) & postfix
    End If
Exit Function
ex:
    GetReadyStringA = ifEmpty
End Function

Public Sub fixAllGFSShapesC()
'Замена английской C на русскую С
Dim shp As Visio.Shape
    
    For Each shp In A.gfsShapes
        SetCellVal shp, "Prop.Unit", Replace(cellVal(shp, "Prop.Unit", visUnitsString), "C", "С")
    Next shp
End Sub

Public Function ClearString(ByVal txt As String) As String
'Функция получает на вход строку и возвращает округленное до сотых число (если строка представляла число) или исходную строку
Dim tmpVal As Variant
    On Error Resume Next
    txt = Round(txt, 2)
    ClearString = txt
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

