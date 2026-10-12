Attribute VB_Name = "m_tools"
Option Explicit




Public Function PFB_isDoor(ByRef aO_Shape As Visio.Shape) As Boolean
'Функция возвращает Истина, если фигура - дверной проем, в противном случае - Ложь
    
'---Проверяем, является ли фигура фигурой конструкций
    If aO_Shape.CellExists("User.ShapeClass", 0) = False Or aO_Shape.CellExists("User.ShapeType", 0) = False Then
        PFB_isDoor = False
        Exit Function
    End If

'---Проверяем, является ли фигура фигурой ДВЕРЬ или ПРОЕМ
    If aO_Shape.Cells("User.ShapeClass").Result(visNumber) = 3 And _
        (aO_Shape.Cells("User.ShapeType").Result(visNumber) = 10 Or aO_Shape.Cells("User.ShapeType").Result(visNumber) = 25) Then
        PFB_isDoor = True
        Exit Function
    End If
PFB_isDoor = False
End Function

Public Function PFB_isPlace(ByRef aO_Shape As Visio.Shape) As Boolean
'Функция возвращает Истина, если фигура - дверной проем, в противном случае - Ложь
    
'---Проверяем, является ли фигура фигурой конструкций
    If aO_Shape.CellExists("User.ShapeClass", 0) = False Or aO_Shape.CellExists("User.ShapeType", 0) = False Then
        PFB_isPlace = False
        Exit Function
    End If

'---Проверяем, является ли фигура фигурой МЕСТО
    If aO_Shape.Cells("User.ShapeClass").Result(visNumber) = 5 And _
        aO_Shape.Cells("User.ShapeType").Result(visNumber) = 38 Then
        PFB_isPlace = True
        Exit Function
    End If
PFB_isPlace = False
End Function

Public Function PFB_isWall(ByRef aO_Shape As Visio.Shape) As Boolean
'Функция возвращает Истина, если фигура - стена, в противном случае - Ложь
Dim shapeType As Integer
'---Проверяем, является ли фигура фигурой конструкций
    If aO_Shape.CellExists("User.ShapeClass", 0) = False Or aO_Shape.CellExists("User.ShapeType", 0) = False Then
        PFB_isWall = False
        Exit Function
    End If

'---Проверяем, является ли фигура фигурой СТЕНА
    shapeType = aO_Shape.Cells("User.ShapeType").Result(visNumber)
    If aO_Shape.Cells("User.ShapeClass").Result(visNumber) = 3 And _
        (shapeType = 44 Or shapeType = 6) Then
        PFB_isWall = True
        Exit Function
    End If
PFB_isWall = False
End Function



Public Function GetNeighbors(ByRef shp As Visio.Shape) As Collection
Dim sel As Visio.Selection
Dim shpChild As Visio.Shape
Dim col As Collection
    
    Set sel = shp.SpatialNeighbors(visSpatialTouching + visSpatialOverlap, 0, 0)
    
    If sel.count > 0 Then
        Set col = New Collection
        
        For Each shpChild In sel
            If PFB_isDoor(shpChild) Then
                Debug.Print shpChild
            ElseIf PFB_isPlace(shpChild) Then
                Debug.Print shpChild
            End If
        Next shpChild
    Else
        Set GetNeighbors = Nothing
    End If
    
    

    
End Function


Public Function Interpolate(ByVal x As Single, ByVal x0 As Single, ByVal x1 As Single, ByVal y0 As Single, ByVal y1 As Single) As Single
    Interpolate = ((x - x0) / (x1 - x0)) * (y1 - y0) + y0
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
