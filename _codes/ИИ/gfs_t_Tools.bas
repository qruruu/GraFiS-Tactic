Attribute VB_Name = "gfs_t_Tools"
'-----------------------------------------Модуль инструментальных функций----------------------------------------------
Option Explicit


Public Function IsModelShape(ByRef shp As Visio.Shape) As Boolean
' Функция проверяет, является ли фигура модельной
    If isWall(shp) Then
        IsModelShape = True
        Exit Function
    End If
    If isDoor(shp) Then
        IsModelShape = True
        Exit Function
    End If
    If IsGFSShape(shp) Then
        IsModelShape = True
        Exit Function
    End If
IsModelShape = False
End Function


Public Function isWall(ByRef aO_Shape As Visio.Shape) As Boolean
'Функция возвращает Истина, если фигура - стена, в противном случае - Ложь
Dim shapeType As Integer
'---Проверяем, является ли фигура фигурой конструкций
    If aO_Shape.CellExists("User.ShapeClass", 0) = False Or aO_Shape.CellExists("User.ShapeType", 0) = False Then
        isWall = False
        Exit Function
    End If

'---Проверяем, является ли фигура фигурой СТЕНА
    shapeType = aO_Shape.Cells("User.ShapeType").result(visNumber)
    If aO_Shape.Cells("User.ShapeClass").result(visNumber) = 3 And _
        (shapeType = 44 Or shapeType = 6) Then
        isWall = True
        Exit Function
    End If
isWall = False
End Function

Public Function isDoor(ByRef aO_Shape As Visio.Shape) As Boolean
'Функция возвращает Истина, если фигура - дверной проем, в противном случае - Ложь
Dim shapeType As Integer
'---Проверяем, является ли фигура фигурой конструкций
    If aO_Shape.CellExists("User.ShapeClass", 0) = False Or aO_Shape.CellExists("User.ShapeType", 0) = False Then
        isDoor = False
        Exit Function
    End If

'---Проверяем, является ли фигура фигурой ДВЕРЬ или ПРОЕМ
    shapeType = aO_Shape.Cells("User.ShapeType").result(visNumber)
    If aO_Shape.Cells("User.ShapeClass").result(visNumber) = 3 And _
        (shapeType = 10 Or shapeType = 25) Then
        isDoor = True
        Exit Function
    End If
isDoor = False
End Function

Public Function isWindow(ByRef aO_Shape As Visio.Shape) As Boolean
'Функция возвращает Истина, если фигура - окно, в противном случае - Ложь
Dim shapeType As Integer
'---Проверяем, является ли фигура фигурой конструкций
    If aO_Shape.CellExists("User.ShapeClass", 0) = False Or aO_Shape.CellExists("User.ShapeType", 0) = False Then
        isWindow = False
        Exit Function
    End If

'---Проверяем, является ли фигура фигурой ОКНО
    shapeType = aO_Shape.Cells("User.ShapeType").result(visNumber)
    If aO_Shape.Cells("User.ShapeClass").result(visNumber) = 3 And _
        (shapeType = 45) Then
        isWindow = True
        Exit Function
    End If
isWindow = False
End Function

Public Function IsGFSShape(ByRef shp As Visio.Shape, Optional ByVal useManeure As Boolean = True) As Boolean
'Функция возвращает True, если фигура является фигурой ГраФиС
Dim i As Integer

    'Проверяем, является ли фигура фигурой ГраФиС
    If useManeure Then      'Если нужно учитывать проверку на маневр
        If shp.CellExists("User.IndexPers", 0) = True Then
            'Если имеется ячейка опции Маневра и ее значение показывает, что
            If shp.CellExists("Actions.MainManeure", 0) = True Then
                If shp.Cells("Actions.MainManeure.Checked").result(visNumber) = 0 Then
                    IsGFSShape = True       'Фигура ГраФиС и не маневренная
                Else
                    IsGFSShape = False      'Фигура ГраФиС и маневренная
                End If
            Else
                IsGFSShape = True       'Фигура ГраФиС и не имеет ячейки Маневр
            End If
        Else
            IsGFSShape = False      'Фигура не ГраФиС
        End If
    Else                    'если не нужно учитывать проверку на маневр
        IsGFSShape = shp.CellExists("User.IndexPers", 0)
    End If

End Function

' Функция проверки, является ли фигура точечной
Public Function IsPointGeometry(ByVal shp As Visio.Shape) As Boolean
    Dim cellCount As Integer
    Dim hasGeometry As Boolean

    ' Точечные фигуры обычно:
    ' 1. Имеют минимальную геометрию
    ' 2. Являются маркерами или символами
    ' 3. Не имеют сложной формы

    ' Проверяем наличие геометрии
    On Error Resume Next
    hasGeometry = shp.SectionExists(visSectionFirstComponent, False)
    On Error GoTo 0

    ' Если нет секции геометрии или это простой маркер
    If Not hasGeometry Then
        IsPointGeometry = True
        Exit Function
    End If

    ' Проверяем количество вершин в первой строке геометрии
    Dim vertexCount As Integer
    On Error Resume Next
    vertexCount = shp.rowCount(visSectionFirstComponent)   ', visRowFirst)
    On Error GoTo 0

    ' Если мало вершин (1-2), считаем точечной
    If vertexCount <= 2 Then
        IsPointGeometry = True
    Else
        IsPointGeometry = False
    End If
End Function

Public Function FormatNumberWithDot(ByVal value As Double, Optional ByVal decimals As Integer = 0) As String
    Dim formatted As String

    ' Форматируем число стандартным способом
    If decimals = 0 Then
        formatted = Format(value, "0")
    Else
        formatted = Format(value, "0." & String(decimals, "0"))
    End If

    ' Заменяем системный разделитель (запятую) на точку
    FormatNumberWithDot = Replace(formatted, ",", ".")
End Function

'Public Function YSL(ByVal level As Byte) As String
'Dim i As Byte
'Dim result As String
'
'    result = ""
'    For i = 1 To level
'        result = result & YS  ' "--"  'YS
'    Next i
'YSL = result
'End Function



'=============================================================================
' Дюймы в Миллиметры
'=============================================================================
Public Function InchesToMM(ByVal inches As Double) As Double
    InchesToMM = inches * INCH_TO_MM
End Function

'=============================================================================
' Миллиметры в Дюймы (для записи значений обратно в Visio)
'=============================================================================
Public Function MMToInches(ByVal mm As Double) As Double
    MMToInches = mm * MM_TO_INCH
End Function


'=============================================================================
' Проверка того является ли строка числом и в этом случае перевод в строку с точкой
'=============================================================================
Public Function StrNumCheck(s As String, Optional ByVal decimals As Integer = 1) As String
Dim ss As String

    On Error GoTo ex
    ss = CDbl(s)
    StrNumCheck = FormatNumberWithDot(ss, decimals)
    Exit Function
ex:
    StrNumCheck = s
End Function



' Оборачивает значение в двойные кавычки YAML.
' Внутренние обратные слеши и двойные кавычки экранируются.
Public Function QuoteYamlString(ByVal value As String) As String
Dim escapedValue As String

    ' Возможно это не нужно ----------------------------------------
    escapedValue = Replace(value, Chr(92), Chr(92) & Chr(92))
    escapedValue = Replace(escapedValue, Chr(34), Chr(92) & Chr(34))
    ' --------------------------------------------------------------
    QuoteYamlString = Chr(34) & escapedValue & Chr(34)

End Function



Public Function IsOneDShape(ByVal shp As Visio.Shape) As Boolean
    On Error GoTo NotOneD
    IsOneDShape = shp.OneD
    If Not IsOneDShape Then
        IsOneDShape = shp.CellExists("BeginX", 0) And shp.CellExists("BeginY", 0) And shp.CellExists("EndX", 0) And shp.CellExists("EndY", 0)
    End If
    Exit Function
NotOneD:
    IsOneDShape = False
End Function


Public Sub FixZIndex_P(Optional ByVal pg As Visio.Page = Nothing)
'Прока исправляет положение фигур ГраФиС относительно стен и прочего
Dim vsoSelection As Visio.Selection

    On Error GoTo ex
    If pg Is Nothing Then Set pg = Application.ActivePage
    If pg Is Nothing Then Exit Sub

    '---перемещаем вперед //Техника;ПТВ;Рукавные линии;Водоисточники;Очаг;Статисты и посредники
    Set vsoSelection = pg.CreateSelection(visSelTypeByLayer, visSelModeSkipSuper, "Техника;ПТВ;Рукавные линии;Водоисточники;Очаг;Статисты и посредники")

    vsoSelection.BringToFront

    '---перемещаем вперед //ГДЗС;Подписи рукавов
    Set vsoSelection = pg.CreateSelection(visSelTypeByLayer, visSelModeSkipSuper, "ГДЗС;Подписи рукавов;Очаг;Управление СиС")

    vsoSelection.BringToFront

    '---перемещаем назад //Задымление
    Set vsoSelection = pg.CreateSelection(visSelTypeByLayer, visSelModeSkipSuper, "Задымление")

    vsoSelection.SendToBack

Exit Sub
ex:

End Sub
