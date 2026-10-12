Attribute VB_Name = "gfs_m_Model"
' Модуль для сборки модели пожара в формета yaml
Option Explicit


Public Function GetModelData(ByRef vsoPage As Visio.Page, ByVal g_ModelShapeFilter As String) As String
' Основаня процедура получения данных модели в YAML
Dim vsoShape As Visio.Shape
Dim yamlOutput As String
Dim s As String
Dim persIndex As indexPers

Dim walls As Collection
Dim nozzles As Collection
Dim hoses As Collection
Dim equipment As Collection
Dim gdzs As Collection
Dim watersupply As Collection
Dim management As Collection
Dim svaz As Collection
Dim ladders As Collection
Dim fireObjects As Collection
Dim labels As Collection
Dim otherTechs As Collection
Dim mainTechs As Collection
Dim sitePlan As Collection
Dim unclassified As Collection
Dim space As String

'    Set vsoPage = ActivePage
    yamlOutput = ""


    Set walls = New Collection
    Set nozzles = New Collection
    Set hoses = New Collection
    Set equipment = New Collection
    Set gdzs = New Collection
    Set watersupply = New Collection
    Set management = New Collection
    Set svaz = New Collection
    Set ladders = New Collection
    Set fireObjects = New Collection
    Set labels = New Collection
    Set otherTechs = New Collection
    Set mainTechs = New Collection
    Set sitePlan = New Collection
    Set unclassified = New Collection

    ' Заголовок документа
    yamlOutput = yamlOutput & "# Модель боевых действий по тушению пожара" & vbCrLf

    ' Метаданные
    yamlOutput = yamlOutput & "metadata:" & vbCrLf
'    yamlOutput = yamlOutput & "  version: """ & Application.ActiveDocument.Version & """" & vbCrLf
    s = vsoPage.Document.Description
    If s = "" Then s = "Ситуационная модель пожара"
    yamlOutput = yamlOutput & "  description: > " & vbCrLf
    yamlOutput = yamlOutput & "    """ & s & """" & vbCrLf
    yamlOutput = yamlOutput & "  author: """ & vsoPage.Document.Creator & """" & vbCrLf
    yamlOutput = yamlOutput & "  created_at: """ & Now() & """" & vbCrLf
    yamlOutput = yamlOutput & "  units:" & vbCrLf
    yamlOutput = yamlOutput & "    length: mm" & vbCrLf
    yamlOutput = yamlOutput & "    flow: lps" & vbCrLf
    yamlOutput = yamlOutput & "    pressure: mwc" & vbCrLf

    yamlOutput = yamlOutput & "  crs:" & vbCrLf
    yamlOutput = yamlOutput & "    type: local" & vbCrLf
    yamlOutput = yamlOutput & "    origin: [0, 0]" & vbCrLf
    yamlOutput = yamlOutput & "    axis_y: up" & vbCrLf
    yamlOutput = yamlOutput & "    north_deg: 90" & vbCrLf
    yamlOutput = yamlOutput & vbCrLf


    ' Общие данные об обстановке
    ' Будет реализовано позже

    ' Общие сведения о состоянии модели (странице)
    yamlOutput = yamlOutput & "state:" & vbCrLf
    yamlOutput = yamlOutput & "  state_name: " & vsoPage.name & vbCrLf
'    yamlOutput = yamlOutput & "  state_time: не указано" & vbCrLf  ' В ГраФиС не реализовано, но нужно иметь таткую заглушку
    space = "  "


    For Each vsoShape In vsoPage.Shapes
        If ModelShapeAllowed(vsoShape, g_ModelShapeFilter) And isWall(vsoShape) Then
            AddUniqueCollectionItem walls, vsoShape
        End If

        ' Проверяем, является ли фигура фигурой ГраФиС (фигурой УГО)
        If IsGFSShape(vsoShape) Then
            If ModelShapeAllowed(vsoShape, g_ModelShapeFilter) Then
                persIndex = cellVal(vsoShape, "User.IndexPers")
                Select Case GetObjectClass(persIndex)
                    Case Is = FireModelClass.fmcCommand
                        AddUniqueCollectionItem management, vsoShape
                    Case Is = FireModelClass.fmcComms
                        AddUniqueCollectionItem svaz, vsoShape
                    Case Is = FireModelClass.fmcFireEquipment
                        AddUniqueCollectionItem equipment, vsoShape
                    Case Is = FireModelClass.fmcFireObjects
                        AddUniqueCollectionItem fireObjects, vsoShape
                    Case Is = FireModelClass.fmcFireTruck
                        AddUniqueCollectionItem mainTechs, vsoShape
                    Case Is = FireModelClass.fmcGDZS
                        AddUniqueCollectionItem gdzs, vsoShape
                    Case Is = FireModelClass.fmcHoseLine
                        AddUniqueCollectionItem hoses, vsoShape
                    Case Is = FireModelClass.fmcLabels
                        AddUniqueCollectionItem labels, vsoShape
                    Case Is = FireModelClass.fmcLadders
                        AddUniqueCollectionItem ladders, vsoShape
                    Case Is = FireModelClass.fmcNozzle
                        AddUniqueCollectionItem nozzles, vsoShape
                    Case Is = FireModelClass.fmcOtherTech
                        AddUniqueCollectionItem otherTechs, vsoShape
                    Case Is = FireModelClass.fmcSitePlan
                        AddUniqueCollectionItem sitePlan, vsoShape
                    Case Is = FireModelClass.fmcUndefined
                        AddUniqueCollectionItem unclassified, vsoShape
                    Case Is = FireModelClass.fmcWaterSupply
                        AddUniqueCollectionItem watersupply, vsoShape
                End Select
            End If
        End If
    Next vsoShape

    ' Формирование блока ограждающих конструкций
    If walls.Count > 0 Then
        yamlOutput = yamlOutput & space & "walls:" & vbCrLf
        For Each vsoShape In walls
            yamlOutput = yamlOutput & GetWallYaml(vsoShape, space & "  ")
'            yamlOutput = yamlOutput & GetShapeYaml(vsoShape, space & "  ")
        Next vsoShape
    End If

    ' Формирование блоков СиС
    ' Пожар
    If fireObjects.Count > 0 Then
        yamlOutput = yamlOutput & space & "fire:" & vbCrLf
        For Each vsoShape In fireObjects
            yamlOutput = yamlOutput & GetShapeYaml(vsoShape, space & "  ")
        Next vsoShape
    End If
    ' Пожарная техника
    If mainTechs.Count > 0 Then
        yamlOutput = yamlOutput & space & "fire_units:" & vbCrLf
        For Each vsoShape In mainTechs
            yamlOutput = yamlOutput & GetShapeYaml(vsoShape, space & "  ")
        Next vsoShape
    End If
    ' Приборы подачи ОТВ
    If nozzles.Count > 0 Then
        yamlOutput = yamlOutput & space & "nozzles:" & vbCrLf
        For Each vsoShape In nozzles
            yamlOutput = yamlOutput & GetShapeYaml(vsoShape, space & "  ")
        Next vsoShape
    End If
    ' ПТВ
    If equipment.Count > 0 Then
        yamlOutput = yamlOutput & space & "equipment:" & vbCrLf
        For Each vsoShape In equipment
            yamlOutput = yamlOutput & GetShapeYaml(vsoShape, space & "  ")
        Next vsoShape
    End If
    ' ГДЗС
    If gdzs.Count > 0 Then
        yamlOutput = yamlOutput & space & "gdzs:" & vbCrLf
        For Each vsoShape In gdzs
            yamlOutput = yamlOutput & GetShapeYaml(vsoShape, space & "  ")
        Next vsoShape
    End If
    ' Управление
    If management.Count > 0 Then
        yamlOutput = yamlOutput & space & "management:" & vbCrLf
        For Each vsoShape In management
            yamlOutput = yamlOutput & GetShapeYaml(vsoShape, space & "  ")
        Next vsoShape
    End If
    ' Связь
    If svaz.Count > 0 Then
        yamlOutput = yamlOutput & space & "communications:" & vbCrLf
        For Each vsoShape In svaz
            yamlOutput = yamlOutput & GetShapeYaml(vsoShape, space & "  ")
        Next vsoShape
    End If
    ' Лестницы
    If ladders.Count > 0 Then
        yamlOutput = yamlOutput & space & "ladders:" & vbCrLf
        For Each vsoShape In ladders
            yamlOutput = yamlOutput & GetShapeYaml(vsoShape, space & "  ")
        Next vsoShape
    End If
    ' Техника прочее
    If otherTechs.Count > 0 Then
        yamlOutput = yamlOutput & space & "other_technic:" & vbCrLf
        For Each vsoShape In otherTechs
            yamlOutput = yamlOutput & GetShapeYaml(vsoShape, space & "  ")
        Next vsoShape
    End If
    ' Водоисточники
    If watersupply.Count > 0 Then
        yamlOutput = yamlOutput & space & "water_sources:" & vbCrLf
        For Each vsoShape In watersupply
            yamlOutput = yamlOutput & GetShapeYaml(vsoShape, space & "  ")
        Next vsoShape
    End If
    ' Рукавные линии
    If hoses.Count > 0 Then
        yamlOutput = yamlOutput & space & "hoses:" & vbCrLf
        For Each vsoShape In hoses
            yamlOutput = yamlOutput & GetShapeYaml(vsoShape, space & "  ")
        Next vsoShape
    End If


GetModelData = yamlOutput
End Function





' Applies a YAML model fragment to a Visio page.
Public Function ApplyModelYaml(ByVal yamlText As String, Optional ByVal pageName As String = "") As String
    Dim pg As Visio.Page, wasEventsEnabled As Boolean, i As Long
    Dim errNumber As Long, errDescription As String, errSource As String
    Dim response As String
    Dim undoScope As Long, scopeToClose As Long
    wasEventsEnabled = Application.EventsEnabled
    On Error GoTo Failed
    If Application.ActiveDocument Is Nothing Then Err.Raise vbObjectError + 2501, "ApplyModelYaml", "No active drawing document"
    If Application.ActiveDocument.Type <> visTypeDrawing Then Err.Raise vbObjectError + 2502, "ApplyModelYaml", "Active document is not a drawing"
    If Len(Trim$(pageName)) = 0 Then
        Set pg = Application.ActivePage
    Else
        Set pg = Application.ActiveDocument.Pages.Item(pageName)
    End If
    If pg Is Nothing Then Err.Raise vbObjectError + 2503, "ApplyModelYaml", "No target page"
    Application.EventsEnabled = False
    undoScope = Application.BeginUndoScope("GraFiS: apply model")
    For i = 0 To YamlElementCount(yamlText) - 1
        response = response & ApplyYamlElement(pg, YamlElementBlock(yamlText, i)) & vbCrLf
    Next i
    FixZIndex_P pg
    ApplyModelYaml = response
Finished:
    Application.EventsEnabled = False
    scopeToClose = undoScope
    undoScope = 0
    If scopeToClose <> 0 Then Application.EndUndoScope scopeToClose, (errNumber = 0)
    Application.EventsEnabled = wasEventsEnabled
    If errNumber <> 0 Then Err.Raise errNumber, errSource, errDescription
    Exit Function
Failed:
    errNumber = Err.Number
    errDescription = Err.Description
    errSource = Err.Source
    Resume Finished
End Function

' Returns the number of model elements recognised in a YAML fragment.
Public Function CountModelYamlElements(ByVal yamlText As String) As Long
'"   Возвращает число YAML-элементов, распознанных в переданной строке.
'"   Элементом считается строка формата - id: ....
'"   Используется в тестах для проверки, что парсер видит элементы, в том числе вложенные, например проёмы внутри стены.

    CountModelYamlElements = YamlElementCount(yamlText)
End Function

Public Function ApplyYamlElement(ByVal pg As Visio.Page, ByVal elementYaml As String) As String
'"   Получает отдельный YAML-блок элемента.
'"   Извлекает идентификатор id, имя мастера visio_master_name, тип геометрии и координаты.
'"   Ищет фигуру на странице по имени, совпадающему с id.
'"   Если фигура найдена, обновляет её свойства и геометрию.
'"   Если фигура не найдена, получает мастер через существующую функцию GetMaster, размещает его на странице и задаёт новой фигуре имя из id.

    Dim shapeId As String, masterName As String, geometryType As String
    Dim shp As Visio.Shape, mst As Visio.master, points As Collection
    Dim response As String
    shapeId = YamlScalar(Trim$(Mid$(Split(Replace(elementYaml, vbCrLf, vbLf), vbLf)(0), InStr(Split(Replace(elementYaml, vbCrLf, vbLf), vbLf)(0), "id:") + 3)))
'    shapeId = YamlValue(elementYaml, "id")
    If Len(shapeId) = 0 Then
        Err.Raise vbObjectError + 2504, "ApplyYamlElement", "Shape ID is missing"
    End If
    response = "ShapeId=" & shapeId & ": "
'    Set shp = FindShapeById(pg, "n" & shapeId)
    Set shp = FindShapeByName(pg, shapeId)
    geometryType = YamlNestedValue(elementYaml, "geometry", "type")
    Set points = YamlCoordinates(elementYaml)

    If shp Is Nothing Then
        masterName = YamlValue(elementYaml, "visio_master_name")
        Set mst = GetMaster(masterName)
        If mst Is Nothing Then
'            Err.Raise vbObjectError + 2101, , "Master not found: " & masterName
            Err.Raise vbObjectError + 2505, "ApplyYamlElement", "Master not found: " & masterName
        End If
        Set shp = DropMaster(pg, mst, points)
        On Error Resume Next
        shp.name = shapeId
        On Error GoTo 0
    End If

    ' Наполняем свойсвта и геометрию фигур
    ApplyYamlProperties shp, elementYaml
    ApplyYamlGeometry shp, geometryType, points, YamlNestedValue(elementYaml, "geometry", "heading_deg")

    ' Склеиваем геометрию отдельных фигур (стены, окна, двери)
    If isWall(shp) Or isWindow(shp) Or isDoor(shp) Then
        Dim moveSelection As Visio.Selection
        Dim previousEvents As Boolean
        previousEvents = Application.EventsEnabled
        On Error GoTo MoveFailed
        Set moveSelection = pg.CreateSelection(visSelTypeEmpty)
        moveSelection.Select shp, visSelect
        Application.EventsEnabled = True
        moveSelection.Move 0.01, 0#
        Application.EventsEnabled = previousEvents
        On Error GoTo 0
    End If

    ' Для рукавных линий нужно добаить приклеивание к фигурам
    Dim persIndex As indexPers
    persIndex = cellVal(shp, "User.IndexPers")
    If persIndex = indexPers.ipRukavLineNapor Or persIndex = indexPers.ipRukavLineVsas Then
        ' Собственно соединение
        ConnectHoseLine shp
        ' Пересчет длины
        SetCellVal shp, "User.LineLenight", shp.LengthIU / 39.37, visMeters
    End If


    ApplyYamlElement = response & "Data successfuly set."
    Exit Function
MoveFailed:
    Dim moveError As Long, moveDescription As String, moveSource As String
    moveError = Err.Number
    moveDescription = Err.Description
    moveSource = Err.Source
    Application.EventsEnabled = previousEvents
    Err.Raise moveError, moveSource, moveDescription
End Function


' Переделать!
Public Function DropMaster(ByVal pg As Visio.Page, ByVal mst As Visio.master, ByVal points As Collection) As Visio.Shape
'"   Размещает мастер mst на странице pg.
'"   В качестве позиции используется первая точка геометрии.
'"   Координаты YAML интерпретируются в миллиметрах и переводятся в внутренние единицы Visio - дюймы.

    Dim x As Double, y As Double, pt As Variant
    If points.Count > 0 Then
        pt = points(1): x = pt(0) / 25.4: y = pt(1) / 25.4
    End If
    Set DropMaster = pg.Drop(mst, x, y)
End Function

Public Sub ApplyYamlProperties(ByVal shp As Visio.Shape, ByVal elementYaml As String)
'"   Находит в YAML блок properties.
'"   Перебирает пары "имя свойства - значение" внутри этого блока.
'"   Заменяет символы подчёркивания в имени свойства на пробелы.
'"   Передаёт имя и значение в SetShapeProperty.

    Dim lines() As String, i As Long, propertyIndent As Long, indent As Long
    Dim name As String, value As String
    lines = Split(Replace(elementYaml, vbCrLf, vbLf), vbLf)
    For i = 0 To UBound(lines)
        If Trim$(lines(i)) = "properties:" Then
            propertyIndent = IndentOf(lines(i))
            i = i + 1
            Do While i <= UBound(lines)
                indent = IndentOf(lines(i))
                If Len(Trim$(lines(i))) > 0 And indent <= propertyIndent Then Exit Do
                If indent > propertyIndent And InStr(lines(i), ":") > 0 Then
                    name = Trim$(Left$(Trim$(lines(i)), InStr(Trim$(lines(i)), ":") - 1))
                    value = Trim$(Mid$(Trim$(lines(i)), InStr(Trim$(lines(i)), ":") + 1))
                    SetShapeProperty shp, Replace(name, "_", " "), YamlScalar(value)
                End If
                i = i + 1
            Loop
            Exit Sub
        End If
    Next i
End Sub

Public Sub SetShapeProperty(ByVal shp As Visio.Shape, ByVal label As String, ByVal value As String)
'"   Получает секцию пользовательских свойств visSectionProp фигуры.
'"   Если секции нет, создаёт её.
'"   Ищет строку свойства по отображаемому имени - полю Label.
'"   Если свойство отсутствует, добавляет новую строку свойства и задаёт ей Label.
'"   Для числовых, денежных, логических свойств и длительностей записывает значение как формулу.
'"   Для остальных типов записывает значение как текстовую строку Visio.

    Dim sec As Visio.Section, rw As Visio.row, rowIndex As Long, propType As Long, rowName As String
    On Error Resume Next
    Set sec = shp.Section(visSectionProp)
    On Error GoTo 0
    If sec Is Nothing Then shp.AddSection visSectionProp: Set sec = shp.Section(visSectionProp)
    For rowIndex = 0 To sec.Count - 1
        Set rw = sec.row(rowIndex)
        If StrComp(Trim$(rw.Cell(visCustPropsLabel).ResultStr(visUnitsString)), label, vbTextCompare) = 0 Then Exit For
    Next rowIndex
    If rowIndex = sec.Count Then
        rowIndex = shp.AddNamedRow(visSectionProp, "Prop" & CStr(sec.Count + 1), visTagDefault)
        shp.CellsSRC(visSectionProp, rowIndex, visCustPropsLabel).FormulaU = VisioString(label)
    End If
    rowName = shp.CellsSRC(visSectionProp, rowIndex, visCustPropsValue).name
    propType = shp.CellsSRC(visSectionProp, rowIndex, visCustPropsType).ResultIU
    If propType = visPropTypeNumber Or propType = visPropTypeCurrency Or propType = visPropTypeDuration Or propType = visPropTypeBool Then
'        shp.CellsSRC(visSectionProp, rowIndex, visCustPropsValue).FormulaU = Replace(value, ",", ".")
        SetCellVal shp, rowName, Replace(value, ",", ".")
    ElseIf propType = visPropTypeDate Then
'        shp.CellsSRC(visSectionProp, rowIndex, visCustPropsValue).FormulaU = """" & value & """"  ' VisioString(value)
        SetCellVal shp, rowName, value, visDate
    ElseIf propType = visPropTypeString Then
        SetCellVal shp, rowName, value, visUnitsString
    Else
        SetCellVal shp, rowName, value, visUnitsString
    End If
End Sub

Public Sub ApplyYamlGeometry(ByVal shp As Visio.Shape, ByVal geometryType As String, ByVal points As Collection, ByVal heading As String)
'"   Применяет геометрию к фигуре в соответствии с полем geometry.type.
'"   Для Point устанавливает PinX, PinY и, если он задан, угол Angle из heading_deg.
'"   Для LineString устанавливает координаты начала и конца через ячейки BeginX, BeginY, EndX и EndY.
'"   Для Line, Polygon, MultiPolygon и MultiLine передаёт набор точек в SetPolylineGeometry.

    Dim pt As Variant, x As Double, y As Double
    If points.Count = 0 Then Exit Sub
    pt = points(1): x = pt(0): y = pt(1)
'    Select Case LCase$(geometryType)
    Select Case geometryType
        Case "Point"
            SetCellVal shp, "PinX", CStr(x) & " mm"
            SetCellVal shp, "PinY", CStr(y) & " mm"
            If Len(heading) > 0 Then
                SetCellVal shp, "Angle", heading & " deg"
            End If
        Case "LineString"

            If points.Count < 2 Then Err.Raise vbObjectError + 2102, , "LineString requires two coordinates."

            ' Нужно добавить проверку на 1d behaviour - shp.OneD
            If IsOneDShape(shp) Then
                shp.CellsU("BeginX").FormulaU = Replace(CStr(x), ",", ".") & " mm"
                shp.CellsU("BeginY").FormulaU = Replace(CStr(y), ",", ".") & " mm"
                pt = points(2): x = pt(0): y = pt(1)
                shp.CellsU("EndX").FormulaU = Replace(CStr(x), ",", ".") & " mm"
                shp.CellsU("EndY").FormulaU = Replace(CStr(y), ",", ".") & " mm"
'            Else
'                SetPolylineGeometry shp, points, False
            End If

            ' Альтернативно (дли оттличия стен и рукавных линий и вообще фигур с поведением 1d с геометрией и без)
            If shp.SectionExists(visSectionFirstComponent, False) Then
                SetPolylineGeometry shp, points
            End If

        Case Is = "MultiLine"
            SetPolylineGeometry shp, points
        Case "Polygon"
            SetPolygonGeometry shp, points
        Case Is = "MultiPolygon"
            SetPolygonGeometry shp, points
    End Select
End Sub

Public Sub SetPolylineGeometry(ByVal shp As Visio.Shape, ByVal points As Collection)
    Dim pointCount As Long
    Dim i As Long
    Dim pointValue As Variant
    Dim firstPoint As Variant
    Dim lastPoint As Variant
    Dim x As Double
    Dim y As Double
    Dim minX As Double
    Dim maxX As Double
    Dim minY As Double
    Dim maxY As Double
    Dim widthValue As Double
    Dim heightValue As Double
    Dim scopeId As Long
    Dim savedErrorNumber As Long
    Dim savedErrorDescription As String

    Dim newX As Double
    Dim newY As Double

    If shp Is Nothing Then
        Err.Raise vbObjectError + 1000, "SetPolylineGeometry", "The shape argument is Nothing."
    End If
    If points Is Nothing Then
        Err.Raise vbObjectError + 1001, "SetPolylineGeometry", "The points collection is Nothing."
    End If

    pointCount = points.Count
    If pointCount < 2 Then
        Err.Raise vbObjectError + 1002, "SetPolylineGeometry", _
                  "The points collection must contain at least two points."
    End If

    On Error GoTo InvalidPoint
    firstPoint = points(1)
    lastPoint = points(points.Count)

    minX = CDbl(firstPoint(0))
    maxX = CDbl(lastPoint(0))
    minY = CDbl(firstPoint(1))
    maxY = CDbl(lastPoint(1))

    SetCellVal shp, "BeginX", minX, VALUE_MODE
    SetCellVal shp, "EndX", maxX, VALUE_MODE
    SetCellVal shp, "BeginY", minY, VALUE_MODE
    SetCellVal shp, "EndY", maxY, VALUE_MODE


'    For i = 2 To pointCount
'        pointValue = points(i)
'        x = CDbl(pointValue(0))
'        y = CDbl(pointValue(1))
'        If x < minX Then minX = x
'        If x > maxX Then maxX = x
'        If y < minY Then minY = y
'        If y > maxY Then maxY = y
'    Next i
    On Error GoTo 0

'    widthValue = maxX - minX
    widthValue = cellVal(shp, "Width", visMillimeters)
    heightValue = maxY - minY
'    heightValue = cellVal(shp, "Height", visMillimeters)
'    If widthValue <= 0# Or heightValue <= 0# Then
'        Err.Raise vbObjectError + 1003, "SetPolylineGeometry", _
'                  "The points must define a non-zero bounding rectangle."
'    End If

    scopeId = Application.BeginUndoScope("Replace shape geometry with absolute polyline")
    On Error GoTo OperationFailed

    ' Make the shape coordinate system parallel to the page axes.
'    shp.CellsU("Angle").ResultIU = 0#
'    shp.CellsU("Width").FormulaU = Replace(CStr(widthValue), ",", ".") & " mm"
    shp.CellsU("Height").FormulaU = Replace(CStr(heightValue), ",", ".") & " mm"
'    shp.CellsU("LocPinX").FormulaU = Replace(CStr(widthValue / 2#), ",", ".") & " mm"
'    shp.CellsU("LocPinY").FormulaU = Replace(CStr(heightValue / 2#), ",", ".") & " mm"
'    shp.CellsU("PinX").FormulaU = Replace(CStr((minX + maxX) / 2#), ",", ".") & " mm"
'    shp.CellsU("PinY").FormulaU = Replace(CStr((minY + maxY) / 2#), ",", ".") & " mm"

    ' Remove all existing geometry rows except the initial MoveTo row.
    ClearGeometryRows shp

    firstPoint = points(1)
    shp.rowType(visSectionFirstComponent, 1) = visTagMoveTo
'    shp.CellsSRC(visSectionFirstComponent, 1, visX).FormulaU = _
'        Replace(CStr(CDbl(firstPoint(0)) - minX), ",", ".") & " mm"
'    shp.CellsSRC(visSectionFirstComponent, 1, visY).FormulaU = _
'        Replace(CStr(CDbl(firstPoint(1)) - minY), ",", ".") & " mm"
'    shp.CellsSRC(visSectionFirstComponent, 1, visX).FormulaU = _
'        "=Width*" & Replace(CStr((CDbl(firstPoint(0)) - minX) / widthValue), ",", ".")
'    shp.CellsSRC(visSectionFirstComponent, 1, visY).FormulaU = _
'        "=Height*" & Replace(CStr((CDbl(firstPoint(1)) - minY) / heightValue), ",", ".")
    shp.XYFromPage MMToInches(firstPoint(0)), MMToInches(firstPoint(1)), newX, newY
    newX = InchesToMM(newX)
    newY = InchesToMM(newY)
    shp.CellsSRC(visSectionFirstComponent, 1, visX).FormulaU = _
        "=Width*" & Replace(CStr(CDbl(newX) / widthValue), ",", ".")
    shp.CellsSRC(visSectionFirstComponent, 1, visY).FormulaU = _
        "=Height*" & Replace(CStr(CDbl(newY) / heightValue), ",", ".")

    For i = 2 To pointCount
        pointValue = points(i)
        shp.AddRow visSectionFirstComponent, visRowLast, visTagLineTo
'        shp.CellsSRC(visSectionFirstComponent, visRowLast, visX).FormulaU = _
'            Replace(CStr(CDbl(pointValue(0)) - minX), ",", ".") & " mm"
'        shp.CellsSRC(visSectionFirstComponent, visRowLast, visY).FormulaU = _
'            Replace(CStr(CDbl(pointValue(1)) - minY), ",", ".") & " mm"
'        shp.CellsSRC(visSectionFirstComponent, visRowLast, visX).FormulaU = _
'            "=Width*" & Replace(CStr((CDbl(pointValue(0)) - minX) / widthValue), ",", ".")
'        shp.CellsSRC(visSectionFirstComponent, visRowLast, visY).FormulaU = _
'            "=Height*" & Replace(CStr((CDbl(pointValue(1)) - minY) / heightValue), ",", ".")
        shp.XYFromPage MMToInches(pointValue(0)), MMToInches(pointValue(1)), newX, newY
        newX = InchesToMM(newX)
        newY = InchesToMM(newY)
        shp.CellsSRC(visSectionFirstComponent, visRowLast, visX).FormulaU = _
            "=Width*" & Replace(CStr(CDbl(newX) / widthValue), ",", ".")
        shp.CellsSRC(visSectionFirstComponent, visRowLast, visY).FormulaU = _
            "=Height*" & Replace(CStr(CDbl(newY) / heightValue), ",", ".")


    Next i

'    If closePath And pointCount > 2 Then
'        shp.AddRow visSectionFirstComponent, visRowLast, visTagLineTo
'        shp.CellsSRC(visSectionFirstComponent, visRowLast, visX).FormulaU = _
'            Replace(CStr(CDbl(firstPoint(0)) - minX), ",", ".") & " mm"
'        shp.CellsSRC(visSectionFirstComponent, visRowLast, visY).FormulaU = _
'            Replace(CStr(CDbl(firstPoint(1)) - minY), ",", ".") & " mm"
'    End If

    Application.EndUndoScope scopeId, True
    Exit Sub

InvalidPoint:
    Err.Raise vbObjectError + 1004, "SetPolylineGeometry", _
              "Each collection item must contain numeric x and y values at indexes 0 and 1."

OperationFailed:
    savedErrorNumber = Err.Number
    savedErrorDescription = Err.Description
    Application.EndUndoScope scopeId, False
    Err.Raise savedErrorNumber, "SetPolylineGeometry", savedErrorDescription
End Sub

Public Sub SetPolygonGeometry(ByVal shp As Visio.Shape, ByVal points As Collection)
    Dim pointCount As Long
    Dim i As Long
    Dim pointValue As Variant
    Dim firstPoint As Variant
    Dim x As Double
    Dim y As Double
    Dim minX As Double
    Dim maxX As Double
    Dim minY As Double
    Dim maxY As Double
    Dim widthValue As Double
    Dim heightValue As Double
    Dim scopeId As Long
    Dim savedErrorNumber As Long
    Dim savedErrorDescription As String

    Dim newX As Double
    Dim newY As Double

    If shp Is Nothing Then
        Err.Raise vbObjectError + 1000, "SetPolylineGeometry", "The shape argument is Nothing."
    End If
    If points Is Nothing Then
        Err.Raise vbObjectError + 1001, "SetPolylineGeometry", "The points collection is Nothing."
    End If

    pointCount = points.Count
    If pointCount < 2 Then
        Err.Raise vbObjectError + 1002, "SetPolylineGeometry", _
                  "The points collection must contain at least two points."
    End If

    On Error GoTo InvalidPoint
    firstPoint = points(1)
    minX = CDbl(firstPoint(0))
    maxX = minX
    minY = CDbl(firstPoint(1))
    maxY = minY

    For i = 2 To pointCount
        pointValue = points(i)
        x = CDbl(pointValue(0))
        y = CDbl(pointValue(1))
        If x < minX Then minX = x
        If x > maxX Then maxX = x
        If y < minY Then minY = y
        If y > maxY Then maxY = y
    Next i
    On Error GoTo 0

    widthValue = maxX - minX
    heightValue = maxY - minY
    If widthValue <= 0# Or heightValue <= 0# Then
        Err.Raise vbObjectError + 1003, "SetPolylineGeometry", _
                  "The points must define a non-zero bounding rectangle."
    End If

    scopeId = Application.BeginUndoScope("Replace shape geometry with absolute polyline")
    On Error GoTo OperationFailed

    ' Make the shape coordinate system parallel to the page axes.
    shp.CellsU("Angle").ResultIU = 0#
    shp.CellsU("Width").FormulaU = Replace(CStr(widthValue), ",", ".") & " mm"
    shp.CellsU("Height").FormulaU = Replace(CStr(heightValue), ",", ".") & " mm"
    shp.CellsU("LocPinX").FormulaU = Replace(CStr(widthValue / 2#), ",", ".") & " mm"
    shp.CellsU("LocPinY").FormulaU = Replace(CStr(heightValue / 2#), ",", ".") & " mm"
    shp.CellsU("PinX").FormulaU = Replace(CStr((minX + maxX) / 2#), ",", ".") & " mm"
    shp.CellsU("PinY").FormulaU = Replace(CStr((minY + maxY) / 2#), ",", ".") & " mm"

    ' Remove all existing geometry rows except the initial MoveTo row.
    ClearGeometryRows shp

    firstPoint = points(1)
    shp.rowType(visSectionFirstComponent, 1) = visTagMoveTo
'    shp.CellsSRC(visSectionFirstComponent, 1, visX).FormulaU = _
'        Replace(CStr(CDbl(firstPoint(0)) - minX), ",", ".") & " mm"
'    shp.CellsSRC(visSectionFirstComponent, 1, visY).FormulaU = _
'        Replace(CStr(CDbl(firstPoint(1)) - minY), ",", ".") & " mm"
    shp.XYFromPage MMToInches(firstPoint(0)), MMToInches(firstPoint(1)), newX, newY
    newX = InchesToMM(newX)
    newY = InchesToMM(newY)
    shp.CellsSRC(visSectionFirstComponent, 1, visX).FormulaU = _
        "=Width*" & Replace(CStr(CDbl(newX) / widthValue), ",", ".")
    shp.CellsSRC(visSectionFirstComponent, 1, visY).FormulaU = _
        "=Height*" & Replace(CStr(CDbl(newY) / heightValue), ",", ".")

    For i = 2 To pointCount
        pointValue = points(i)
        shp.AddRow visSectionFirstComponent, visRowLast, visTagLineTo
'        shp.CellsSRC(visSectionFirstComponent, visRowLast, visX).FormulaU = _
'            Replace(CStr(CDbl(pointValue(0)) - minX), ",", ".") & " mm"
'        shp.CellsSRC(visSectionFirstComponent, visRowLast, visY).FormulaU = _
'            Replace(CStr(CDbl(pointValue(1)) - minY), ",", ".") & " mm"
        shp.XYFromPage MMToInches(pointValue(0)), MMToInches(pointValue(1)), newX, newY
        newX = InchesToMM(newX)
        newY = InchesToMM(newY)
        shp.CellsSRC(visSectionFirstComponent, visRowLast, visX).FormulaU = _
            "=Width*" & Replace(CStr(CDbl(newX) / widthValue), ",", ".")
        shp.CellsSRC(visSectionFirstComponent, visRowLast, visY).FormulaU = _
            "=Height*" & Replace(CStr(CDbl(newY) / heightValue), ",", ".")
    Next i

'    If closePath And pointCount > 2 Then
        shp.AddRow visSectionFirstComponent, visRowLast, visTagLineTo
'        shp.CellsSRC(visSectionFirstComponent, visRowLast, visX).FormulaU = _
'            Replace(CStr(CDbl(firstPoint(0)) - minX), ",", ".") & " mm"
'        shp.CellsSRC(visSectionFirstComponent, visRowLast, visY).FormulaU = _
'            Replace(CStr(CDbl(firstPoint(1)) - minY), ",", ".") & " mm"
        shp.XYFromPage MMToInches(firstPoint(0)), MMToInches(firstPoint(1)), newX, newY
        newX = InchesToMM(newX)
        newY = InchesToMM(newY)
        shp.CellsSRC(visSectionFirstComponent, visRowLast, visX).FormulaU = _
            "=Width*" & Replace(CStr(CDbl(newX) / widthValue), ",", ".")
        shp.CellsSRC(visSectionFirstComponent, visRowLast, visY).FormulaU = _
            "=Height*" & Replace(CStr(CDbl(newY) / heightValue), ",", ".")

'    End If

    Application.EndUndoScope scopeId, True
    Exit Sub

InvalidPoint:
    Err.Raise vbObjectError + 1004, "SetPolylineGeometry", _
              "Each collection item must contain numeric x and y values at indexes 0 and 1."

OperationFailed:
    savedErrorNumber = Err.Number
    savedErrorDescription = Err.Description
    Application.EndUndoScope scopeId, False
    Err.Raise savedErrorNumber, "SetPolylineGeometry", savedErrorDescription
End Sub

Public Sub ClearGeometryRows(ByVal shp As Visio.Shape)
    Const ROW_INDEX As Long = 2

    If shp Is Nothing Then
        Err.Raise vbObjectError + 1005, "ClearGeometryRows", "The shape argument is Nothing."
    End If

    ' Remove rows one by one at the specified index.
    If shp.SectionExists(visSectionFirstComponent, False) Then
        Do While shp.RowExists(visSectionFirstComponent, ROW_INDEX, False)
            shp.DeleteRow visSectionFirstComponent, ROW_INDEX
        Loop
    End If
End Sub






Public Function YamlElementCount(ByVal yamlText As String) As Long
'"   Подсчитывает строки YAML, начинающиеся с - id:.
'"   Используется основной процедурой для определения количества элементов.

    Dim lines() As String, i As Long
    lines = Split(Replace(yamlText, vbCrLf, vbLf), vbLf)
    For i = 0 To UBound(lines)
        If LCase$(Trim$(lines(i))) Like "- id:*" Then YamlElementCount = YamlElementCount + 1
    Next i
End Function

Public Function YamlElementBlock(ByVal yamlText As String, ByVal elementIndex As Long) As String
'"   Возвращает YAML-текст элемента по его порядковому номеру.
'"   Начинает формирование блока со строки - id:.
'"   Использует отступы YAML для определения границ блока.

    Dim lines() As String, i As Long, found As Long, baseIndent As Long
    lines = Split(Replace(yamlText, vbCrLf, vbLf), vbLf)
    For i = 0 To UBound(lines)
        If LCase$(Trim$(lines(i))) Like "- id:*" Then
            If found = elementIndex Then
                baseIndent = IndentOf(lines(i))
                Do
                    YamlElementBlock = YamlElementBlock & lines(i) & vbLf: i = i + 1
                    If i > UBound(lines) Then Exit Do
                    If LCase$(Trim$(lines(i))) Like "- id:*" And IndentOf(lines(i)) <= baseIndent Then Exit Do
                Loop
                Exit Function
            End If
            found = found + 1
        End If
    Next i
End Function

Public Function YamlValue(ByVal elementYaml As String, ByVal key As String) As String
'"   Извлекает простое значение верхнего уровня из блока элемента.
'"   Используется, например, для получения id и visio_master_name.

    YamlValue = YamlNestedValue(elementYaml, "", key)
End Function

Public Function YamlNestedValue(ByVal elementYaml As String, ByVal parentKey As String, ByVal key As String) As String
'"   Ищет значение поля внутри указанного вложенного блока YAML.
'"   Например, получает type и heading_deg из блока geometry.

    Dim lines() As String, i As Long, parentIndent As Long, inParent As Boolean, t As String
    lines = Split(Replace(elementYaml, vbCrLf, vbLf), vbLf)
    For i = 0 To UBound(lines)
        t = Trim$(lines(i))
        If Len(parentKey) > 0 And t = parentKey & ":" Then parentIndent = IndentOf(lines(i)): inParent = True
        If inParent And Len(t) > 0 And IndentOf(lines(i)) <= parentIndent And t <> parentKey & ":" Then Exit Function
        If (Len(parentKey) = 0 Or inParent) And LCase$(Left$(t, Len(key) + 1)) = LCase$(key & ":") Then
            YamlNestedValue = YamlScalar(Trim$(Mid$(t, Len(key) + 2))): Exit Function
        End If
    Next i
End Function

Public Function YamlCoordinates(ByVal elementYaml As String) As Collection
'"   Находит строку или многострочный блок coordinates.
'"   С помощью регулярного выражения извлекает пары числовых координат [x, y].
'"   Поддерживает дробные числа с точкой либо запятой.
'"   Возвращает координаты как коллекцию пар чисел.

    Dim lines() As String, i As Long, j As Long, coordIndent As Long, text As String, re As Object, matches As Object, m As Object
    Set YamlCoordinates = New Collection
    lines = Split(Replace(elementYaml, vbCrLf, vbLf), vbLf)
    For i = 0 To UBound(lines)
        If LCase$(Trim$(lines(i))) Like "coordinates:*" Then
            coordIndent = IndentOf(lines(i)): text = lines(i)
            For j = i + 1 To UBound(lines)
                If Len(Trim$(lines(j))) > 0 And IndentOf(lines(j)) <= coordIndent Then Exit For
                text = text & " " & lines(j)
            Next j
            Exit For
        End If
    Next i
    Set re = CreateObject("VBScript.RegExp"): re.Global = True: re.Pattern = "\[\s*(-?[0-9]+(?:[\.,][0-9]+)?)\s*,\s*(-?[0-9]+(?:[\.,][0-9]+)?)"
    Set matches = re.Execute(text)
    For Each m In matches
        YamlCoordinates.Add Array(CDbl(Replace(m.SubMatches(0), ".", ",")), CDbl(Replace(m.SubMatches(1), ".", ",")))
    Next m
End Function

Public Function YamlScalar(ByVal value As String) As String
'"   Удаляет внешние одинарные или двойные кавычки у простого YAML-значения.
'"   Используется для получения текстового значения без YAML-оформления.

    value = Trim$(value)
    If Len(value) > 1 Then
        If (Left$(value, 1) = """" And Right$(value, 1) = """") Or (Left$(value, 1) = "'" And Right$(value, 1) = "'") Then value = Mid$(value, 2, Len(value) - 2)
    End If
    YamlScalar = Replace(value, "\""", """")
End Function

Public Function IndentOf(ByVal text As String) As Long
'"   Возвращает число пробелов в начале строки.
'"   Нужна парсеру для определения вложенности YAML-блоков.

    IndentOf = Len(text) - Len(LTrim$(text))
End Function

Public Function VisioString(ByVal value As String) As String
'"   Формирует строковую формулу Visio в двойных кавычках.
'"   Двойные кавычки внутри значения удваиваются, чтобы они не нарушали синтаксис формулы.

    VisioString = """" & Replace(value, """", """") & """"
End Function



Public Function ModelShapeAllowed(ByVal shp As Visio.Shape, ByVal g_ModelShapeFilter As String) As Boolean
    Dim names() As String
    Dim i As Long
    Dim item As String

    If Len(Trim$(g_ModelShapeFilter)) = 0 Then
        ModelShapeAllowed = True
        Exit Function
    End If

    names = Split(g_ModelShapeFilter, ";")
    For i = LBound(names) To UBound(names)
        item = Trim$(names(i))
        If Len(item) > 0 Then
            If StrComp(item, shp.name, vbTextCompare) = 0 Then
                ModelShapeAllowed = True
                Exit Function
            End If

            On Error Resume Next
            If StrComp(item, shp.ID, vbTextCompare) = 0 Then
                ModelShapeAllowed = True
                On Error GoTo 0
                Exit Function
            End If
            On Error GoTo 0
        End If
    Next i
End Function
