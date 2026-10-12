VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Dim ButEvent As Class1
Dim C_ConnectionsTrace As c_HoseConnector
Dim WithEvents LineAppEvents As Visio.Application
Attribute LineAppEvents.VB_VarHelpID = -1


Private WithEvents gfsSession As c_GFSStencilSession
Private gfsUIReady As Boolean

Private Sub GFSCloseCore(ByVal doc As IVDocument)
    On Error Resume Next
    Set LineAppEvents = Nothing
    Set C_ConnectionsTrace = Nothing
'Процедура закрытия документа и удаления его рабочих элементов

'---Очищаем объект ButEvent и удаляем кнопку "Рукав" с панели управления "Превращения"
    Set ButEvent = Nothing
    DeleteButtonLine
    DeleteButtonMLine
    DeleteButtonVHose
    DeleteButtonNormalize
    DeleteButtonNRSSettings
    DeleteButtonNRSReport
    
'---В случае, если на панели "Превращения" нет ни одной кнопки, удаляем её
    If Application.CommandBars("Превращения").Controls.Count = 0 Then RemoveTBImagination
'---В случае, если на панели "Насосно-рукавные системы" нет ни одной кнопки, удаляем её
    If Application.CommandBars("Насосно-рукавные системы").Controls.Count = 0 Then RemoveTBNRSCalc
    
'---Очищаем переменную приложения
Set LineAppEvents = Nothing
Set C_ConnectionsTrace = Nothing
End Sub


Private Sub GFSOpenedCore(ByVal doc As IVDocument)
    If Not GFSHasEditableDrawing() Then Exit Sub

'    On Error GoTo ex
'---Показываем окно свойств
    Application.ActiveWindow.Windows.ItemFromID(visWinIDCustProp).Visible = True

'---Добавляем ячейки "User.FireTime", "User.CurrentTime"
    AddTimeUserCells

'---Объявляем переменную приложения для дальнейшего реагирования на изменение содержимого ячеек
    Set LineAppEvents = Visio.Application
'---Активируем экземпляр класса для отслеживания соединений
    Set C_ConnectionsTrace = New c_HoseConnector

'---Импортируем мастера
    sp_MastersImport

'---Добавляем ячейку Аспект (если еще не была добавлена)
    If Not Application.ActivePage.PageSheet.CellExists("User.GFS_Aspect", 0) Then
        Application.ActivePage.PageSheet.AddNamedRow visSectionUser, "GFS_Aspect", 0
        Application.ActivePage.PageSheet.Cells("User.GFS_Aspect").FormulaU = 1
    End If

'---ОБновляем/экспортируем в активный документ стили трафарета
    '---Проверяем не является ли активный документ документом цветовой схемы
    If Application.ActiveDocument.DocumentSheet.CellExists("User.GFSColorTheme", 0) = 0 Then
        StyleExport
    End If

'---Создаем панель управления "Превращения" и добавляем на нее кнопку "Обратить в линию"
    AddTBImagination        'Добавляем тулбокс Обрашщения
    AddButtonLine           'Добавляем кнопку рабочей линии
    AddButtonMLine          'Добавляем фигуру магистральной линии
    AddButtonVHose          'Добавляем кнопку всасывающей линии
'---Создаем панель управления "Насосно-рукавные системы"
    AddTBNRSCalc
    AddButtonNormalize      'Добавляем кнопку "Расчет НРС"
    AddButtonNRSSettings    'Добавляем кнопку "Настройки расчета НРС"
    AddButtonNRSReport      'Добавляем кнопку "Отчет расчета НРС"

Set ButEvent = New Class1

'---Активируем опцию приклеивания к контурам фигур
    Application.ActiveDocument.GlueSettings = visGlueToGeometry + visGlueToGuides + visGlueToConnectionPoints + visGlueToVertices

'---Добавляем для документа своство "FireTime"
    sm_AddFireTime
    
'---Проверяем наличие обновлений
'    fmsgCheckNewVersion.CheckUpdates
    
Exit Sub
ex:
    SaveLog Err, "Document_DocumentOpened"
End Sub


Private Sub LineAppEvents_CellChangedCore(ByVal cell As IVCell)
'Процедура обновления списков в фигурах
Dim ShpInd As Long
'---Проверяем имя ячейки
    
    If cell.Name = "Prop.HoseMaterial" Or cell.Name = "Prop.HoseDiameter" Then
        ShpInd = cell.Shape.ID
        '---Запускаем процедуру получения СПИСКОВ диаметров рукавов
        HoseDiametersListImport ShpInd, cell.Shape
        '---Запускаем процедуру получения ЗНАЧЕНИЙ Сопротивлений рукавов
        HoseResistanceValueImport ShpInd, cell.Shape
        '---Запускаем процедуру получения ЗНАЧЕНИЙ Пропускной способности рукавов
        HoseMaxFlowValueImport ShpInd, cell.Shape
        '---Запускаем процедуру получения ЗНАЧЕНИЙ Массы рукавов
        HoseWeightValueImport ShpInd, cell.Shape
    End If
    
    
    
'MsgBox Cell.Shape.Index
'MsgBox Cell.Shape.ID

'В случае, если произошло изменение не нужной ячейки прекращаем событие
End Sub

Private Sub sp_TemplatesImport()
    Visio.Application.ActivePage.Drop ThisDocument.Masters("ВсасывающийРукав"), 0, 0
End Sub


Private Sub sp_MastersImport()
    If Not GFSHasEditableDrawing() Then Exit Sub
'---Импортируем мастера

    MasterImportSub "Линии НРС.vss", "ВсасывающийРукав"
    MasterImportSub "Линии НРС.vss", "ВсасывающийРукав1000"
    MasterImportSub "Линии НРС.vss", "НапорноВсасывающийРукав"
    
End Sub


Private Sub sm_AddFireTime()
    If Not GFSHasEditableDrawing() Then Exit Sub
'Процедура добавляет в документ свойство - время начала пожара, в случае его отсутствия

    If Application.ActiveDocument.DocumentSheet.CellExists("User.FireTime", 0) = False Then
        Application.ActiveDocument.DocumentSheet.AddNamedRow visSectionUser, "FireTime", 0
        Application.ActiveDocument.DocumentSheet.Cells("User.FireTime").FormulaU = "Now()"
    End If

End Sub

Private Sub LineAppEvents_ShapeAdded(ByVal Shape As IVShape)
'Событие добавления на лист фигуры
Dim v_Ctrl As CommandBarControl
Dim SecExists As Boolean
    If Not GFSIsActiveDrawingShape(Shape) Then Exit Sub
    
'---Включаем обработку ошибок
    On Error GoTo Tail

'---Проверяем какая кнопка нажата и в зависимости от этого выполняем действие
    For Each v_Ctrl In Application.CommandBars("Превращения").Controls
        If v_Ctrl.State = msoButtonDown Then
            Select Case v_Ctrl.Caption
                Case Is = "Рукав"
                    '---Запускаем процедуру обращения в рабочую рукавную линию
                    If IsSelectedOneShape(False) Then
                    '---Если выбрана хоть одна фигура - пытаемся ее обратить
                        If Not IsHavingUserSection(False) And Not IsSquare(False) Then
                        '---Обращаем фигуру в фигуру рабочей рукавной линии
'                            ButEvent.MakeHoseLine
                            MakeHoseLine 51, 0
                        End If
                    End If
                Case Is = "Всасывающий рукав"
                    '---Запускаем процедуру обращения во всасывающую рукавную линию
                    If IsSelectedOneShape(False) Then
                    '---Если выбрана хоть одна фигура - пытаемся ее обратить
                        If Not IsHavingUserSection(False) And Not IsSquare(False) Then
                        '---Обращаем фигуру в фигуру всасывающей рукавной линии
'                            ButEvent.MakeVHoseLine
                            MakeVHoseLine
                        End If
                    End If
                Case Is = "Магистральная линия"
                    '---Запускаем процедуру обращения в магистральную рукавную линию
                    If IsSelectedOneShape(False) Then
                    '---Если выбрана хоть одна фигура - пытаемся ее обратить
                        If Not IsHavingUserSection(False) And Not IsSquare(False) Then
                        '---Обращаем фигуру в фигуру магистральной рукавной линии
'                            ButEvent.MakeMagHoseLine
                            MakeHoseLine 77, 1
                        End If
                    End If
            End Select
        End If
    Next v_Ctrl
    
Exit Sub
Tail:
    MsgBox "В ходе выполнения программы произошла ошибка! Если она будет повторяться - обратитесь к разработчику."
    SaveLog Err, "Document_DocumentOpened"
End Sub

Private Sub AddTimeUserCells()
'Прока добавляет ячейки "User.FireTime", "User.CurrentTime"
Dim docSheet As Visio.Shape
Dim cell As Visio.cell
    If Not GFSHasEditableDrawing() Then Exit Sub

    Set docSheet = Application.ActiveDocument.DocumentSheet
    
    If Not docSheet.CellExists("User.FireTime", 0) Then
        docSheet.AddNamedRow visSectionUser, "FireTime", visTagDefault
        docSheet.Cells("User.FireTime").FormulaU = "Now()"
    End If
    If Not docSheet.CellExists("User.CurrentTime", 0) Then
        docSheet.AddNamedRow visSectionUser, "CurrentTime", visTagDefault
        docSheet.Cells("User.CurrentTime").FormulaU = "User.FireTime"
    End If

End Sub

Private Sub LineAppEvents_CellChanged(ByVal cell As IVCell)
    Dim eventKey As String
    On Error GoTo EventFailed
    eventKey = GFSBeginCellEvent(cell)
    If Len(eventKey) = 0 Then Exit Sub
    LineAppEvents_CellChangedCore cell
EventFinished:
    GFSEndCellEvent eventKey
    Exit Sub
EventFailed:
    SaveLog Err, "LineAppEvents_CellChanged"
    Resume EventFinished
End Sub

Private Sub Document_DocumentOpened(ByVal doc As IVDocument)
    On Error GoTo Failed
    Set gfsSession = New c_GFSStencilSession
    gfsSession.Start ThisDocument
    Exit Sub
Failed:
    Debug.Print "GraFiS startup: " & Err.Number & " " & Err.Description
End Sub

Private Sub gfsSession_DrawingReady(ByVal drawingWindow As Visio.Window)
    Dim drawing As Visio.Document
    On Error GoTo Failed
    Set drawing = drawingWindow.Document
    If drawing.Type <> visTypeDrawing Then Exit Sub
    If drawing.ReadOnly Then Exit Sub
    drawingWindow.Activate
    If Not (Application.ActiveDocument Is drawing) Then Exit Sub
    If Not gfsUIReady Then
        GFSOpenedCore ThisDocument
        gfsUIReady = True
    End If
    GFSPrepareDrawing drawing
    Exit Sub
Failed:
    Err.Raise Err.Number, "GraFiS drawing initialization", Err.Description
End Sub

Private Sub GFSPrepareDrawing(ByVal drawing As Visio.Document)
    If drawing Is Nothing Then Exit Sub
    If drawing.Type <> visTypeDrawing Then Exit Sub
    If drawing.ReadOnly Then Exit Sub
    If Not (Application.ActiveDocument Is drawing) Then Exit Sub
    AddTimeUserCells
    sp_MastersImport
    sm_AddFireTime
    If drawing.DocumentSheet.CellExistsU("User.GFSColorTheme", 0) = 0 Then
        StyleExport
    End If
    If Not Application.ActivePage.PageSheet.CellExistsU("User.GFS_Aspect", 0) Then
        Application.ActivePage.PageSheet.AddNamedRow visSectionUser, "GFS_Aspect", visTagDefault
        Application.ActivePage.PageSheet.CellsU("User.GFS_Aspect").FormulaU = "1"
    End If
    drawing.GlueSettings = visGlueToGeometry + visGlueToGuides + visGlueToConnectionPoints + visGlueToVertices
End Sub

Private Sub Document_BeforeDocumentClose(ByVal doc As IVDocument)
    On Error Resume Next
    GFSCloseCore doc
    If Not gfsSession Is Nothing Then
        gfsSession.MarkRuntimeClean
        gfsSession.Stop
    End If
    Set gfsSession = Nothing
    gfsUIReady = False
End Sub
