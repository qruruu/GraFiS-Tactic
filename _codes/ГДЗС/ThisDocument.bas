VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Dim WithEvents GDZSAppEvents As Visio.Application
Attribute GDZSAppEvents.VB_VarHelpID = -1

Private Sub Document_BeforeDocumentClose(ByVal doc As IVDocument)
    On Error Resume Next
    Set GDZSAppEvents = Nothing
'Процедура при закрытии документа

'---Очищаем переменную приложения
Set GDZSAppEvents = Nothing
End Sub

Private Sub Document_DocumentOpened(ByVal doc As IVDocument)
'Процедура при открытии документа

'---Добавляем ячейки "User.FireTime", "User.CurrentTime"
    AddTimeUserCells

    On Error Resume Next
'---Показываем окно свойств
    Application.ActiveWindow.Windows.ItemFromID(visWinIDCustProp).Visible = True

'---ОБновляем/экспортируем в активный документ стили трафарета
    '---Проверяем не является ли активный документ документом цветовой схемы
    If Application.ActiveDocument.DocumentSheet.CellExists("User.GFSColorTheme", 0) = 0 Then
        StyleExport
    End If

'---Объявляем переменную приложения для дальнейшего реагирования на изменение содержимого ячеек
Set GDZSAppEvents = Visio.Application

''---Проверяем наличие обновлений
'    fmsgCheckNewVersion.CheckUpdates

End Sub

Private Sub GDZSAppEvents_CellChangedCore(ByVal cell As IVCell)
'Процедура обновления списков в фигурах
Dim ShpInd As Long

'---Проверяем, имеет ли ячейка родительскую фигуру (например, является стилем)
    If cell.Shape Is Nothing Then Exit Sub
    
'---Проверяем имя ячейки
    If Not IsShapeLinkedToDataAndDropFirst(cell.Shape) Then
        If cell.Name = "Prop.AirDevice" Then
            ShpInd = cell.Shape.ID
            '---Запускаем процедуру получения списков аппаратов
            AirDevicesListImport ShpInd, cell.Shape
            '---Запускаем процедуру получения ТТХ для указанной модели ДАСВ
            GetTTH ShpInd, cell.Shape
        ElseIf cell.Name = "Prop.FogRMK" Then
            ShpInd = cell.Shape.ID
            '---Запускаем процедуру получения ТТХ для указанной модели Дымососов
            GetTTH ShpInd, cell.Shape
        End If
    End If

'В случае, если произошло изменение не нужной ячейки прекращаем событие
End Sub

Private Sub AddTimeUserCells()
'Прока добавляет ячейки "User.FireTime", "User.CurrentTime"
Dim docSheet As Visio.Shape
Dim cell As Visio.cell

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

Private Sub GDZSAppEvents_CellChanged(ByVal cell As IVCell)
    Dim eventKey As String
    On Error GoTo EventFailed
    eventKey = GFSBeginCellEvent(cell)
    If Len(eventKey) = 0 Then Exit Sub
    GDZSAppEvents_CellChangedCore cell
EventFinished:
    GFSEndCellEvent eventKey
    Exit Sub
EventFailed:
    SaveLog Err, "GDZSAppEvents_CellChanged"
    Resume EventFinished
End Sub
