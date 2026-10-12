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

Private WithEvents gfsSession As c_GFSStencilSession
Private gfsUIReady As Boolean

Private Sub GFSCloseCore(ByVal doc As IVDocument)
    On Error Resume Next
    Set GDZSAppEvents = Nothing
'Процедура при закрытии документа

'---Очищаем переменную приложения
Set GDZSAppEvents = Nothing
End Sub

Private Sub GFSOpenedCore(ByVal doc As IVDocument)
    If Not GFSHasEditableDrawing() Then Exit Sub
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
    If drawing.DocumentSheet.CellExistsU("User.GFSColorTheme", 0) = 0 Then
        StyleExport
    End If
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
