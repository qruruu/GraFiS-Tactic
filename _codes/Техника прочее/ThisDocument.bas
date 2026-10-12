VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Dim WithEvents vdAppEventsTech2 As Visio.Application
Attribute vdAppEventsTech2.VB_VarHelpID = -1

Private WithEvents gfsSession As c_GFSStencilSession
Private gfsUIReady As Boolean

Private Sub GFSCloseCore(ByVal doc As IVDocument)
    On Error Resume Next
    Set vdAppEventsTech2 = Nothing
'Процедура при закрытии документа

'---Очищаем переменную приложения
Set vdAppEventsTech2 = Nothing
End Sub

Private Sub GFSOpenedCore(ByVal doc As IVDocument)
    If Not GFSHasEditableDrawing() Then Exit Sub
'Процедура при открытии документа

'---Добавляем ячейки "User.FireTime", "User.CurrentTime"
    AddTimeUserCells

'---Показываем окно свойств
    Application.ActiveWindow.Windows.ItemFromID(visWinIDCustProp).Visible = True
    
'---Обновляем/экспортируем в активный документ стили трафарета
    '---Проверяем не является ли активный документ документом цветовой схемы
    If Application.ActiveDocument.DocumentSheet.CellExists("User.GFSColorTheme", 0) = 0 Then
        StyleExport
    End If
    
'---Объявляем переменную приложения для дальнейшего реагирования на изменение содержимого ячеек
    Set vdAppEventsTech2 = Visio.Application
    
''---Проверяем наличие обновлений
'    fmsgCheckNewVersion.CheckUpdates

End Sub

Private Sub vdAppEventsTech2_CellChangedCore(ByVal cell As IVCell)
'Процедура обновления списков в фигурах
Dim ShpInd As Integer

    On Error GoTo EX
'---Проверяем имя ячейки

    If cell.Name = "Prop.Set" Then
        '---Запускаем процедуру получения списков моделей
'        ShpInd = cell.Shape.ID
        ModelsListImport cell.Shape
    ElseIf cell.Name = "Prop.Model" Then
        '---Процедура получения ТТХ - СДЕЛАТЬ
        If Not IsShapeHaveCalloutAndDropFirst(cell.Shape) Then
'            ShpInd = cell.Shape.ID
            GetTTH cell.Shape
        End If
    End If

'В случае, если произошло изменение не нужной ячейки прекращаем событие
Exit Sub
EX:
    SaveLog Err, "vdAppEventsTech2_CellChanged"
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

Private Sub vdAppEventsTech2_CellChanged(ByVal cell As IVCell)
    Dim eventKey As String
    On Error GoTo EventFailed
    eventKey = GFSBeginCellEvent(cell)
    If Len(eventKey) = 0 Then Exit Sub
    vdAppEventsTech2_CellChangedCore cell
EventFinished:
    GFSEndCellEvent eventKey
    Exit Sub
EventFailed:
    SaveLog Err, "vdAppEventsTech2_CellChanged"
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
