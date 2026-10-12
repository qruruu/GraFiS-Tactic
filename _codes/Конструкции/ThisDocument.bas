VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private WithEvents app As Visio.Application
Attribute app.VB_VarHelpID = -1
Private cellChangedCount As Long
Const cellChangedInterval = 1000

Private ButEvent As c_Buttons


Private WithEvents gfsSession As c_GFSStencilSession
Private gfsUIReady As Boolean

Private Sub app_CellChanged(ByVal Cell As IVCell)
    If ButEvent Is Nothing Then Exit Sub
'---Один раз в выполняем обновление иконок на кнопках
    cellChangedCount = cellChangedCount + 1
    If cellChangedCount > cellChangedInterval Then
        ButEvent.PictureRefresh
        cellChangedCount = 0
    End If
End Sub

Private Sub GFSOpenedCore(ByVal doc As IVDocument)
    If Not GFSHasEditableDrawing() Then Exit Sub
    OpenDoc
End Sub


Private Sub GFSCloseCore(ByVal doc As IVDocument)
    On Error Resume Next
    Set app = Nothing
    CloseDoc
End Sub


Private Sub OpenDoc()
    If Not GFSHasEditableDrawing() Then Exit Sub
'---Показываем окно свойств
    Application.ActiveWindow.Windows.ItemFromID(visWinIDCustProp).Visible = True
    
'---Импортируем мастера
    sp_MastersImport
    
'---Создаем панель управления "Конструкции" и добавляем на нее кнопки
    AddTB_Constructions
    AddButtons
    
'---Активируем объект отслеживания нажатия кнопок
    Set ButEvent = New c_Buttons
    
'---Активируем объект отслеживания изменений в приложении для 201х версий
    If Application.version > 12 Then
        Set app = Visio.Application
        cellChangedCount = cellChangedInterval - 10
    End If
    
''---Проверяем наличие обновлений
'    fmsgCheckNewVersion.CheckUpdates
End Sub

Private Sub CloseDoc()
'Обрабатываем закрытие документа

'---Деактивируем объект отслеживания нажатия кнопок
    Set ButEvent = Nothing
    
'---Деактивируем объект отслеживания изменений в приложении для 201х версий
    If Application.version > 12 Then Set app = Nothing
    
'---Удаляем панель управления "СпецФункции"
    RemoveTB_Constructions
End Sub


Private Sub sp_MastersImport()
    If Not GFSHasEditableDrawing() Then Exit Sub
'---Импортируем мастера

'---Масштаб 1:200
    MasterImportSub "Забор"
    MasterImportSub "Забор2"
    MasterImportSub "Забор3"
    MasterImportSub "Забор4"
    MasterImportSub "ЖДПолотно"
    MasterImportSub "ЖДПолотно2"
    MasterImportSub "Обрыв"
    MasterImportSub "Ров"
    MasterImportSub "Насыпь"
    MasterImportSub "ТрамвайныеПути"
'---Масштаб 1:1000
    MasterImportSub "Забор_1000"
    MasterImportSub "Забор2_1000"
    MasterImportSub "Забор3_1000"
    MasterImportSub "Забор4_1000"
    MasterImportSub "ЖДПолотно_1000"
    MasterImportSub "ЖДПолотно2_1000"
    
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
    sp_MastersImport
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
