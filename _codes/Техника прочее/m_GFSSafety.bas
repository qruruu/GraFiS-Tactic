Attribute VB_Name = "m_GFSSafety"
Option Explicit
Private activeCellEvents As Object

Public Function GFSResolveShape(ByVal shapeID As Long, Optional ByVal sourceShape As Visio.Shape = Nothing) As Visio.Shape
    If Not sourceShape Is Nothing Then
        If sourceShape.ID <> shapeID Then Err.Raise vbObjectError + 2401, "GFSResolveShape", "Shape ID does not match event context"
        Set GFSResolveShape = sourceShape
    Else
        If Application.ActivePage Is Nothing Then Err.Raise vbObjectError + 2402, "GFSResolveShape", "No active drawing page"
        Set GFSResolveShape = Application.ActivePage.Shapes.ItemFromID(shapeID)
    End If
End Function

Public Function GFSIsDrawingShape(ByVal shp As Visio.Shape) As Boolean
    On Error GoTo NotDrawing
    If shp Is Nothing Then Exit Function
    If shp.ContainingMasterID >= 0 Then Exit Function
    If shp.ContainingPage Is Nothing Then Exit Function
    If shp.Document.Type <> visTypeDrawing Then Exit Function
    GFSIsDrawingShape = True
NotDrawing:
End Function

Public Function GFSIsActiveDrawingShape(ByVal shp As Visio.Shape) As Boolean
    On Error GoTo NotActive
    If Not GFSIsDrawingShape(shp) Then Exit Function
    If Application.ActivePage Is Nothing Then Exit Function
    If Not (shp.ContainingPage Is Application.ActivePage) Then Exit Function
    GFSIsActiveDrawingShape = True
NotActive:
End Function

Public Function GFSBeginCellEvent(ByVal cell As Visio.Cell) As String
    Dim eventKey As String
    On Error GoTo NotEntered
    If cell Is Nothing Then Exit Function
    If Not GFSIsDrawingShape(cell.Shape) Then Exit Function
    eventKey = CStr(cell.Shape.Document.ID) & ":" & CStr(cell.Shape.ContainingPage.ID) & _
        ":" & CStr(cell.Shape.ID) & ":" & cell.NameU
    If activeCellEvents Is Nothing Then Set activeCellEvents = CreateObject("Scripting.Dictionary")
    If activeCellEvents.Exists(eventKey) Then Exit Function
    activeCellEvents.Add eventKey, True
    GFSBeginCellEvent = eventKey
NotEntered:
End Function

Public Sub GFSEndCellEvent(ByVal eventKey As String)
    On Error Resume Next
    If Len(eventKey) = 0 Then Exit Sub
    If activeCellEvents Is Nothing Then Exit Sub
    If activeCellEvents.Exists(eventKey) Then activeCellEvents.Remove eventKey
End Sub

Public Function GFSCurrentTime(ByVal shp As Visio.Shape) As Date
    Dim docSheet As Visio.Shape
    If shp Is Nothing Then Err.Raise vbObjectError + 2403, "GFSCurrentTime", "No source shape"
    If shp.Document.Type <> visTypeDrawing Then Err.Raise vbObjectError + 2404, "GFSCurrentTime", "Time properties belong to a drawing, not a stencil"
    Set docSheet = shp.Document.DocumentSheet
    If Not docSheet.CellExistsU("User.FireTime", 0) Then
        docSheet.AddNamedRow visSectionUser, "FireTime", visTagDefault
        docSheet.CellsU("User.FireTime").FormulaU = "Now()"
    End If
    If Not docSheet.CellExistsU("User.CurrentTime", 0) Then
        docSheet.AddNamedRow visSectionUser, "CurrentTime", visTagDefault
        docSheet.CellsU("User.CurrentTime").FormulaU = "User.FireTime"
    End If
    GFSCurrentTime = docSheet.CellsU("User.CurrentTime").Result(visDate)
End Function

Public Function GFSHasEditableDrawing() As Boolean
    Dim drawing As Visio.Document, drawingWindow As Visio.Window
    On Error GoTo NotDrawing
    Set drawing = Application.ActiveDocument
    Set drawingWindow = Application.ActiveWindow
    If drawing Is Nothing Then Exit Function
    If drawingWindow Is Nothing Then Exit Function
    If drawing.Type <> visTypeDrawing Then Exit Function
    If drawing.ReadOnly Then Exit Function
    If drawingWindow.Type <> visDrawing Then Exit Function
    If drawingWindow.Page Is Nothing Then Exit Function
    If Not (drawingWindow.Document Is drawing) Then Exit Function
    GFSHasEditableDrawing = True
NotDrawing:
End Function
