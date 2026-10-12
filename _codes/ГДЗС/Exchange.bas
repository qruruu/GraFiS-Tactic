Attribute VB_Name = "Exchange"
'------------------------Модуль для процедур импорта ТТХ-------------------
'-----------------------------------Управляющий блок-------------------------------------------------
Public Sub GetTTH(ShpIndex As Long, Optional ByVal sourceShape As Visio.Shape = Nothing)
'Управляющая процедура импорта ТТХ аппаратов
'---Объявляем переменные
Dim shp As Visio.Shape
Dim IndexPers As Integer

On Error GoTo ex

'---Проверяем к какой именно фигуре относится данная ячейка
    Set shp = GFSResolveShape(ShpIndex, sourceShape)
    IndexPers = shp.Cells("User.IndexPers")
    
'---Запускаем процедуру получения значений ТТХ для текущей фигуры
    Select Case IndexPers
        Case Is = 46
            GetValuesOfCellsFromTable ShpIndex, "ДАСВ", shp
        Case Is = 90
            GetValuesOfCellsFromTable ShpIndex, "ДАСК", shp
        Case Is = 49 'Дымососы
            FogRMKGetValuesOfCellsFromTable ShpIndex, "Дымососы", shp
    End Select

Exit Sub
ex:
    SaveLog Err, "GetTTH", CStr(ShpIndex)
End Sub



