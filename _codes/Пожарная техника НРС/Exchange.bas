Attribute VB_Name = "Exchange"
'------------------------Модуль для процедур импорта ТТХ-------------------
'-----------------------------------Управляющий блок-------------------------------------------------
Public Sub GetTTH(ShpIndex As Long, Optional ByVal sourceShape As Visio.Shape = Nothing)
'Управляющая процедура импорта ТТХ автомобилей
'---Объявляем переменные
Dim shp As Visio.Shape
Dim IndexPers As Integer

'---Проверяем к какой именно фигуре относится данная ячейка
    Set shp = GFSResolveShape(ShpIndex, sourceShape)
    IndexPers = shp.Cells("User.IndexPers")
    
'---Запускаем процедуру получения относительного списка Модели(Набор) для текущей фигуры
Select Case IndexPers
    Case Is = 1
        GetValuesOfCellsFromTable ShpIndex, "З_Автоцистерны", shp
    Case Is = 2
        GetValuesOfCellsFromTable ShpIndex, "З_АНР", shp
    Case Is = 3
        GetValuesOfCellsFromTable ShpIndex, "З_АЛ", shp
    Case Is = 4
        GetValuesOfCellsFromTable ShpIndex, "З_АКП", shp
    Case Is = 5
        GetValuesOfCellsFromTable ShpIndex, "З_АСО", shp
    Case Is = 6
        GetValuesOfCellsFromTable ShpIndex, "З_АТ", shp
    Case Is = 7
        GetValuesOfCellsFromTable ShpIndex, "З_АД", shp
    Case Is = 8
        GetValuesOfCellsFromTable ShpIndex, "З_ПНС", shp
    Case Is = 9
        GetValuesOfCellsFromTable ShpIndex, "З_АА", shp
    Case Is = 10
        GetValuesOfCellsFromTable ShpIndex, "З_АВ", shp
    Case Is = 11
        GetValuesOfCellsFromTable ShpIndex, "З_АКТ", shp
    Case Is = 12
        GetValuesOfCellsFromTable ShpIndex, "З_АП", shp
    Case Is = 13
        GetValuesOfCellsFromTable ShpIndex, "З_АГВТ", shp
    Case Is = 14
        GetValuesOfCellsFromTable ShpIndex, "З_АГТ", shp
    Case Is = 15
        GetValuesOfCellsFromTable ShpIndex, "З_АГДЗС", shp
    Case Is = 16
        GetValuesOfCellsFromTable ShpIndex, "З_ПКС", shp
    Case Is = 17
        GetValuesOfCellsFromTable ShpIndex, "З_ЛБ", shp
    Case Is = 18
        GetValuesOfCellsFromTable ShpIndex, "З_АСА", shp
    Case Is = 19
        GetValuesOfCellsFromTable ShpIndex, "З_АШ", shp
    Case Is = 20
        GetValuesOfCellsFromTable ShpIndex, "З_АР", shp
    Case Is = 161
        GetValuesOfCellsFromTable ShpIndex, "З_АЦЛ", shp
    Case Is = 162
        GetValuesOfCellsFromTable ShpIndex, "З_АЦКП", shp
    Case Is = 163
        GetValuesOfCellsFromTable ShpIndex, "З_АПП", shp
End Select



End Sub


