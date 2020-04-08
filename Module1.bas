Attribute VB_Name = "Module1"

Sub обновление()
'сохранение файла Data в связи с не корректной выгрузкой из 1С
FilePath = "Q:\!calc.start\Data.xlsx"
    Workbooks.Open Filename:=FilePath
   ' Workbooks("Z:\!Отдел маркетинга и рекламы\БАЗА ДАННЫХ\Передача дел\Пономарев\Data.xlsx").Activate
   ' ActiveWorkbook.Sheets(1).Activate
    'Range("A1").Value = 5
   ActiveWorkbook.Save 'As Filename:=FilePath
    Workbooks("Data.xlsx").Close ' Filename:=FilePath

ActiveWorkbook.Unprotect Password:="0709"
Sheets("Калькулятор").Unprotect Password:="0709"
ActiveSheet.Shapes.Range(Array("Button 1")).Select
    With Selection.Font
        .Name = "Arial"
        .FontStyle = "Полужирный"
        .Size = 14
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
    End With
Application.Calculation = xlAutomatic
ActiveWorkbook.RefreshAll

' Автопересчет Макрос
 
 
' обновление Макрос


     
  '  Range("A28").Select
   
  '  ActiveSheet.PivotTables("Сводная таблица3").PivotCache.Refresh
   
   Sheets("Калькулятор").Protect Password:="0709", Contents:=True, AllowFiltering:=True, AllowUsingPivotTables:=True, UserInterfaceOnly:=True
ActiveWorkbook.Protect Password:="0709"

Dim a As Integer
Dim b As Variant
b = Sheets("Калькулятор").Cells(3, 44)
If b = 0 Then
a = MsgBox("!!!УКАЖИТЕ СВОЙСТВА ОБЪЕКТА ОЦЕНКИ", 48)
'MsgBox "!!!УКАЖИТЕ СВОЙСТВА ОБЪЕКТА ОЦЕНКИ"
Range("AR3").Select
End If
Range("AR3").Select
End Sub

