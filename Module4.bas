Attribute VB_Name = "Module4"
Sub СбросФильтров()
Attribute СбросФильтров.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ФИЛЬТРАЦИЯ
Dim a As Integer
a = MsgBox("Процесс может занять длительное время, Вы так же можете снять фильтры вручную", 48)
msg = MsgBox("Хотите СБРОСИТЬ свойства ОБЪЕКТА СРАВНЕНИЯ?", vbYesNo, "????")
If msg = vbYes Then
 

Sheets("Калькулятор").Unprotect Password:="0709"
 
'сброс фильтра сводной таблицы

    Range("C27").Select
    ActiveSheet.PivotTables("Сводная таблица3").ClearAllFilters

Sheets("Калькулятор").Protect Password:="0709", Contents:=True, AllowUsingPivotTables:=True, UserInterfaceOnly:=True

 'установка фильтра
 Range("B17").Select
  ActiveSheet.PivotTables("Сводная таблица3").PivotFields( _
        "[TDSheet].[Город].[Город]").VisibleItemsList = Array( _
        "[TDSheet].[Город].&[Новосибирск]")
   
   Range("AP3").Select
End If


 'сброс фильтра объекта оценки
Sheets("Калькулятор").Unprotect Password:="0709"
msg = MsgBox("Хотите СБРОСИТЬ свойства ОБЪЕКТА ОЦЕНКИ?", vbYesNo, "????")
If msg = vbYes Then
 
    Sheets("Калькулятор").Cells(5, 5) = "введите данные"
    
    Sheets("Калькулятор").Cells(5, 7) = "введите данные"
    Sheets("Калькулятор").Cells(5, 8) = "введите данные"
    Sheets("Калькулятор").Cells(5, 8) = "введите данные"
    Sheets("Калькулятор").Cells(5, 9) = "введите данные"
    Sheets("Калькулятор").Cells(5, 10) = "введите данные"
 
 
 Range("AE26:AE70000").Select
    Selection.Replace What:="Да", Replacement:="Нет", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    
Sheets("Калькулятор").Cells(3, 44) = 0
    
    
Range("AR3").Select
End If

        
Sheets("Калькулятор").Protect Password:="0709", Contents:=True, AllowUsingPivotTables:=True, UserInterfaceOnly:=True

End Sub

