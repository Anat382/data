Attribute VB_Name = "Module3"
Sub Разгруппировка()
'
' Разгруппировка Макрос
'
Sheets("Калькулятор").Unprotect Password:="0709"

'
    Range("L27").Select
    ActiveSheet.PivotTables("Сводная таблица3").PivotFields( _
        "[TDSheet].[Группа цен  исх].[Группа цен  исх]").DrilledDown = True
    
Sheets("Калькулятор").Protect Password:="0709", Contents:=True, AllowUsingPivotTables:=True, UserInterfaceOnly:=True

Range("AR3").Select
End Sub


' AllowFormattingCells:=True, AllowFiltering:=True,
