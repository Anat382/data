Attribute VB_Name = "Module2"
Sub группировка()
Attribute группировка.VB_ProcData.VB_Invoke_Func = " \n14"
'
' группировка Макрос
'
Sheets("Калькулятор").Unprotect Password:="0709"

'
    Range("L27").Select
    ActiveSheet.PivotTables("Сводная таблица3").PivotFields( _
        "[TDSheet].[Группа цен  исх].[Группа цен  исх]").DrilledDown = False
        
Sheets("Калькулятор").Protect Password:="0709", Contents:=True, AllowUsingPivotTables:=True, UserInterfaceOnly:=True
Dim a As Integer
a = MsgBox("Для расчета рекомендованной цены, нажмите РАЗГРУППИРОВАТЬ", 48)
'a = MsgBox("Критичное сообщение, сообщение об ошибке", 16)
'a = MsgBox("Сообщение с вопросом", 32)
'a = MsgBox("Предупреждающее сообщение", 48)
'a = MsgBox("Информационное сообщение", 64)
Range("AG26").Select

End Sub
