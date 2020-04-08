VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Лист1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Worksheet_Change(ByVal Target As Range)
  'изменение цены
 
    Select Case Target.Address
        Case Range("B1").Address
            MsgBox "!!!Произошло изменение " & "ИНТЕРВАЛА ЦЕН,обновите сводную таблицу"
            Sheets("Калькулятор").Unprotect Password:="0709"
            ActiveSheet.Shapes.Range(Array("Button 1")).Select
    With Selection.Font
        .Name = "Arial"
        .FontStyle = "Полужирный"
        .Size = 15
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 3
    End With
Sheets("Калькулятор").Protect Password:="0709", Contents:=True, AllowFiltering:=True, AllowUsingPivotTables:=True, UserInterfaceOnly:=True
  'изменение площади
            Case Range("B2").Address
            MsgBox "!!!Произошло изменение " & "ИНТЕРВАЛА ПЛОЩАДИ,обновите сводную таблицу"
           Sheets("Калькулятор").Unprotect Password:="0709"
            ActiveSheet.Shapes.Range(Array("Button 1")).Select
    With Selection.Font
        .Name = "Arial"
        .FontStyle = "Полужирный"
        .Size = 15
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 3
    End With
Sheets("Калькулятор").Protect Password:="0709", Contents:=True, AllowFiltering:=True, AllowUsingPivotTables:=True, UserInterfaceOnly:=True
       
       Case Range("AR3").Address
            MsgBox "!!!Укажите остальные свойства объекта оценки"
            Range("E5").Select
  End Select
  End Sub
  





