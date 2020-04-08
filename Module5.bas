Attribute VB_Name = "Module5"
Sub Печать()
Attribute Печать.VB_ProcData.VB_Invoke_Func = " \n14"
'
'Подгатовка данных для Печать Макрос
'
 Sheets("Калькулятор").Select
    Range("E2:J5").Select
    Selection.CopyPicture Appearance:=xlScreen, Format:=xlBitmap
    Sheets("На печать").Select
    Range("B8").Select
    ActiveSheet.Paste
    Selection.ShapeRange.IncrementLeft 7.5
    Selection.ShapeRange.IncrementTop -5.6249606299
    Range("B16").Select
    Sheets("Калькулятор").Select
    ActiveWindow.SmallScroll ToRight:=2
    Range("AO2:AR18").Select
    Selection.Copy
    Application.CutCopyMode = False
    Selection.CopyPicture Appearance:=xlScreen, Format:=xlBitmap
    Sheets("На печать").Select
    Range("B16").Select
    ActiveSheet.Paste
    Selection.ShapeRange.IncrementLeft 7.5
    Selection.ShapeRange.IncrementTop 6.5624409449
    Range("B27").Select
    Sheets("Калькулятор").Select
    Range("B26:L134").Select
    ActiveWindow.ScrollRow = 26
    Range("B26:L134,AC26:AC134").Select
    Range("AC26").Activate
    ActiveWindow.SmallScroll ToRight:=6
    ActiveWindow.ScrollRow = 26
    Range("B26:L134,AC26:AC134,AP26:AP134").Select
    Range("AP26").Activate
    ActiveWindow.SmallScroll ToRight:=1
    ActiveWindow.ScrollRow = 26
    Range("B26:L134,AC26:AC134,AP26:AP134,AR26:AR134").Select
    Range("AR26").Activate
    Selection.Copy
    Sheets("На печать").Select
    Range("B33").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=6
    
    'Обработка
    
    
    
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection
        .HorizontalAlignment = xlLeft
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
    End With
    ActiveSheet.Range("$B$32:$P$141").AutoFilter Field:=1, Criteria1:=Array( _
        "1"), Operator:=xlFilterValues
        ' , "11", "12", "14", "19", "2", "20", "21", "22", "24", "25", "27", "30", "32", "33", "36", _
        "4", "42", "5", "6", "77"
        Range("M9:P5").Select
   
End Sub
