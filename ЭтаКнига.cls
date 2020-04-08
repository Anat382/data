VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ЭтаКнига"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Workbook_Open()
Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"
Dim a As Integer
a = MsgBox("!!!ЧТОБЫ ПРОДОЛЖИТЬ РАБОТУ,НАЖМИТЕ НА КНОПКУ обновить сводную таблицу", 16)
  'MsgBox "!!!ЧТОБЫ ПРОДОЛЖИТЬ РАБОТУ,НАЖМИТЕ НА КНОПКУ обновить сводную таблицу"
End Sub

Private Sub App_WorkbookOpen(ByVal Wb As Workbook) 'при открытии существующей книги
  App.Calculation = xlCalculationManual
End Sub


' Спрятать
Sub HideRibbon()
Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"
End Sub
' Открыть
Sub ShowRibbon()
Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
End Sub


Private Sub Workbook_BeforeClose(Cancel As Boolean)
ThisWorkbook.Saved = True
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
Dim xName As String
xName = "CancelBeforeSave"
 
If Not Evaluate("=ISREF('" & xName & "'!A1)") Then
    Sheets.Add(after:=Worksheets(Worksheets.Count)).Name = xName & ""
    Sheets(xName & "").Move after:=Worksheets(Worksheets.Count)
    Sheets(xName & "").Visible = False
    Exit Sub
End If
    Cancel = True
End Sub






