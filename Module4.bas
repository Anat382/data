Attribute VB_Name = "Module4"
Sub �������������()
Attribute �������������.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ����������
Dim a As Integer
a = MsgBox("������� ����� ������ ���������� �����, �� ��� �� ������ ����� ������� �������", 48)
msg = MsgBox("������ �������� �������� ������� ���������?", vbYesNo, "????")
If msg = vbYes Then
 

Sheets("�����������").Unprotect Password:="0709"
 
'����� ������� ������� �������

    Range("C27").Select
    ActiveSheet.PivotTables("������� �������3").ClearAllFilters

Sheets("�����������").Protect Password:="0709", Contents:=True, AllowUsingPivotTables:=True, UserInterfaceOnly:=True

 '��������� �������
 Range("B17").Select
  ActiveSheet.PivotTables("������� �������3").PivotFields( _
        "[TDSheet].[�����].[�����]").VisibleItemsList = Array( _
        "[TDSheet].[�����].&[�����������]")
   
   Range("AP3").Select
End If


 '����� ������� ������� ������
Sheets("�����������").Unprotect Password:="0709"
msg = MsgBox("������ �������� �������� ������� ������?", vbYesNo, "????")
If msg = vbYes Then
 
    Sheets("�����������").Cells(5, 5) = "������� ������"
    
    Sheets("�����������").Cells(5, 7) = "������� ������"
    Sheets("�����������").Cells(5, 8) = "������� ������"
    Sheets("�����������").Cells(5, 8) = "������� ������"
    Sheets("�����������").Cells(5, 9) = "������� ������"
    Sheets("�����������").Cells(5, 10) = "������� ������"
 
 
 Range("AE26:AE70000").Select
    Selection.Replace What:="��", Replacement:="���", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    
    
Sheets("�����������").Cells(3, 44) = 0
    
    
Range("AR3").Select
End If

        
Sheets("�����������").Protect Password:="0709", Contents:=True, AllowUsingPivotTables:=True, UserInterfaceOnly:=True

End Sub

