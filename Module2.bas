Attribute VB_Name = "Module2"
Sub �����������()
Attribute �����������.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ����������� ������
'
Sheets("�����������").Unprotect Password:="0709"

'
    Range("L27").Select
    ActiveSheet.PivotTables("������� �������3").PivotFields( _
        "[TDSheet].[������ ���  ���].[������ ���  ���]").DrilledDown = False
        
Sheets("�����������").Protect Password:="0709", Contents:=True, AllowUsingPivotTables:=True, UserInterfaceOnly:=True
Dim a As Integer
a = MsgBox("��� ������� ��������������� ����, ������� ���������������", 48)
'a = MsgBox("��������� ���������, ��������� �� ������", 16)
'a = MsgBox("��������� � ��������", 32)
'a = MsgBox("��������������� ���������", 48)
'a = MsgBox("�������������� ���������", 64)
Range("AG26").Select

End Sub
