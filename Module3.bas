Attribute VB_Name = "Module3"
Sub ��������������()
'
' �������������� ������
'
Sheets("�����������").Unprotect Password:="0709"

'
    Range("L27").Select
    ActiveSheet.PivotTables("������� �������3").PivotFields( _
        "[TDSheet].[������ ���  ���].[������ ���  ���]").DrilledDown = True
    
Sheets("�����������").Protect Password:="0709", Contents:=True, AllowUsingPivotTables:=True, UserInterfaceOnly:=True

Range("AR3").Select
End Sub


' AllowFormattingCells:=True, AllowFiltering:=True,
