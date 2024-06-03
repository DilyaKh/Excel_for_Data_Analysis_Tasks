Attribute VB_Name = "Select_and_Fill"


' -------------------------------------------------------------

Sub Fill_yellow_from_selected_cell_to_last_non_empty_cell_in_column()
Attribute Fill_yellow_from_selected_cell_to_last_non_empty_cell_in_column.VB_Description = "�������� ������� ������ ������ �� ���������� ������ �� ��������� ����������� ������ � �������"
Attribute Fill_yellow_from_selected_cell_to_last_non_empty_cell_in_column.VB_ProcData.VB_Invoke_Func = " \n14"
' �������� ������� ������ ������ �� ���������� ������ 
' �� ��������� ����������� ������ � �������

    ' ��������� ���������� ���������:
    Range(Selection, Selection.End(xlDown)).Select
    
    ' ���������� ��� ����������� ���������
    ' ������� ���� ����� ������ ������:
    Selection.Interior.Color = vbYellow
    
End Sub


' -------------------------------------------------------------

Sub Fill_yellow_from_top_cell_to_last_non_empty_cell_in_column()
' �������� ������� ������ ������ �� ������ ������ ������� �������
' �� ��������� ����������� ������ � �������

    ' ���������� ����������, � ������� ����� ���������
    ' ����� ��������� �������:
    Dim activeColumn As Integer
    

    ' ���������� � ���������� ������ ��������� �������:
    activeColumn = ActiveCell.Column
    
    ' ��������� ���������� ���������:
    Range(Cells(1, activeColumn), Selection.End(xlDown)).Select
    
    ' ���������� ��� ����������� ���������
    ' ������� ���� ����� ������ ������:
    Selection.Interior.Color = vbYellow

End Sub


' -------------------------------------------------------------

Sub Fill_yellow_entire_column_of_selected_cell()
' �������� ������ ������ ���� �������,
' � ������� ������������� ���������� ������

    ' ���������� ����������, � ������� ����� ���������
    ' ����� ��������� �������:
    Dim activeColumn As Integer
    
    
    ' ���������� � ���������� ������ ��������� �������:
    activeColumn = ActiveCell.Column
    
    ' ��������� ���������� ���������:
    Range(Columns(activeColumn), Columns(activeColumn)).Select
    
    ' ���������� ��� ����������� ���������
    ' ������� ���� ����� ������ ������:
    Selection.Interior.Color = vbYellow

End Sub

' -------------------------------------------------------------
