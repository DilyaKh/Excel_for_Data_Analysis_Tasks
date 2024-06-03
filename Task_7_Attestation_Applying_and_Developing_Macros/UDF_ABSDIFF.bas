Attribute VB_Name = "UDF_ABSDIFF"
Function ABSDIFF(arg1 As Variant, arg2 As Variant) As Variant
Attribute ABSDIFF.VB_Description = "���������� ������ �������� ���� ����������"
Attribute ABSDIFF.VB_ProcData.VB_Invoke_Func = " \n14"
' ������ ���������������� ������� ���������
' ������ ������� ���� ����������.

' Args: ��������� �� ���� ��� ��������� (��� ����� ���� �����).

' Returns: ���������� ����� - ������ �������� ���� ����������.

' ������������� �������� ����, ��� ��� ��������� - ��������� ����.
' ���� ��� ������� �� �����������, ������������ ������ #�����!.


    ' �������� ����, ��� ��� ��������� - ��������� ����:
    If Not IsNumeric(arg1) Or Not IsNumeric(arg2) Then
        ' ���� ��� ��������� �� �������� ���������,
        ' ������� ���������� ������ #�����!:
        ABSDIFF = CVErr(xlErrNum)
        ' ����� �� �������:
        Exit Function
    End If

    ' ���������� ������ �������� ���� ���������� � ��� �������:
    ABSDIFF = Abs(arg1 - arg2)

End Function

' -------------------------------------------------------------
