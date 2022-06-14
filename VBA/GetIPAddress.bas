Attribute VB_Name = "GetIPAddress"
Option Explicit


Public Function IP�A�h���X���擾() As String()
    'IP_Address�V�[�g��A�񂩂�IP�A�h���X�̂ݎ擾���܂��B
    
    '�߂�l
        'IP�A�h���X() As String
    
    
    Dim IP_AddressSheet As Worksheet
    Set IP_AddressSheet = ThisWorkbook.Worksheets("IP_Address")
    
    Dim �ŏI�s As Long
    �ŏI�s = �ŏI�s���擾(IP_AddressSheet, ��:=1)
    
    '(�s, 1)�̓񎟌��z��
    Dim A��̒l() As Variant
    With IP_AddressSheet
        A��̒l() = .Range(.Range("A1"), .Range("A" & �ŏI�s)).Value
    End With
    
    Dim IP�A�h���X() As String
    Dim index As Long
    index = 0
    
    Dim �s
    For �s = 1 To �ŏI�s
        Dim �l As String
        �l = A��̒l(�s, 1)
        If IsIPv4�A�h���X(�l) Then
            ReDim Preserve IP�A�h���X(index)
            IP�A�h���X(index) = �l
            index = index + 1
        End If
    Next
    
    IP�A�h���X���擾 = IP�A�h���X()
End Function


Private Function �ŏI�s���擾(�V�[�g As Worksheet, �� As Long) As Long
    '�V�[�g�̏I�[���������ɒl�̓������Z����T�����܂��B
    '�l�̓������Z��������΍s�ԍ���Ԃ��܂��B
    '�l�̓������Z�����Ȃ����1��Ԃ��܂��B
    
    '����
        '�V�[�g As Worksheet
        '�T������V�[�g
        
        '�� As Long
        '�T�������
    
    
    Dim �I�[ As Long
    �I�[ = �V�[�g.Rows.Count
    
    Dim �ŏI�s As Long
    If �V�[�g.Cells(�I�[, ��) = "" Then
        �ŏI�s = �V�[�g.Cells(�I�[, ��).End(xlUp).Row
    Else
        �ŏI�s = �V�[�g.Cells(�I�[, ��)
    End If
    
    �ŏI�s���擾 = �ŏI�s
End Function


Private Function IsIPv4�A�h���X(ByVal �l) As Boolean
    'IPv4�A�h���X�\�L�ł���� True ��Ԃ��܂��B
    '10�i���\�L�łȂ���� False ��Ԃ��܂��B
    '�h�b�g��3����A�e���̐��l�� 0 �` 255 �ł����
    'IPv4�A�h���X�\�L�ł���Ɣ��f���܂��B
    
    
    Dim �e��() As String
    �e��() = Split(CStr(�l), ".")
    
    Dim �h�b�g�̐� As Long
    �h�b�g�̐� = UBound(�e��())
    If Not �h�b�g�̐� = 3 Then
        IsIPv4�A�h���X = False
        Exit Function
    End If
    
    Dim ���l
    For Each ���l In �e��()
        If Not ((0 <= ���l) And (���l <= 255)) Then
            IsIPv4�A�h���X = False
            Exit Function
        End If
    Next
    
    IsIPv4�A�h���X = True
End Function
