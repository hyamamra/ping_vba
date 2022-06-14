Attribute VB_Name = "RunPingAndOutputCSV"
Option Explicit


Public Function Ping�����s����csv���o��(IP�A�h���X() As String, ByVal �� As Long, ByVal �^�C���A�E�g As Long)
    '���̃u�b�N�̓��K�w�� PingModule.ps1 ���쐬���A���s���܂��B
    '���s���ʂ� yyyy-mm-dd_hh-nn-ss_Ping.csv �ɏo�͂��܂��B
    
    
    Dim ���s�t�@�C���̃p�X As String
    ���s�t�@�C���̃p�X = ���s�t�@�C�����쐬(IP�A�h���X(), ��, �^�C���A�E�g)
    
    Call ps1�t�@�C�������s(���s�t�@�C���̃p�X, True)
End Function


Private Function ���s�t�@�C�����쐬(IP�A�h���X() As String, ByVal �� As Long, ByVal �^�C���A�E�g As Long) As String
    '�߂�l
        '���s�t�@�C���̃p�X As String
    
    '�Q�Ɛݒ�
        'Microsoft Scripting Runtime
    
    
    Dim �p�X As String
    �p�X = ThisWorkbook.Path & "\ps1\PingModule.ps1"
    ���s�t�@�C�����쐬 = �p�X
    
    Dim FSO As New FileSystemObject
    
    Dim ps1 As Object
    Set ps1 = FSO.OpenTextFile(�p�X, ForWriting, True, TristateTrue)
    
    Dim �R�}���h As String
    �R�}���h = �R�}���h���쐬(IP�A�h���X(), ��, �^�C���A�E�g)
    
    Call ps1.WriteLine(�R�}���h)
    Call ps1.Close
    
    Set ps1 = Nothing
    Set FSO = Nothing
End Function


Private Function �R�}���h���쐬(IP�A�h���X() As String, ByVal �� As Long, ByVal �^�C���A�E�g As Long) As String
    Dim ���� As String
    ���� = Format(Now, "yyyy-mm-dd_hh-nn-ss")
    
    Dim PingCSV�̃p�X As String
    PingCSV�̃p�X = ThisWorkbook.Path & "\Ping_" & ���� & ".csv'"
    
    Dim PingPS1�̃p�X As String
    PingPS1�̃p�X = ThisWorkbook.Path & "\ps1\Ping.ps1"
    
    Dim ShowBalloonPS1�̃p�X As String
    ShowBalloonPS1�̃p�X = ThisWorkbook.Path & "\ps1\ShowBalloon.ps1"
    
    Dim �R�}���h As String
    �R�}���h = "$IP_Addresses = @(" & vbNewLine
    
    Dim IP
    For Each IP In IP�A�h���X()
        �R�}���h = �R�}���h & "    '" & IP & "'" & vbNewLine
    Next
    
    �R�}���h = �R�}���h & ")" & vbNewLine & vbNewLine _
        & "$path = '" & PingCSV�̃p�X & vbNewLine _
        & "foreach ($IP in $IP_Addresses) {" & vbNewLine _
        & "    $value = " & PingPS1�̃p�X _
        & " $IP " & �� & " " & �^�C���A�E�g & vbNewLine _
        & "    Add-Content -Path $path -Value $value" & vbNewLine _
        & "}" & vbNewLine & vbNewLine _
        & "$message = 'Ping.csv�t�@�C���̎�荞�݂��J�n�ł��܂��B'" & vbNewLine _
        & "$title = 'Ping���������܂����B'" & vbNewLine _
        & "$icon = 'Info'" & vbNewLine _
        & "powershell.exe -Sta -NoProfile -WindowStyle Hidden " _
        & "-ExecutionPolicy RemoteSigned -File " & ShowBalloonPS1�̃p�X _
        & " $message $title $icon"
    
    �R�}���h���쐬 = �R�}���h
End Function


Private Function ps1�t�@�C�������s(ByVal �R�}���h As String, Optional �\�� As Boolean = False, Optional �������� As Boolean = False)
    'PowerShell�ŃR�}���h�����s���܂��B
    'https://docs.microsoft.com/ja-jp/powershell/module/microsoft.powershell.core/about/about_powershell_exe?view=powershell-5.1
    
    '����
        '�R�}���h As String
        
        '�\�� As Boolean = False
        '(�ȗ���) �E�B���h�E�̕\���ݒ�
        
        '�������� As Boolean = False
        '(�ȗ���) �R�}���h�̏I����ҋ@
    
    '�Q�Ɛݒ�
        'Windows Script Host Object Model
    
    
    Dim �\���ݒ� As Long
    If �\�� Then
        �\���ݒ� = 1
    End If
    
    Dim WSH As New IWshRuntimeLibrary.WshShell
    
    '-Sta : �V���O���X���b�h�Ŏ��s
    '-NoProfile : �J�n���Ƀv���t�@�C����ǂݍ��܂Ȃ�
    'WindowStyle : �E�C���h�E�̕\�����@�i�ŏ����j
    '-ExecutionPolicy : ���s�|���V�[�i���[�J���Ȃ珐���Ȃ��Ŏ��s�j
    ps1�t�@�C�������s = WSH.Run("powershell.exe -Sta -NoProfile " _
        & "-WindowStyle Minimized -ExecutionPolicy RemoteSigned " _
        & "-Command " & �R�}���h, �\���ݒ�, ��������)
    
    Set WSH = Nothing
End Function
