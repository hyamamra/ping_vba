VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PingForm 
   Caption         =   "Ping Form"
   ClientHeight    =   3248
   ClientLeft      =   96
   ClientTop       =   416
   ClientWidth     =   4568
   OleObjectBlob   =   "PingForm.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "PingForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub StartButton_Click()
    Call Ping
End Sub


Private Function Ping()
    Dim IP�A�h���X() As String
    IP�A�h���X() = IP�A�h���X���擾()
    
    Dim �� As Long
    �� = PingForm.��.Value
    
    Dim �^�C���A�E�g As Long
    �^�C���A�E�g = PingForm.�^�C���A�E�g.Value
    
    Call Ping�����s����csv���o��(IP�A�h���X(), ��, �^�C���A�E�g)
End Function


Private Sub ��_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call Esc�L�[�ŏI��(KeyCode)
End Sub


Private Sub �^�C���A�E�g_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call Esc�L�[�ŏI��(KeyCode)
End Sub


Private Sub StartButton_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call Esc�L�[�ŏI��(KeyCode)
End Sub


Private Function Esc�L�[�ŏI��(ByVal KeyCode As Long)
    If KeyCode = 27 Then
        Call Unload(PingForm)
    End If
End Function
