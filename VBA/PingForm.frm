VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PingForm 
   Caption         =   "Ping Form"
   ClientHeight    =   3248
   ClientLeft      =   96
   ClientTop       =   416
   ClientWidth     =   4568
   OleObjectBlob   =   "PingForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
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
    Dim IPアドレス() As String
    IPアドレス() = IPアドレスを取得()
    
    Dim 回数 As Long
    回数 = PingForm.回数.Value
    
    Dim タイムアウト As Long
    タイムアウト = PingForm.タイムアウト.Value
    
    Call Pingを実行してcsvを出力(IPアドレス(), 回数, タイムアウト)
End Function


Private Sub 回数_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call Escキーで終了(KeyCode)
End Sub


Private Sub タイムアウト_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call Escキーで終了(KeyCode)
End Sub


Private Sub StartButton_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Call Escキーで終了(KeyCode)
End Sub


Private Function Escキーで終了(ByVal KeyCode As Long)
    If KeyCode = 27 Then
        Call Unload(PingForm)
    End If
End Function
