Attribute VB_Name = "GetIPAddress"
Option Explicit


Public Function IPアドレスを取得() As String()
    'IP_AddressシートのA列からIPアドレスのみ取得します。
    
    '戻り値
        'IPアドレス() As String
    
    
    Dim IP_AddressSheet As Worksheet
    Set IP_AddressSheet = ThisWorkbook.Worksheets("IP_Address")
    
    Dim 最終行 As Long
    最終行 = 最終行を取得(IP_AddressSheet, 列:=1)
    
    '(行, 1)の二次元配列
    Dim A列の値() As Variant
    With IP_AddressSheet
        A列の値() = .Range(.Range("A1"), .Range("A" & 最終行)).Value
    End With
    
    Dim IPアドレス() As String
    Dim index As Long
    index = 0
    
    Dim 行
    For 行 = 1 To 最終行
        Dim 値 As String
        値 = A列の値(行, 1)
        If IsIPv4アドレス(値) Then
            ReDim Preserve IPアドレス(index)
            IPアドレス(index) = 値
            index = index + 1
        End If
    Next
    
    IPアドレスを取得 = IPアドレス()
End Function


Private Function 最終行を取得(シート As Worksheet, 列 As Long) As Long
    'シートの終端から上方向に値の入ったセルを探索します。
    '値の入ったセルがあれば行番号を返します。
    '値の入ったセルがなければ1を返します。
    
    '引数
        'シート As Worksheet
        '探索するシート
        
        '列 As Long
        '探索する列
    
    
    Dim 終端 As Long
    終端 = シート.Rows.Count
    
    Dim 最終行 As Long
    If シート.Cells(終端, 列) = "" Then
        最終行 = シート.Cells(終端, 列).End(xlUp).Row
    Else
        最終行 = シート.Cells(終端, 列)
    End If
    
    最終行を取得 = 最終行
End Function


Private Function IsIPv4アドレス(ByVal 値) As Boolean
    'IPv4アドレス表記であれば True を返します。
    '10進数表記でなければ False を返します。
    'ドットが3つあり、各桁の数値が 0 ～ 255 であれば
    'IPv4アドレス表記であると判断します。
    
    
    Dim 各桁() As String
    各桁() = Split(CStr(値), ".")
    
    Dim ドットの数 As Long
    ドットの数 = UBound(各桁())
    If Not ドットの数 = 3 Then
        IsIPv4アドレス = False
        Exit Function
    End If
    
    Dim 数値
    For Each 数値 In 各桁()
        If Not ((0 <= 数値) And (数値 <= 255)) Then
            IsIPv4アドレス = False
            Exit Function
        End If
    Next
    
    IsIPv4アドレス = True
End Function
