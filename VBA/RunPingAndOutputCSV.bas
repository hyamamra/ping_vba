Attribute VB_Name = "RunPingAndOutputCSV"
Option Explicit


Public Function Pingを実行してcsvを出力(IPアドレス() As String, ByVal 回数 As Long, ByVal タイムアウト As Long)
    'このブックの同階層に PingModule.ps1 を作成し、実行します。
    '実行結果を yyyy-mm-dd_hh-nn-ss_Ping.csv に出力します。
    
    
    Dim 実行ファイルのパス As String
    実行ファイルのパス = 実行ファイルを作成(IPアドレス(), 回数, タイムアウト)
    
    Call ps1ファイルを実行(実行ファイルのパス, True)
End Function


Private Function 実行ファイルを作成(IPアドレス() As String, ByVal 回数 As Long, ByVal タイムアウト As Long) As String
    '戻り値
        '実行ファイルのパス As String
    
    '参照設定
        'Microsoft Scripting Runtime
    
    
    Dim パス As String
    パス = ThisWorkbook.Path & "\ps1\PingModule.ps1"
    実行ファイルを作成 = パス
    
    Dim FSO As New FileSystemObject
    
    Dim ps1 As Object
    Set ps1 = FSO.OpenTextFile(パス, ForWriting, True, TristateTrue)
    
    Dim コマンド As String
    コマンド = コマンドを作成(IPアドレス(), 回数, タイムアウト)
    
    Call ps1.WriteLine(コマンド)
    Call ps1.Close
    
    Set ps1 = Nothing
    Set FSO = Nothing
End Function


Private Function コマンドを作成(IPアドレス() As String, ByVal 回数 As Long, ByVal タイムアウト As Long) As String
    Dim 日時 As String
    日時 = Format(Now, "yyyy-mm-dd_hh-nn-ss")
    
    Dim PingCSVのパス As String
    PingCSVのパス = ThisWorkbook.Path & "\Ping_" & 日時 & ".csv'"
    
    Dim PingPS1のパス As String
    PingPS1のパス = ThisWorkbook.Path & "\ps1\Ping.ps1"
    
    Dim ShowBalloonPS1のパス As String
    ShowBalloonPS1のパス = ThisWorkbook.Path & "\ps1\ShowBalloon.ps1"
    
    Dim コマンド As String
    コマンド = "$IP_Addresses = @(" & vbNewLine
    
    Dim IP
    For Each IP In IPアドレス()
        コマンド = コマンド & "    '" & IP & "'" & vbNewLine
    Next
    
    コマンド = コマンド & ")" & vbNewLine & vbNewLine _
        & "$path = '" & PingCSVのパス & vbNewLine _
        & "foreach ($IP in $IP_Addresses) {" & vbNewLine _
        & "    $value = " & PingPS1のパス _
        & " $IP " & 回数 & " " & タイムアウト & vbNewLine _
        & "    Add-Content -Path $path -Value $value" & vbNewLine _
        & "}" & vbNewLine & vbNewLine _
        & "$message = 'Ping.csvファイルの取り込みを開始できます。'" & vbNewLine _
        & "$title = 'Pingが完了しました。'" & vbNewLine _
        & "$icon = 'Info'" & vbNewLine _
        & "powershell.exe -Sta -NoProfile -WindowStyle Hidden " _
        & "-ExecutionPolicy RemoteSigned -File " & ShowBalloonPS1のパス _
        & " $message $title $icon"
    
    コマンドを作成 = コマンド
End Function


Private Function ps1ファイルを実行(ByVal コマンド As String, Optional 表示 As Boolean = False, Optional 同期処理 As Boolean = False)
    'PowerShellでコマンドを実行します。
    'https://docs.microsoft.com/ja-jp/powershell/module/microsoft.powershell.core/about/about_powershell_exe?view=powershell-5.1
    
    '引数
        'コマンド As String
        
        '表示 As Boolean = False
        '(省略可) ウィンドウの表示設定
        
        '同期処理 As Boolean = False
        '(省略可) コマンドの終了を待機
    
    '参照設定
        'Windows Script Host Object Model
    
    
    Dim 表示設定 As Long
    If 表示 Then
        表示設定 = 1
    End If
    
    Dim WSH As New IWshRuntimeLibrary.WshShell
    
    '-Sta : シングルスレッドで実行
    '-NoProfile : 開始時にプロファイルを読み込まない
    'WindowStyle : ウインドウの表示方法（最小化）
    '-ExecutionPolicy : 実行ポリシー（ローカルなら署名なしで実行可）
    ps1ファイルを実行 = WSH.Run("powershell.exe -Sta -NoProfile " _
        & "-WindowStyle Minimized -ExecutionPolicy RemoteSigned " _
        & "-Command " & コマンド, 表示設定, 同期処理)
    
    Set WSH = Nothing
End Function
