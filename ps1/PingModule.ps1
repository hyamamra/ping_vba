$IP_Addresses = @(
    '192.168.128.190'
    '192.168.128.255'
)

$path = 'C:\Users\main\Important_files\Excel_VBA\Ping\Ping_2022-06-14_14-36-43.csv'
foreach ($IP in $IP_Addresses) {
    $value = C:\Users\main\Important_files\Excel_VBA\Ping\ps1\Ping.ps1 $IP 6 1000
    Add-Content -Path $path -Value $value
}

$message = 'Ping.csvファイルの取り込みを開始できます。'
$title = 'Pingが完了しました。'
$icon = 'Info'
powershell.exe -Sta -NoProfile -WindowStyle Hidden -ExecutionPolicy RemoteSigned -File C:\Users\main\Important_files\Excel_VBA\Ping\ps1\ShowBalloon.ps1 $message $title $icon
