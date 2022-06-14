# pingを実行します。
# 実行結果に 'TTL=' が含まれていれば疎通成功と判断します。
# 成功なら終了コード = 0
# 失敗なら終了コード = 1
# 出力 = IPアドレス, 終了コード

param(
    [string]$IP_Address,
    [int]$number_of_times = 1,
    [int]$wait_time = 1000
)

$ping = ping $IP_Address -n $number_of_times -w $wait_time

$ping = [string]::join('', $ping)

if ($ping.contains('TTL=')) {
    $status = 0
}
else {
    $status = 1
}

$result = $IP_Address + ',' + $status
Write-Output $result
