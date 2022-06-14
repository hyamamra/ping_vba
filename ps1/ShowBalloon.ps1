# バルーン通知を表示します。
# $icon
# 'Error', 'Warning', 'Info', 'None'

param(
    [string]$message = 'message',
    [string]$Title = 'PowerShell',
    [string]$icon = 'Info'
)

[reflection.assembly]::loadwithpartialname('System.Windows.Forms')
# [reflection.assembly]::loadwithpartialname('System.Drawing')
$notify = new-object system.windows.forms.notifyicon
$notify.icon = [System.Drawing.SystemIcons]::Application
$notify.visible = $true
$notify.showballoontip(0, $title, $message, [system.windows.forms.tooltipicon]::$icon)
