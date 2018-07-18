##############################################################
# 古いログを削除する
##############################################################
$LogPath = "C:\WU_Log"
$LogFiles = Join-Path $LogPath "*.log"
$DeleteDay = (Get-Date).AddMonths(-3)

dir $LogFiles | ? {$_.Attributes -notmatch "Directory"} | ? {$_.LastWriteTime -lt $DeleteDay } | del



