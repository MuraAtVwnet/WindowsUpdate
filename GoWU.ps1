#################################################################
# WU より指定経過日の時に Windows Update スクリプトを投入する
#################################################################
param (
	$LC_WuDelay = 7,			# WU 実行ディレイ(WU 日からの経過日)
	[ValidateSet("Full", "Minimum")][string]$LC_WuOption = "Minimum"	# オプション
	)

# スクリプトの配置場所
$LC_ScriptDir = Split-Path $MyInvocation.MyCommand.Path -Parent

# Windows Update スクリプト
$LC_WuScript = Join-Path $LC_ScriptDir "AutoWindowsUpdate.ps1"

$LC_LogPath = "C:\WU_Log"
$LC_LogName = "GoWU"
##########################################################################
# ログ出力
##########################################################################
function Log(
			$LogString
			){

	$Now = Get-Date

	$Log = $Now.ToString("yyyy/MM/dd HH:mm:ss.fff") + " "
	$Log += $LogString

	if( $LC_LogName -eq $null ){
		$LC_LogName = "LOG"
	}

	$LogFile = $LC_LogName + "_" +$Now.ToString("yyyy-MM") + ".log"

	# ログフォルダーがなかったら作成
	if( -not (Test-Path $LC_LogPath) ) {
		New-Item $LC_LogPath -Type Directory
	}

	$LogFileName = Join-Path $LC_LogPath $LogFile

	Write-Output $Log | Out-File -FilePath $LogFileName -Encoding utf8 -append

	Return $Log
}

###############################################
# Windows Update 日を取得する(日本)
###############################################
function GetWindowsUpdateDay([datetime]$TergetDate){

    # 1日の曜日と US Windows Update 日のオフセット ハッシュテーブル
    $DayOfWeek2WUOffset = @{
        [System.DayOfWeek]"Wednesday"   = 13    # 水曜日
        [System.DayOfWeek]"Thursday"    = 12    # 木曜日
        [System.DayOfWeek]"Friday"      = 11    # 金曜日
        [System.DayOfWeek]"Saturday"    = 10    # 土曜日
        [System.DayOfWeek]"Sunday"      = 9     # 日曜日
        [System.DayOfWeek]"Monday"      = 8     # 月曜日
        [System.DayOfWeek]"Tuesday"     = 7     # 火曜日
    }

    # 年月が指定されていない(default)
    if( $TergetDate -eq $null ){
        # 今の日時
        $TergetDate = Get-Date
    }

    # 1日
    $1stDay = [datetime]$TergetDate.ToString("yyyy/MM/1")

    # US Windows Update 日のオフセット
    $Offset = $DayOfWeek2WUOffset[$1stDay.DayOfWeek]

    if( $Offset -ne $null ){
        # US Windows Update 日
        $WUDayUS = $1stDay.AddDays($Offset)

        # 日本の Windows Update 日(US Windows Update の翌日)
        $WUDay = $WUDayUS.AddDays(1)
    }
    else{
        $WUDay = $null
    }

    return $WUDay
}

###############################################
# Main
###############################################

# 今月の WU 日
$WUDay = GetWindowsUpdateDay

# 今日
$Now = Get-Date
$Today = [datetime]$Now.ToString("yyyy/MM/dd")

# 差
$Diff = New-TimeSpan $WUDay $Today
$DiffDays = $Diff.Days

Log "[INFO] Diff : $DiffDays / Delay $LC_WuDelay"

# 指定経過日
if( $DiffDays -eq $LC_WuDelay ){
	# WU 実施
	Log "[INFO] Go Windows Update : $LC_WuOption"
	. $LC_WuScript $LC_WuOption
}

