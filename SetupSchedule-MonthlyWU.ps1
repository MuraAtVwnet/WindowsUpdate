#################################################################
# スケジュール登録
#################################################################
param (
	$StartTime,				# スケジュール開始時刻(00:00)
	[int]$WuDelay = 7,		# WU 実行ディレイ(WU 日からの経過日)
	[ValidateSet("Full", "Minimum")][string]$WuOption = "Minimum",	# オプション
	[switch]$ConsiderationBU	# Build Update を考慮
	)

##########################################################################
# Usage
##########################################################################
function Usage(){
	echo "Usage..."
	echo "    RegistSchedule.ps1 StartTime(99:99) WuDelayDate(9) WuOptin( Minimum | Full ) [-ConsiderationBU]"
	exit
}

##########################################################################
# 開始時刻の整形
##########################################################################
function SetStartTime( $BackupTime ){
	# 99:99 に桁数そろえる
	$hh = ($BackupTime.Split(":"))[0]
	$mm = ($BackupTime.Split(":"))[1]
	if( $hh.Length -eq 1 ){
		$hh = "0" + $hh
	}
	if( $mm.Length -eq 1 ){
		$mm = "0" + $mm
	}

	$BackupTime = $hh + ":" + $mm

	# 中身確認
	if( $BackupTime -notmatch "^[0-9]{2,2}:[0-9]{2,2}$" ){
		echo "[FAIL] バックアップ時刻が $BackupTerget が正しくない : $BackupTime"
		exit
	}
	else{
		if( ([int]$hh -gt 23) -or ([int]$mm -gt 59) ){
			echo "[FAIL] バックアップ時刻が $BackupTerget が正しくない : $BackupTime"
			exit
		}
		else{
			return $BackupTime
		}
	}
}

#######################################################
# 管理権限で実行されているか確認
#######################################################
function HaveIAdministrativePrivileges(){
	$WindowsPrincipal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
	$IsRoleStatus = $WindowsPrincipal.IsInRole("Administrators")
	return $IsRoleStatus
}

#######################################################
# ログ削除時刻
#######################################################
function GetLogRemoveStart( $WUTime ){
	$LogRemoveDateTime = (([datetime]$WUTime).AddHours(-1))
	$LogRemoveStart = $LogRemoveDateTime.ToString("HH:mm")
	return $LogRemoveStart
}

#######################################################
# スケジュール登録
#######################################################
function EntorySchedule( $FullTaskName, $Script, $RunTime, $Option, $ConsiderationBU = $false ){


	if( $ConsiderationBU ){
		SCHTASKS /Create /tn $FullTaskName /tr "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe $Script $Option -ConsiderationBU" /ru "SYSTEM" /sc daily /st $RunTime /f
	}
	else{
		SCHTASKS /Create /tn $FullTaskName /tr "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe $Script $Option" /ru "SYSTEM" /sc daily /st $RunTime /f
	}
}


##########################################################################
# main
##########################################################################

if( $StartTime -eq $null ){
	Usage
}

if( -not (HaveIAdministrativePrivileges) ){
	echo "[FAIL] 管理権限で実行してください。"
	exit
}

# 開始時刻を整形
$StartTime = SetStartTime $StartTime

# タスク名
$FullTaskName = "\MURA\Go Windows Update"

# スクリプト
$Script = "C:\WindowsUpdate\GoWU.ps1"

# オプション
$Option = [string]$WuDelay + " " + $WuOption

# スケジュール登録
EntorySchedule $FullTaskName $Script $StartTime $Option $ConsiderationBU

echo "以下で Windows Update スケジュールを登録しました"
echo "タスク名   : $FullTaskName"
echo "開始時刻   : $StartTime"
echo "Delay 日数 : $WuDelay"
echo "オプション : $WuOption"
echo "BU 考慮    : $ConsiderationBU"
echo ""

# Log 削除開始時刻
$LogRemoveStart = GetLogRemoveStart $StartTime

# タスク名
$FullTaskName = "\MURA\Auto Windows Update Log Remove"

# スクリプト
$Script = "C:\WindowsUpdate\RemoveLog.ps1"

# ログ削除 スケジュール登録
EntorySchedule $FullTaskName $Script $LogRemoveStart

echo "以下でログ削除スケジュールを登録しました"
echo "タスク名   : $FullTaskName"
echo "開始時刻   : $LogRemoveStart"
