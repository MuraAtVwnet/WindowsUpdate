#################################################################
# スケジュール登録
#################################################################
param (
	$StartTime,				# スケジュール開始時刻(00:00)
	[ValidateSet("Full", "Minimum")][string]$WuOption = "Minimum"	# オプション
	)

##########################################################################
# Usage
##########################################################################
function Usage(){
	echo "Usage..."
	echo "    RegistSchedule.ps1 StartTime(99:99) WuDelayDate(9) WuOptin( Minimum | Full )"
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
# スケジュール登録
#######################################################
function EntorySchedule( $FullTaskName, $Script, $RunTime, $Option ){

	SCHTASKS /Create /tn $FullTaskName /tr "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe $Script $Option" /ru "SYSTEM" /sc daily /st $RunTime /f
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
$FullTaskName = "\MURA\Auto Windows Update"

# スクリプト
$Script = "C:\WindowsUpdate\AutoWindowsUpdate.ps1"

# オプション
$Option = [string]$WuOption

# スケジュール登録
EntorySchedule $FullTaskName $Script $StartTime $Option

echo "以下でスケジュール登録しました"
echo "タスク名   : $FullTaskName"
echo "開始時刻   : $StartTime"
echo "オプション : $WuOption"
