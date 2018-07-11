<#
.SYNOPSIS
Windows Update スクリプトを投入し再起動します
PowerShell プロンプトを閉じないように注意する必要がありません

<CommonParameters> はサポートしていません

.DESCRIPTION
・フルアップデート(-Option Full)
    全ての更新プログラムを適用します

・ミニマムアップデート(-Option Minimum)
    重要な更新プログラムのみを適用します
    オプションを省略した場合はこのモードになります

.EXAMPLE
PS C:\WindowsUpdate> .\SubmitWU.ps1 Full

全ての更新プログラムを適用

.EXAMPLE
PS C:\WindowsUpdate> .\SubmitWU.ps1

重要な更新プログラムのみを適用

.PARAMETER Option
操作モード
    Full: 全ての更新プログラムを適用
    Minimum: 重要な更新プログラムのみを適用
    省略: 重要な更新プログラムのみを適用

<CommonParameters> はサポートしていません

.LINK
http://www.vwnet.jp/Windows/PowerShell/FullAutoWU.htm
#>


##########################################################################
#
# Windows Update PowerShell 投入
#
#   参考サイト: http://yamanxworld.blogspot.jp/2010/07/windows-scripting-windows-update_05.html
#
##########################################################################
param (
        [ValidateSet("Full", "Minimum")][string]$Option # アップデートオプション
    )

# スクリプトの配置場所
$G_MyName = "C:\WindowsUpdate\AutoWindowsUpdate.ps1"

# 最大適用更新数
$G_MaxUpdateNumber = 100

# スクリプトの配置場所
$C_ScriptDir = Split-Path $MyInvocation.MyCommand.Path -Parent

# 完了時刻記録ファイル
$G_SetTimeStampFilePath = "C:\WindowsUpdate"
$G_SetTimeStampFileName = "WU_TimeStamp.txt"

# リラン判定日数
$C_ReRunDate = 5

$G_LogPath = "C:\WU_Log"
$G_LogName = "WU_Log.txt"
##########################################################################
# ログ出力
##########################################################################
function Log(
            $LogString
            ){

    $Now = Get-Date

    $Log = "{0:0000}-{1:00}-{2:00} " -f $Now.Year, $Now.Month, $Now.Day
    $Log += "{0:00}:{1:00}:{2:00}.{3:000} " -f $Now.Hour, $Now.Minute, $Now.Second, $Now.Millisecond
    $Log += $LogString

    if( $G_LogName -eq $null ){
        $G_LogName = "LOG"
    }

    $LogFile = $G_LogName +"_{0:0000}-{1:00}-{2:00}.log" -f $Now.Year, $Now.Month, $Now.Day

    # ログフォルダーがなかったら作成
    if( -not (Test-Path $G_LogPath) ) {
        New-Item $G_LogPath -Type Directory
    }

    $LogFileName = Join-Path $G_LogPath $LogFile

    Write-Output $Log | Out-File -FilePath $LogFileName -Encoding Default -append

    Return $Log
}

##########################################################################
# レジストリ追加/更新
##########################################################################
function RegSet( $RegPath, $RegKey, $RegKeyType, $RegKeyValue ){
    # レジストリそのものの有無確認
    $Elements = $RegPath -split "\\"
    $RegPath = ""
    $FirstLoop = $True
    foreach ($Element in $Elements ){
        if($FirstLoop){
            $FirstLoop = $False
        }
        else{
            $RegPath += "\"
        }
        $RegPath += $Element
        if( -not (test-path $RegPath) ){
            Log "Add Registry : $RegPath"
            md $RegPath
        }
    }

    # Key有無確認
    $Result = Get-ItemProperty $RegPath -name $RegKey -ErrorAction SilentlyContinue
    # キーがあった時
    if( $Result -ne $null ){
        Set-ItemProperty $RegPath -name $RegKey -Value $RegKeyValue
    }
    # キーが無かった時
    else{
        # キーを追加する
        New-ItemProperty $RegPath -name $RegKey -PropertyType $RegKeyType -Value $RegKeyValue
    }
    Get-ItemProperty $RegPath -name $RegKey
}

##########################################################################
# Autoexec.ps1 有効
##########################################################################
function EnableAutoexec( $ScriptName, $Option ){
    if( -not (Test-Path $ScriptName)){
        Log "[FAIL] $ScriptName not found !!"
        exit
    }

    ### boot 時に autoexec.ps1 を自動実行するレジストリ設定

    $RegPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\Scripts\Startup\0"
    $RegKey = "FileSysPath"
    $RegKeyType = "String"
    $RegKeyValue = "C:\\Windows\\System32\\GroupPolicy\\Machine"
    RegSet $RegPath $RegKey $RegKeyType $RegKeyValue

    $RegKey = "PSScriptOrder"
    $RegKeyType = "DWord"
    $RegKeyValue = 3
    RegSet $RegPath $RegKey $RegKeyType $RegKeyValue

    $RegPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\Scripts\Startup\0\0"
    $RegKey = "Script"
    $RegKeyType = "String"
    $RegKeyValue = "autoexec.ps1"
    RegSet $RegPath $RegKey $RegKeyType $RegKeyValue

    $RegKey = "Parameters"
    $RegKeyType = "String"
    $RegKeyValue = "$Option"
    RegSet $RegPath $RegKey $RegKeyType $RegKeyValue

    $RegKey = "IsPowershell"
    $RegKeyType = "DWord"
    $RegKeyValue = 1
    RegSet $RegPath $RegKey $RegKeyType $RegKeyValue

    $RegPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\State\Machine\Scripts\Startup\0"
    $RegKey = "FileSysPath"
    $RegKeyType = "String"
    $RegKeyValue = "C:\\Windows\\System32\\GroupPolicy\\Machine"
    RegSet $RegPath $RegKey $RegKeyType $RegKeyValue

    $RegKey = "PSScriptOrder"
    $RegKeyType = "DWord"
    $RegKeyValue = 3
    RegSet $RegPath $RegKey $RegKeyType $RegKeyValue

    $RegPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\State\Machine\Scripts\Startup\0\0"
    $RegKey = "Script"
    $RegKeyType = "String"
    $RegKeyValue = "autoexec.ps1"
    RegSet $RegPath $RegKey $RegKeyType $RegKeyValue

    $RegKey = "Parameters"
    $RegKeyType = "String"
    $RegKeyValue = "$Option"
    RegSet $RegPath $RegKey $RegKeyType $RegKeyValue

    ### 自動実行するスクリプトを autoexec.ps1 に上書きコピー
    $TergetPath = "C:\Windows\System32\GroupPolicy\Machine\Scripts\Startup"
    $TergetFile = Join-Path $TergetPath "autoexec.ps1"
    Log "$ScriptName → $TergetFile"
    if( -not (Test-Path $TergetPath)){
        md $TergetPath
        Log "md $TergetPath"
    }
    copy $ScriptName $TergetFile -Force

    Log "[INFO] $ScriptName → $TergetFile copied"

    if( -not (Test-Path $TergetFile)){
        Log "[FAIL] autoexec.ps1 copy failed !!"
        exit
    }
}

##########################################################################
# Autoexec.ps1 無効
##########################################################################
function DisableAutoexec(){
    $RegPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\Scripts\Startup\0"
    if(Test-Path $RegPath){
        Remove-Item $RegPath -Recurse -Force -Confirm:$false
    }

    $RegPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\State\Machine\Scripts\Startup\0"
    if(Test-Path $RegPath){
        Remove-Item $RegPath -Recurse -Force -Confirm:$false
    }

    $Terget = "c:\Windows\System32\GroupPolicy\Machine\Scripts\Startup\autoexec.ps1"
    if(Test-Path $Terget){
        del $Terget -Force
    }
    Log "[INFO] Autoexec Disabled."
}

##########################################################################
# タイムスタンプファイル読み込み
##########################################################################
function GetTimeStampFile($SetTimeStampFilePath, $SetTimeStampFileName){

    $SetTimeStampFileFullName = Join-Path $SetTimeStampFilePath $SetTimeStampFileName

    if( -not (Test-Path $SetTimeStampFilePath)){
        # Path が存在しない
        return $null
    }

    if( -not (Test-Path $SetTimeStampFileFullName)){
        # ファイルが存在しない
        return $null
    }

    [datetime]$FinishTime = Get-Content -Path $SetTimeStampFileFullName -Encoding UTF8
    return $FinishTime
}

##################################################
# リラン判定
##################################################
function IsReRun(){

	$FinishTime = GetTimeStampFile $G_SetTimeStampFilePath $G_SetTimeStampFileName
	# 存在確認
	if( $FinishTime -ne $null ){
		# WU 完了日が新しい時は rerun と判断
		if( $FinishTime -gt (Get-Date).AddDays( -$C_ReRunDate ) ){
			Return $true
		}
	}

	Return $false
}

##########################################################################
#
# main
#
##########################################################################
if (-not(([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))) {
    Log "実行には管理権限が必要です"
    exit
}

# 自動起動スクリプト停止
DisableAutoexec

$OSData = Get-WmiObject Win32_OperatingSystem
$BuildNumber = $OSData.BuildNumber
$strVersion = $OSData.Version
$strVersion = $strVersion.Replace( ".$BuildNumber", "" )
$Version = [decimal]$strVersion
if( $Version -lt 6.1 ){
    Log "Windows Server 2008 R2 / Windows 7 以降しかサポートしていません"
    exit
}

# スクリプト存在確認
if( -not (Test-Path $G_MyName)){
    Log "[FAIL] $G_MyName not found !!"
    exit
}

# リラン判定
if( IsReRun ){
	Log "[INFO] リランなので Windows Update スキップ"
	exit
}

Log "--- Submit Windows Update ---"

# アップデートタイプコントロール
if( $Option -match "ful" ){
    Log "Full Update"
    $Option = "Full"
}
else{
    Log "Minimum Update"
    $Option = "Minimum"
}

# 既知の問題(KB2962824)対応
if(($Version -ge 6.3) -and ($Version -lt 6.4)){
    $OSData = Get-WmiObject Win32_OperatingSystem
    if( $OSData.Caption -match "Windows Server 2012 R2" ){
        Log "Windows Server 2012 R2"
        $OriginalProgressPreference = $ProgressPreference
        $ProgressPreference="SilentlyContinue"
        $BitLocker = Get-WindowsFeature BitLocker
        if( $BitLocker.Installed -eq $false ){
            Log "BitLocker がインストールされていないのでインストール & 再起動"
            Add-WindowsFeature BitLocker
        }
        $ProgressPreference = $OriginalProgressPreference
    }
}

EnableAutoexec $G_MyName $Option
sleep 30
Log "Reboot system now !!"
Restart-Computer -Force

