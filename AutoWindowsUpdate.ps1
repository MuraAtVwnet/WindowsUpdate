<#
.SYNOPSIS
Windows Update と再起動を自動実行します

<CommonParameters> はサポートしていません

.DESCRIPTION
・フルアップデート(-Option Full)
    全ての更新プログラムを適用します

・ミニマムアップデート(-Option Minimum)
    重要な更新プログラムのみを適用します
    オプションを省略した場合はこのモードになります

・Build Update 時の boot loop 抑制(-ConsiderationBU)
    Build Update 時に boot loop に陥らないようにします
    Windows 10 で Windows Update を定期スケジュールを組む場合に指定します。
    初期構築時の Windows Update では使用しません
    (3 時間以内に起動/再起動されている場合はスクリプトを実行しません)

・Microsoft Teams へのメッセージ送信テスト(-MessageTest)
    メッセージ送信だけをします
    (Windows Update はしません)

.EXAMPLE
PS C:\WindowsUpdate> .\AutoWindowsUpdate.ps1 Full

全ての更新プログラムを適用

.EXAMPLE
PS C:\WindowsUpdate> .\AutoWindowsUpdate.ps1

重要な更新プログラムのみを適用

.EXAMPLE
PS C:\WindowsUpdate> .\AutoWindowsUpdate.ps1 Full -ConsiderationBU

Build Update を考慮した全ての更新プログラムを適用

.EXAMPLE
PS C:\WindowsUpdate> .\AutoWindowsUpdate.ps1 -MessageTest

Microsoft Teams へのメッセージ送信テスト

.PARAMETER Option
操作モード
    Full    : 全ての更新プログラムを適用
    Minimum : 重要な更新プログラムのみを適用
    省略    : 重要な更新プログラムのみを適用

.PARAMETER ConsiderationBU
Build Update を考慮
    指定: Build Update 時に boot loop に陥らないようにする
    省略: Build Update を考慮せずに Windows Update する

.PARAMETER MessageTest
Microsoft Teams へのメッセージ送信テスト
    指定: メッセージ送信テストだけを実行する
    省略: 通常処理

<CommonParameters> はサポートしていません

.LINK
http://www.vwnet.jp/Windows/PowerShell/FullAutoWU.htm
#>

##########################################################################
#
# Windows Update PowerShell
#
#  C:\WindowsUpdate にスクリプトを置いて実行する
#
#	参考サイト: http://yamanxworld.blogspot.jp/2010/07/windows-scripting-windows-update_05.html
#
##########################################################################
param (
		[ValidateSet("Full", "Minimum")][string]$Option,	# アップデートオプション
		[switch]$ConsiderationBU,							# Build Update を考慮
		[switch]$MessageTest								# Microsoft Teams メッセージ送信テスト
	)

# スクリプトの配置場所
$G_MyName = "C:\WindowsUpdate\AutoWindowsUpdate.ps1"

# 完了時刻記録ファイル
$G_SetTimeStampFilePath = "C:\WindowsUpdate"
$G_CompleteTimeStampFileName = "WU_TimeStamp.txt"

# 再起動時刻記録ファイル
$G_RebootTimeStampFileName = "Reboot_TimeStamp.txt"

# Build Vertion 記録ファイル
$G_BuildVertionFileName = "Build_Vertion.txt"

# 最大適用更新数
$G_MaxUpdateNumber = 100

# 再起動禁止時間
$G_BootProhibitionTime = 3

# Microsoft Teams メッセージ送信用 URI ファイル名
$G_MicrosoftTeamsUriFileName = "MST_URI.txt"

$G_LogPath = "C:\WU_Log"
$G_LogName = "WU_Log.txt"
##########################################################################
# ログ出力
##########################################################################
function Log(
			$LogString
			){

	$Now = Get-Date

	$Log = $Now.ToString("yyyy/MM/dd HH:mm:ss.fff") + " "
	$Log += $LogString

	if( $G_LogName -eq $null ){
		$G_LogName = "LOG"
	}

	$LogFile = $G_LogName + "_" +$Now.ToString("yyyy-MM-dd") + ".log"

	# ログフォルダーがなかったら作成
	if( -not (Test-Path $G_LogPath) ) {
		New-Item $G_LogPath -Type Directory
	}

	$LogFileName = Join-Path $G_LogPath $LogFile

	Write-Output $Log | Out-File -FilePath $LogFileName -Encoding utf8 -append

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
function EnableAutoexec( $ScriptName, $Option, $ConsiderationBU ){
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
	if( $ConsiderationBU ){
		$RegKeyValue = "$Option -ConsiderationBU"
	}
	else{
		$RegKeyValue = "$Option"
	}
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
	if( $ConsiderationBU ){
		$RegKeyValue = "$Option -ConsiderationBU"
	}
	else{
		$RegKeyValue = "$Option"
	}
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
# WU で再起動されたかの確認
##########################################################################
function IsWURebooted(){

	# 前回再起動時刻
	$RebootTime = GetTimeStampFile $G_SetTimeStampFilePath $G_RebootTimeStampFileName
	if( $RebootTime -eq $null ){
		# 一度も再起動していない
		return $false
	}

	$TimeSpan = New-TimeSpan $RebootTime (Get-Date)
	[int]$TotalHours = $TimeSpan.TotalHours
	if( $TotalHours -le $G_BootProhibitionTime ){
		return $true
	}
	else{
		return $false
	}
}


##########################################################################
# タイムスタンプファイル出力
##########################################################################
function SetTimeStampFile($SetTimeStampFilePath, $SetTimeStampFileName){

	$SetTimeStampFileFullName = Join-Path $SetTimeStampFilePath $SetTimeStampFileName

	if( -not (Test-Path $SetTimeStampFilePath)){
		md $SetTimeStampFilePath
	}

	if( -not (Test-Path $SetTimeStampFilePath)){
		Log "[FAIL] !!!!!!!! $SetTimeStampFilePath not created. !!!!!!!!"
		exit
	}

	$NowTime = Get-Date

	try{
		Set-Content -Value $NowTime.DateTime -Path $SetTimeStampFileFullName  -Encoding UTF8
	}
	catch{
		Log "[FAIL] !!!!!!!! $SetTimeStampFileFullName not created. !!!!!!!!"
		exit
	}
}

##########################################################################
# タイムスタンプファイル読み込み
##########################################################################
function GetTimeStampFile($SetTimeStampFilePath, $SetTimeStampFileName){

	$SetTimeStampFileFullName = Join-Path $SetTimeStampFilePath $SetTimeStampFileName

	if( -not (Test-Path $SetTimeStampFileFullName)){
		return $null
	}

	try{
		[datetime]$GetTime = Get-Content -Path $SetTimeStampFileFullName
	}
	catch{
		Log "[FAIL] !!!!!!!! $SetTimeStampFileFullName read error. !!!!!!!!"
		exit
	}

	return $GetTime
}


##########################################################################
# タイムスタンプファイル削除
##########################################################################
function RemoveTimeStampFile($SetTimeStampFilePath, $SetTimeStampFileName){

	$SetTimeStampFileFullName = Join-Path $SetTimeStampFilePath $SetTimeStampFileName

	if( -not (Test-Path $SetTimeStampFilePath)){
		# Path が存在しない
		return
	}

	if( -not (Test-Path $SetTimeStampFileFullName)){
		# ファイルが存在しない
		return
	}

	del $SetTimeStampFileFullName
	return
}

##########################################################
# Windows Update Reboot を Teames に通知する
##########################################################
function NoticeWU($FilePath, $FileName){

	$FileFullPath = Join-Path $FilePath $FileName

	if( -not (Test-Path $FileFullPath)){
		# URI ファイルが無い時は何もしない
		Log "URI file not found : $FileFullPath"
		return
	}

	# Web API の URL
	[array]$Lines = Get-Content -Path $FileFullPath
	if( $Lines.Count -eq 0 ){
		# データが入っていない
		Log "URI file is empty : $FileFullPath"
		return
	}
	$url = $Lines[0]
	if( $url.Length -le 35 ){
		# URIが短すぎ
		Log "URI data is empty : $FileFullPath"
		return
	}

	# Invoke-RestMethod に渡す Web API の引数を JSON にする
	$HostName = hostname

	$body = ConvertTo-JSON @{
		text = "Windows Update reboot now ! : $HostName"
	}

	# API を叩く
	Invoke-RestMethod -Method Post -Uri $url -Body $body -ContentType 'application/json'
}

##########################################################################
# バージョンファイル出力
##########################################################################
function SetVersionFile($SetVersionFilePath, $SetVersionFileName, $BuildVertion ){

	$SetVersionFileFullName = Join-Path $SetVersionFilePath $SetVersionFileName

	if( -not (Test-Path $SetVersionFilePath)){
		md $SetVersionFilePath
	}

	if( -not (Test-Path $SetVersionFilePath)){
		Log "[FAIL] !!!!!!!! $SetVersionFilePath not created. !!!!!!!!"
		exit
	}

	$RegistryBuildNumber = $BuildVertion.RegistryBuildNumber
	$WinverBuildNumber = $BuildVertion.WinverBuildNumber
	$OSVertion = $BuildVertion.OSVertion

	$VertionDatas = @()

	if($RegistryBuildNumber -ne $null){
		$VertionDatas += $RegistryBuildNumber
	}
	else{
		return
	}

	if( $WinverBuildNumber -ne $null ){
		$VertionDatas += $WinverBuildNumber
	}

	if( $OSVertion -ne $null ){
		$VertionDatas += $OSVertion
	}

	try{
		Set-Content -Value $VertionDatas -Path $SetVersionFileFullName  -Encoding UTF8
	}
	catch{
		Log "[FAIL] !!!!!!!! $SetVersionFileFullName not created. !!!!!!!!"
		exit
	}
}


##########################################################################
# バージョンファイル読み込み
##########################################################################
function GetVersionFile($SetVersionFilePath, $SetVersionFileName){

	$SetVersionFileFullName = Join-Path $SetVersionFilePath $SetVersionFileName

	if( -not (Test-Path $SetVersionFileFullName)){
		return $null
	}

	try{
		[array]$VertionDatas = Get-Content -Path $SetVersionFileFullName
	}
	catch{
		Log "[FAIL] !!!!!!!! $SetVersionFileFullName read error. !!!!!!!!"
		exit
	}

	$ReturnData = New-Object PSObject | Select-Object RegistryBuildNumber, WinverBuildNumber, OSVertion

	if( $VertionDatas.Count -eq 0 ){
		return $null
	}

	if( $VertionDatas.Count -ge 1 ){
		$ReturnData.RegistryBuildNumber = $VertionDatas[0]
	}

	if( $VertionDatas.Count -ge 2 ){
		$ReturnData.WinverBuildNumber = $VertionDatas[1]
	}

	if( $VertionDatas.Count -ge 3 ){
		$ReturnData.OSVertion = $VertionDatas[2]
	}

	return $ReturnData
}


##########################################################################
# Build バージョン 取得
##########################################################################
function GetBuildVersion(){
	$ReturnData = New-Object PSObject | Select-Object RegistryBuildNumber, WinverBuildNumber, OSVertion, Edition

	# ビルド番号詳細
	$RegPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion"
	$RegKey = "BuildLabEx"
	$RegistryBuildNumber = (Get-ItemProperty $RegPath -name $RegKey).$RegKey
	if( $RegistryBuildNumber -ne $null ){
		$ReturnData.RegistryBuildNumber = $RegistryBuildNumber
	}
	else{
		$ReturnData.RegistryBuildNumber = $null
	}

	# Winver のビルド番号
	$RegPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion"
	$RegKey = "CurrentBuild"
	$MajorNumber = (Get-ItemProperty -Path $RegPath -name $RegKey).$RegKey
	$RegKey = "UBR"
	$MinorNumber = (Get-ItemProperty -Path $RegPath -name $RegKey).$RegKey
	$WinverBuildNumber = $MajorNumber + "." + [String]$MinorNumber
	if( $MajorNumber -ne $null ){
		$ReturnData.WinverBuildNumber = $WinverBuildNumber
	}
	else{
		$ReturnData.WinverBuildNumber = $null
	}

	# Winver のバージョン
	$RegPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion"
	$RegKey = "ReleaseId"
	$OSVertion = (Get-ItemProperty $RegPath -name $RegKey -ErrorAction SilentlyContinue).$RegKey
	if( $OSVertion -ne $null ){
		$ReturnData.OSVertion = $OSVertion
	}
	else{
		$ReturnData.OSVertion = $null
	}

	# OS のエディション
	$Win32_OperatingSystem = Get-WmiObject Win32_OperatingSystem
	$OS = $Win32_OperatingSystem.Caption
	$SP = $Win32_OperatingSystem.ServicePackMajorVersion
	if( $SP -ne 0 ){ $OS += " SP" + $SP }
	$ReturnData.Edition = $OS

	return $ReturnData
}

##########################################################
# Windows Update 完了 を Teames に通知する
##########################################################
function NoticeFinishWU($FilePath, $FileName, $BuildVertion){

	$FileFullPath = Join-Path $FilePath $FileName

	if( -not (Test-Path $FileFullPath)){
		# URI ファイルが無い時は何もしない
		Log "URI file not found : $FileFullPath"
		return
	}

	# Web API の URL
	[array]$Lines = Get-Content -Path $FileFullPath
	if( $Lines.Count -eq 0 ){
		# データが入っていない
		Log "URI file is empty : $FileFullPath"
		return
	}
	$url = $Lines[0]
	if( $url.Length -le 35 ){
		# URIが短すぎ
		Log "URI data is empty : $FileFullPath"
		return
	}

	# Invoke-RestMethod に渡す Web API の引数を JSON にする
	$HostName = hostname
	$RegistryBuildNumber = $BuildVertion.RegistryBuildNumber
	$WinverBuildNumber = $BuildVertion.WinverBuildNumber
	$OSVertion = $BuildVertion.OSVertion
	$OSEdition = $BuildVertion.Edition

	$Message = "Windows Update finish : $HostName`n`r"
	$Message += "Registry Build Number : $RegistryBuildNumber`n`r"
	$Message += "Winver Build Number : $WinverBuildNumber`n`r"
	$Message += "OS Vertion : $OSVertion`n`r"
	$Message += "OS Edition : $OSEdition"

	$body = ConvertTo-JSON @{
		text = $Message
	}

	# API を叩く
	Invoke-RestMethod -Method Post -Uri $url -Body $body -ContentType 'application/json'
}

##########################################################
# Windows Update 完了
##########################################################
function FinishWU(){
	# バージョンファイル読み込み
	$LastVertion = GetVersionFile $G_SetTimeStampFilePath $G_BuildVertionFileName

	# Build バージョン 取得
	$NowVertion = GetBuildVersion

	# Build バージョン更新確認
	if( ($LastVertion.RegistryBuildNumber -eq $NowVertion.RegistryBuildNumber) -and
		($LastVertion.WinverBuildNumber   -eq $NowVertion.WinverBuildNumber) -and
		($LastVertion.OSVertion           -eq $NowVertion.OSVertion)){

		# バージョン更新なし
		return
	}
	else{
		# バージョン更新あり

		# バージョンファイル出力
		SetVersionFile $G_SetTimeStampFilePath $G_BuildVertionFileName $NowVertion

		# Windows Update 完了 を Teames に通知する
		NoticeFinishWU $G_SetTimeStampFilePath $G_MicrosoftTeamsUriFileName $NowVertion

		# バージョン情報をログに出力
		$RegistryBuildNumber = $NowVertion.RegistryBuildNumber
		$WinverBuildNumber = $NowVertion.WinverBuildNumber
		$OSVertion = $NowVertion.OSVertion

		Log $RegistryBuildNumber
		Log $WinverBuildNumber
		Log $OSVertion
	}
}


##########################################################################
#
# main
#
##########################################################################
if (-not(([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))) {
	Log "Administrative privileges are required to run"
	exit
}

# Microsoft Teams メッセージ送信テスト
if( $MessageTest ){
	Log "Message send test"
	NoticeWU $G_SetTimeStampFilePath $G_MicrosoftTeamsUriFileName
	exit
}

# 自動起動スクリプト停止
DisableAutoexec

# バージョン確認
$OSData = Get-WmiObject Win32_OperatingSystem
$BuildNumber = $OSData.BuildNumber
$strVersion = $OSData.Version
$strVersion = $strVersion.Replace( ".$BuildNumber", "" )
$Version = [decimal]$strVersion
if( $Version -lt 6.1 ){
	Log "Supports Windows Server 2008 R2/Windows 7 or later"
	exit
}

# スクリプト存在確認
if( -not(Test-Path $G_MyName )){
	Log "Put the script in $G _myname and run it"
	exit
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
			Log "Install & Restart because BitLocker is not installed"
			EnableAutoexec $G_MyName $Option $ConsiderationBU
			Add-WindowsFeature BitLocker -Restart
		}
		$ProgressPreference = $OriginalProgressPreference
	}
}

Log "--- Running Windows Update ---"

# 完了タイムスタンプファイル削除
RemoveTimeStampFile $G_SetTimeStampFilePath $G_CompleteTimeStampFileName

Log "Searching for updates..."
$updateSession = new-object -com "Microsoft.Update.Session"

$updateSearcher = $updateSession.CreateupdateSearcher()

# アップデートタイプコントロール
if( $Option -match "ful" ){
	Log "Full Update"
	$Option = "Full"
	$searchResult = $updateSearcher.Search("IsInstalled=0 and Type='Software'")
}
else{
	Log "Minimum Update"
	$Option = "Minimum"
	$searchResult = $updateSearcher.Search("IsInstalled=0 and Type='Software' and AutoSelectOnWebSites=1")
}


Log "List of applicable items on the machine:"
if ($searchResult.Updates.Count -eq 0) {
	Log "There are no applicable updates."
	SetTimeStampFile $G_SetTimeStampFilePath $G_CompleteTimeStampFileName

	FinishWU
	Log "=-=-=-=-=- Windows Update finished -=-=-=-=-="
}
else{
	$downloadReq = $False
	$i = 0
	foreach ($update in $searchResult.Updates){
		$i++
		if ( $update.IsDownloaded ) {
			$UpdateTitol = $update.Title
			Log "$i : $UpdateTitol (downloaded)"
		}
		else
		{
			$downloadReq = $true
			$UpdateTitol = $update.Title
			Log "$i : $UpdateTitol (not downloaded)"
		}
	}
	if ( $downloadReq ) {
		Log "Creating collection of updates to download..."
		$updatesToDownload = new-object -com "Microsoft.Update.UpdateColl"
		foreach ($update in $searchResult.Updates){
			$updatesToDownload.Add($update) | out-null
		}
		Log "Downloading updates..."
		$downloader = $updateSession.CreateUpdateDownloader()
		$downloader.Updates = $updatesToDownload
		$downloader.Download()
		Log "List of downloaded updates:"
		$i = 0
		foreach ($update in $searchResult.Updates){
			$i++
			if ( $update.IsDownloaded ) {
				$UpdateTitol = $update.Title
				Log "$i : $UpdateTitol (downloaded)"
			}
			else
			{
				$UpdateTitol = $update.Title
				Log "$i : $UpdateTitol (not downloaded)"
			}
		}
	}
	else
	{
		Log "All updates are already downloaded."
	}
	$updatesToInstall = new-object -com "Microsoft.Update.UpdateColl"
	Log "Creating collection of downloaded updates to install..."
	$i = 0
	foreach ($update in $searchResult.Updates){
		if ( $update.IsDownloaded ) {
			$updatesToInstall.Add($update) | out-null
			$i++
			$UpdateTitol = $update.Title
			Log "$i / $G_MaxUpdateNumber : $UpdateTitol (Install)"
			if( $i -ge $G_MaxUpdateNumber ){
				Log "Break max update $G_MaxUpdateNumber"
				break
			}
		}
	}
	if ( $updatesToInstall.Count -eq 0 ) {
		Log "Not ready for installation."
		Log "=-=-=-=-=- Windows Update Abnormal End -=-=-=-=-="
	}
	else
	{
		$InstallCount = $updatesToInstall.Count
		Log "Installing $InstallCount updates..."
		$installer = $updateSession.CreateUpdateInstaller()
		$installer.Updates = $updatesToInstall
		$installationResult = $installer.Install()
		if ( $installationResult.ResultCode -eq 2 ) {
			Log "All updates installed successfully."
		}
		else
		{
			Log "Some updates could not installed."
		}
		if ( $installationResult.RebootRequired ) {
			# Build Update 考慮
			if( $ConsiderationBU ){
				Log "Build Update Consideration decision"
				$Status = IsWURebooted	# 指定時間内に WU 再起動さたか
				if( $Status -eq $true ){
					Log "For Build Updae Consideration, Windows does not Update if uptime is shorter than $G_BootProhibitionTime h : $TotalHours h"
					Log "=-=-=-=-=- Windows Update Abort -=-=-=-=-="
					exit
				}
				else{
					Log "Windows Update continues because it is not Build update"
				}
			}

			# WU 再起動
			Log "One or more updates are requiring reboot."
			Log "[INFO] Autoexec Enabled"
			EnableAutoexec $G_MyName $Option $ConsiderationBU
			sleep 30
			Log "Reboot system now !!"
			SetTimeStampFile $G_SetTimeStampFilePath $G_RebootTimeStampFileName
			NoticeWU $G_SetTimeStampFilePath $G_MicrosoftTeamsUriFileName
			Restart-Computer -Force
		}
		else
		{
			Log "Finished. Reboot are not required."
			SetTimeStampFile $G_SetTimeStampFilePath $G_CompleteTimeStampFileName

			FinishWU
			Log "=-=-=-=-=- Windows Update finished -=-=-=-=-="
		}
	}
}
