[array]$FileList = $null

$Now = Get-Date
$Append = [string]$Now.month + "_" + [string]$now.Day + "_" + `
    [string]$now.year + "_" + [string]$now.hour + "_" + [string]$now.minute `
    + "_" + [string]$now.second

$ExDCLogsFile = ".\ExDC_Events_" + $append + ".txt"
#$ExDCLogs = Get-WinEvent -ProviderName ExDC |fl
$ExDCLogs = Get-EventLog -LogName application -Source exdc | Format-List -Property TimeGenerated,EntryType,Source,EventID,Message
$ExDCLogs | Out-File -FilePath $ExDCLogsFile -Force

$FilesInCurrentFolder = Get-ChildItem
foreach ($a in $FilesInCurrentFolder)
{
	If (($a.name -like ($ExDCLogsFile.replace('.\',''))) -or `
		($a.name -like "ExDC_Step3*") -or `
		($a.name -like "Failed*"))
		{$FileList += [string]$a.fullname}
}

$ZipFilename = (get-location).path + "\ExDCPackagedLogs_" + $append + ".zip" 

if (-not (Test-Path -LiteralPath $ZipFilename)) 
{Set-Content -Path $ZipFilename ("PK" + [char]5 + [char]6 + ("$([char]0)" * 18))} 

$ZipFile = (New-Object -ComObject shell.application).NameSpace($ZipFilename) 
$FileList | ForEach-Object {$zipfile.CopyHere($_)} 
