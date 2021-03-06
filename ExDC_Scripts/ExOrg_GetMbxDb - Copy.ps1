#############################################################################
#                           ExOrg_GetMbxDb.ps1		 						#
#                                     			 							#
#                                 4.0.1    		 							#
#                                     			 							#
#   This Sample Code is provided for the purpose of illustration only       #
#   and is not intended to be used in a production environment.  THIS       #
#   SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT    #
#   WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT    #
#   LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS     #
#   FOR A PARTICULAR PURPOSE.  We grant You a nonexclusive, royalty-free    #
#   right to use and modify the Sample Code and to reproduce and distribute #
#   the object code form of the Sample Code, provided that You agree:       #
#   (i) to not use Our name, logo, or trademarks to market Your software    #
#   product in which the Sample Code is embedded; (ii) to include a valid   #
#   copyright notice on Your software product in which the Sample Code is   #
#   embedded; and (iii) to indemnify, hold harmless, and defend Us and      #
#   Our suppliers from and against any claims or lawsuits, including        #
#   attorneys' fees, that arise or result from the use or distribution      #
#   of the Sample Code.                                                     #
#                                     			 							#
#############################################################################
Param($server,$location,$append,$session_0)

Write-Output -InputObject $PID
Write-Output -InputObject "ExOrg"

$ErrorActionPreference = "Stop"
Trap {
$ErrorText = "ExOrg_GetMbxDb " + "`n" + $server + "`n"
$ErrorText += $_

$ErrorLog = New-Object -TypeName System.Diagnostics.EventLog -ArgumentList Application
$ErrorLog.MachineName = "."
$ErrorLog.Source = "ExDC"
$ErrorLog.WriteEntry($ErrorText,"Error", 100)
}

set-location -LiteralPath $location
$output_location = $location + "\output\ExOrg\GetMbxDb"

if ((Test-Path -LiteralPath $output_location) -eq $false)
    {New-Item -Path $output_location -ItemType directory -Force}

if ($session_0 -ne $null)
{
    $cxnUri = "http://" + $session_0 + "/powershell"
	$session = New-PSSession -configurationName Microsoft.Exchange -ConnectionUri $cxnUri -authentication Kerberos
	Import-PSSession -Session $session -AllowClobber -CommandName Get-MailboxDatabase,get-mailbox,Set-AdServerSettings
    Set-AdServerSettings -ViewEntireForest $true
	
	if (Test-Path -Path HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup)
    	{$exinstall = (Get-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup).MsiInstallPath}
    else
    	{$exinstall = (Get-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup).MsiInstallPath}
	$dll = $exinstall + "\Bin\Microsoft.Exchange.Data.dll"
	[Reflection.Assembly]::LoadFile($dll)
}
else
{
	Add-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.Admin
    ([Microsoft.Exchange.Data.Directory.AdminSessionADSettings]::Instance).ViewEntireForest = $true
}

$ExOrg_GetMbxDb_outputfile = $output_location + "\" + $server + "~~GetMbxDb.txt"

$gmbxdbparams = (Get-Command -Name Get-MailboxDatabase).Parameters.Keys
if ($gmbxdbparams -contains 'IncludePreExchange2013')
	{$mbxdbs = @(Get-MailboxDatabase -status -server $server -IncludePreExchange2013)}
elseif ($gmbxdbparams -contains 'IncludePreExchange2010')
	{$mbxdbs = @(Get-MailboxDatabase -status -server $server -IncludePreExchange2010)}
else
	{$mbxdbs = @(Get-MailboxDatabase -status -server $server)}

$mbxdbs | ForEach-Object `
{
	$MbxDBSize = $MbxDBSize_MB = $MbxDBSize_GB = $MbxCount = $MbxDBVolume = $null

	try
	{
    	$MbxDBPath = $_.EdbFilePath -replace "\\","\\"
    	$MbxDBWMI = Get-WmiObject -Class cim_logicalfile -ComputerName $server -Filter "name='$MbxDBPath'" -Property filesize
    	$MbxDBSize = $MbxDBWMI.filesize
    	$MbxDBSize_MB = [Math]::Round($MbxDBSize/1mb,2)
    	$MbxDBSize_GB = [Math]::Round($MbxDBSize/1gb,2)	
    	$MbxCount = @(Get-Mailbox -Database $_.guid.guid -ResultSize unlimited).count
        if ($_.AvailableNewMailboxSpace -ne $null)
        {
            [Microsoft.Exchange.Data.ByteQuantifiedSize]$WhiteSpace = $_.AvailableNewMailboxSpace
            $WhiteSpaceMB = $WhiteSpace.tomb()            
            $WhiteSpacePercent = [math]::round($WhiteSpace.ToBytes()*100/$MbxDBSize, 2)
        }
        else
        {
            $WhiteSpaceMB = "Not supported in E2k7"
            $WhiteSpacePercent = "Not supported in E2k7"
        }

        $EdbFilePath = $_.EdbFilePath
        $MbxDBVolume = Get-WmiObject -Class win32_volume -ComputerName $server -Property name, freespace, capacity | `
			Where-Object {$EdbFilePath -like $_.Name + '*'} | Sort-Object -Property name | Select-Object -last 1
	}
	catch [System.Runtime.InteropServices.COMException]
	{
		$ErrorText = "ExOrg_GetMbxDb`n" + $_ + "`n" 
		$ErrorText += "Error encountered: [System.Runtime.InteropServices.COMException]"
		$ErrorLog = New-Object -TypeName System.Diagnostics.EventLog -ArgumentList Application
		$ErrorLog.MachineName = "."
		$ErrorLog.Source = "ExDC"
		$ErrorLog.WriteEntry($ErrorText,"Error", 100)
	}
	finally
	{
		$output_ExOrg_GetMbxDb = $Server + "`t" + `
			$_.Name + "`t" + `
			$_.Recovery + "`t" + `
			$_.EdbFilePath + "`t" + `
		    $_.LogFolderPath + "`t" + `
		    $_.MountAtStartup + "`t" + `
		    $_.OfflineAddressBook + "`t" + `
		    $_.PublicFolderDatabase + "`t" + `
		    $_.IssueWarningQuota + "`t" + `
		    $_.ProhibitSendQuota + "`t" + `
		    $_.ProhibitSendReceiveQuota + "`t" + `
		    $_.DatabaseCopies + "`t" + `
		    $_.ReplayLagTimes + "`t" + `
		    $_.TruncationLagTimes + "`t" + `
		    $_.RPCClientAccessServer + "`t" + `
		    $_.MasterServerOrAvailabilityGroup + "`t" + `
		    $_.MasterType + "`t" + `
		    $_.DataMoveReplicationConstraint + "`t" + `
		    $_.ActivationPreference + "`t" + `
		    $_.MailboxRetention + "`t" + `
		    $_.DeletedItemRetention + "`t" + `
		    $_.BackgroundDatabaseMaintenance + "`t" + `
		    $_.Servers + "`t" + `
		    $_.MountedOnServer + "`t" + `
		    $_.CircularLoggingEnabled + "`t" + `
			$MbxDBSize + "`t" + `
		    $MbxDBSize_MB + "`t" + `
		    $MbxDBSize_GB + "`t" + `
			$WhiteSpaceMB + "`t" + `
			$WhiteSpacePercent + "`t" + `
            $MbxCount + "`t" + `
		    $_.AllowFileRestore + "`t" + `
		    $_.RetainDeletedItemsUntilBackup + "`t" + `
		    $_.SnapshotLastFullBackup + "`t" + `
		    $_.SnapshotLastIncrementalBackup + "`t" + `
		    $_.SnapshotLastDifferentialBackup + "`t" + `
		    $_.LastFullBackup + "`t" + `
		    $_.LastIncrementalBackup + "`t" + `
		    $_.LastDifferentialBackup + "`t" + `
            $MbxDBVolume.Name + "`t" + `
            [math]::round($MbxDBVolume.Capacity / 1MB) + "`t" + `
            [math]::round($MbxDBVolume.Freespace / 1MB) + "`t" + `
            [math]::round($MbxDBVolume.Freespace * 100 / $MbxDBVolume.Capacity) + "`t" + `
		    $_.JournalRecipient + "`t" + `
		    $_.MaintenanceSchedule 
		$output_ExOrg_GetMbxDb | Out-File -FilePath $ExOrg_GetMbxDb_outputfile -append 
	}
}

$a = Get-Process -pid $PID

$EventText = "ExOrg_GetMbxDb " + "`n" + $server + "`n"
$vmMB = [int](($a.vm)/1048576)
$wsMB = [int](($a.ws)/1048576)
$pmMB = [int](($a.privatememorysize)/1048576)
$RunTimeInSec = [int](((get-date) - $a.starttime).totalseconds)
$EventText += "VirtualMemorySize `t" + $vmMB + " MB `n"
$EventText += "WorkingSet      `t`t" + $wsMB + " MB `n"
$EventText += "PrivateMemorySize `t" + $pmMB + " MB `n"
$EventText += "Process run time`t`t" + $RunTimeInSec + " sec `n" 

$EventLog = New-Object -TypeName System.Diagnostics.EventLog -ArgumentList Application
$EventLog.MachineName = "."
$EventLog.Source = "ExDC"
$EventLog.WriteEntry($EventText,"Information", 35)

if ($session_0 -ne $null)
{
	Remove-PSSession -Session $session
}
