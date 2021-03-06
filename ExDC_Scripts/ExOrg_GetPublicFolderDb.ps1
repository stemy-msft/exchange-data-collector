#############################################################################
#                      ExOrg_GetPublicFolderDb.ps1							#
#                                     			 							#
#                               4.0.2    		 							#
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
$ErrorText = "ExOrg_GetPublicFolderDb " + "`n" + $server + "`n"
$ErrorText += $_

$ErrorLog = New-Object -TypeName System.Diagnostics.EventLog -ArgumentList Application
$ErrorLog.MachineName = "."
$ErrorLog.Source = "ExDC"
$ErrorLog.WriteEntry($ErrorText,"Error", 100)
}

set-location -LiteralPath $location
$output_location = $location + "\output\ExOrg\GetPublicFolderDb"

if ((Test-Path -LiteralPath $output_location) -eq $false)
    {New-Item -Path $output_location -ItemType directory -Force}

if ($session_0 -ne $null)
{
	$cxnUri = "http://" + $session_0 + "/powershell"
	$session = New-PSSession -configurationName Microsoft.Exchange -ConnectionUri $cxnUri -authentication Kerberos
	Import-PSSession -Session $session -AllowClobber -CommandName Get-PublicFolderDatabase,Set-AdServerSettings
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

$ExOrg_GetPFDb_outputfile = $output_location + "\" + $server + "~~GetPublicFolderDb.txt"

@(Get-PublicFolderDatabase -status -server $server) | ForEach-Object `
{
	try
    {      
        $PFDBPath = $_.EdbFilePath -replace "\\","\\"
	    $PFDBWMI = Get-WmiObject -Class cim_logicalfile -ComputerName $server -Filter "name='$PFDBPath'" -Property filesize
	    $PFDBSize = $PFDBWMI.filesize
	    $PFDBSize_MB = [Math]::Round($PFDBSize/1mb,2)
	    $PFDBSize_GB = [Math]::Round($PFDBSize/1gb,2)	
       	[Microsoft.Exchange.Data.ByteQuantifiedSize]$WhiteSpace = $_.AvailableNewMailboxSpace
        $WhiteSpaceMB = $WhiteSpace.tomb()
    }
    catch [System.Runtime.InteropServices.COMException]
    {
		$ErrorText = "ExOrg_GetPublicFolderDb`n" + $_ + "`n" 
		$ErrorText += "Error encountered: [System.Runtime.InteropServices.COMException]"
		$ErrorLog = New-Object -TypeName System.Diagnostics.EventLog -ArgumentList Application
		$ErrorLog.MachineName = "."
		$ErrorLog.Source = "ExDC"
		$ErrorLog.WriteEntry($ErrorText,"Error", 100)    
    }
    finally
    {
    	$output_ExOrg_GetPFDb = $server + "`t" + `
		   $_.Alias + "`t" + `
		   $_.EdbFilePath + "`t" + `
		   $_.LogFolderPath + "`t" + `
		   $_.MountAtStartup + "`t" + `
		   $_.FirstInstance + "`t" + `
		   $_.MaxItemSize + "`t" + `
	       $_.ItemRetentionPeriod + "`t" + `
	       $_.ProhibitPostQuota + "`t" + `
	       $_.ReplicationSchedule + "`t" + `
	       $_.IssueWarningQuota + "`t" + `
	       $_.Name + "`t" + `
		   $PFDBSize + "`t" + `
	       $PFDBSize_MB + "`t" + `
	       $PFDBSize_GB + "`t" + `
	       $WhiteSpaceMB + "`t" + `
           $_.AllowFileRestore + "`t" + `
           $_.RetainDeletedItemsUntilBackup + "`t" + `
	       $_.SnapshotLastFullBackup + "`t" + `
	       $_.SnapshotLastIncrementalBackup + "`t" + `
	       $_.SnapshotLastDifferentialBackup + "`t" + `
	       $_.LastFullBackup + "`t" + `
	       $_.LastIncrementalBackup + "`t" + `
	       $_.LastDifferentialBackup + "`t" + `
		   $_.MaintenanceSchedule
        $output_ExOrg_GetPFDb | Out-File -FilePath $ExOrg_GetPFDb_outputfile -append 
    }
}

$a = Get-Process -pid $PID

$EventText = "ExOrg_GetPublicFolderDb " + "`n" + $server + "`n"
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
