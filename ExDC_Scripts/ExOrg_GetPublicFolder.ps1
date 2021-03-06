#############################################################################
#                     ExOrg_GetPublicFolder.ps1 		 					#
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
$ErrorText = "ExOrg_GetPublicFolder " + "`n" + $server + "`n"
$ErrorText += $_

$ErrorLog = New-Object -TypeName System.Diagnostics.EventLog -ArgumentList Application
$ErrorLog.MachineName = "."
$ErrorLog.Source = "ExDC"
$ErrorLog.WriteEntry($ErrorText,"Error", 100)
}

set-location -LiteralPath $location
$output_location = $location + "\output\ExOrg\GetPublicFolder"

if ((Test-Path -LiteralPath $output_location) -eq $false)
    {New-Item -Path $output_location -ItemType directory -Force}

if ($session_0 -ne $null)
{
	$cxnUri = "http://" + $server + "/powershell"
	$session = New-PSSession -configurationName Microsoft.Exchange -ConnectionUri $cxnUri -authentication Kerberos
	Import-PSSession -Session $session -AllowClobber -CommandName Get-ExchangeServer,Get-publicfolder,Set-AdServerSettings
    Set-AdServerSettings -ViewEntireForest $true
}
else
{
	Add-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.Admin
    ([Microsoft.Exchange.Data.Directory.AdminSessionADSettings]::Instance).ViewEntireForest = $true
}

#if ((Get-ExchangeServer $session_0).ise15orlater -eq $true)
#{
#	if ((Get-ExchangeServer $server).ise15orlater -eq $true)
#	{
#		Write-Host "Server is Ex2013"
#	}
#	elseif (((Get-ExchangeServer $server).ise15orlater -eq $false) -or ((Get-ExchangeServer $server).ise15orlater -eq $null))
#	{
#		Write-Host "Server is Ex2010"
#	}
#}
$ExOrg_GetPF_outputfile = $output_location + "\" + $server + "~~GetPublicFolder.txt"

@(get-publicfolder -server $server -recurse -Resultsize unlimited) | ForEach-Object `
{	
	$output_ExOrg_GetPF = $_.OriginatingServer.tostring() + "`t" + `
		$_.Name + "`t" + `
		$_.ParentPath + "`t" + `
	    $_.UseDatabaseAgeDefaults + "`t" + `
		$_.AgeLimit + "`t" + `
	    $_.UseDatabaseQuotaDefaults + "`t" + `
		$_.StorageQuota + "`t" + `
	    $_.UseDatabaseReplicationSchedule + "`t" + `
	    $_.UseDatabaseRetentionDefaults  + "`t" + `
	    $_.HasSubFolders + "`t" + `
	    $_.MailEnabled + "`t" + `
	    $_.MaxItemSize + "`t" + `
		$_.Replicas
	$output_ExOrg_GetPF | Out-File -FilePath $ExOrg_GetPF_outputfile -append 
}

$a = Get-Process -pid $PID

$EventText = "ExOrg_GetPublicFolder " + "`n" + $server + "`n"
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
