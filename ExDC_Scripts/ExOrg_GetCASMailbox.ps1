#############################################################################
#                        ExOrg_GetCASMailbox.ps1		 					#
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
Param($location,$server,$i,$PSSession)

Write-Output -InputObject $PID
Write-Output -InputObject "ExOrg"

$ErrorActionPreference = "Stop"
Trap {
$ErrorText = "ExOrg_GetCASMailbox " + "`n" + $server + "`n"
$ErrorText += $_

$ErrorLog = New-Object -TypeName System.Diagnostics.EventLog -ArgumentList Application
$ErrorLog.MachineName = "."
$ErrorLog.Source = "ExDC"
$ErrorLog.WriteEntry($ErrorText,"Error", 100)
}

set-location -LiteralPath $location
$output_location = $location + "\output\ExOrg\GetCASMailbox"

if ((Test-Path -LiteralPath $output_location) -eq $false)
    {New-Item -Path $output_location -ItemType directory -Force}

if ($PSSession -ne $null)
{
	$cxnUri = "http://" + $PSSession + "/powershell"
	$session = New-PSSession -configurationName Microsoft.Exchange -ConnectionUri $cxnUri -authentication Kerberos
	Import-PSSession -Session $session -AllowClobber -CommandName Get-CASMailbox,Set-AdServerSettings
    Set-AdServerSettings -ViewEntireForest $true
}
else
{
	Add-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.Admin
    ([Microsoft.Exchange.Data.Directory.AdminSessionADSettings]::Instance).ViewEntireForest = $true
}

@(Get-Content -Path ".\CheckedMailbox.Set$i.txt") | ForEach-Object `
{
	$ExOrg_GetCASMailbox_outputfile = $output_location + "\\Set$i~~GetCASMailbox.txt"
    $mailbox = $_
	@(Get-CASMailbox -identity $mailbox -ErrorAction continue) | ForEach-Object `
	{
		$output_ExOrg_GetCASMailbox = $mailbox + "`t" + `
			$_.Identity + "`t" + `
			$_.ServerName + "`t" + `
			$_.ActiveSyncMailboxPolicy + "`t" + `
		    $_.ActiveSyncEnabled + "`t" + `
		    $_.HasActiveSyncDevicePartnership + "`t" + `
			$_.OwaMailboxPolicy + "`t" + `
			$_.OWAEnabled + "`t" + `
			$_.ECPEnabled + "`t" + `
			$_.EmwsEnabled + "`t" + `
			$_.PopEnabled + "`t" + `
			$_.ImapEnabled + "`t" + `
			$_.MAPIEnabled + "`t" + `
			$_.MAPIBlockOutlookNonCachedMode + "`t" + `
			$_.MAPIBlockOutlookVersions + "`t" + `
			$_.MAPIBlockOutlookRpcHttp + "`t" + `
			$_.EwsEnabled + "`t" + `
			$_.EwsAllowOutlook + "`t" + `
			$_.EwsAllowMacOutlook + "`t" + `
			$_.EwsAllowEntourage 
		$output_ExOrg_GetCASMailbox | Out-File -FilePath $ExOrg_GetCASMailbox_outputfile -append 
	}
}

$a = Get-Process -pid $PID

$EventText = "ExOrg_GetCASMailbox " + "`n" + $server + "`n"
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

if ($PSSession -ne $null)
{
	try{Remove-PSSession -Session $session}catch{}
}
