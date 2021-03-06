#############################################################################
#                    ExOrg_GetThrottlingPolicy.ps1							#
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
Param($location,$append,$session_0)

Write-Output -InputObject $PID
Write-Output -InputObject "ExOrg"

$ErrorActionPreference = "Stop"
Trap {
$ErrorText = "ExOrg_GetThrottlingPolicy " + "`n" + $server + "`n"
$ErrorText += $_

$ErrorLog = New-Object -TypeName System.Diagnostics.EventLog -ArgumentList Application
$ErrorLog.MachineName = "."
$ErrorLog.Source = "ExDC"
$ErrorLog.WriteEntry($ErrorText,"Error", 100)
}

set-location -LiteralPath $location
$output_location = $location + "\output\ExOrg"

if ((Test-Path -LiteralPath $output_location) -eq $false)
    {New-Item -Path $output_location -ItemType directory -Force}

if ($session_0 -ne $null)
{
	$cxnUri = "http://" + $session_0 + "/powershell"
	$session = New-PSSession -configurationName Microsoft.Exchange -ConnectionUri $cxnUri -authentication Kerberos
	Import-PSSession -Session $session -AllowClobber -CommandName Get-Exchangeserver,Get-ThrottlingPolicy,Set-AdServerSettings
    Set-AdServerSettings -ViewEntireForest $true
}
else
{
	Add-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.Admin
    ([Microsoft.Exchange.Data.Directory.AdminSessionADSettings]::Instance).ViewEntireForest = $true
}

If ((Get-ExchangeServer $session_0).IsE15orLater -eq $true)
	{$Legacy = $false}
	else
	{$Legacy = $true}

$ExOrg_GetThrottlingPolicy_outputfile = $output_location + "\ExOrg_GetThrottlingPolicy.txt"

@(Get-ThrottlingPolicy) | ForEach-Object `
{	
    If ($Legacy = $true)
	{
		$output_ExOrg_GetThrottlingPolicy = $_.Identity + "`t" + `
		$_.Name + "`t" + `
		$_.IsLegacyDefault + "`t" + `
	    $_.EWSMaxConcurrency + "`t" + `
	    $_.EWSPercentTimeInAD + "`t" + `
	    $_.EWSPercentTimeInCAS + "`t" + `
	    $_.EWSPercentTimeInMailboxRPC + "`t" + `
	    $_.EWSMaxSubscriptions + "`t" + `
	    $_.EWSFastSearchTimeoutInSeconds + "`t" + `
	    $_.EWSFindCountLimit + "`t" + `
	    $_.RCAMaxConcurrency + "`t" + `
	    $_.RCAPercentTimeInAD + "`t" + `
	    $_.RCAPercentTimeInCAS + "`t" + `
	    $_.RCAPercentTimeInMailboxRPC + "`t" + `
	    $_.CPAMaxConcurrency + "`t" + `
	    $_.CPAPercentTimeInCAS + "`t" + `
	    $_.CPAPercentTimeInMailboxRPC + "`t" + `
	    $_.AnonymousMaxConcurrency  + "`t" + `
	    $_.AnonymousPercentTimeInAD + "`t" + `
	    $_.AnonymousPercentTimeInCAS + "`t" + `
	    $_.AnonymousPercentTimeInMailboxRPC + "`t" + `
	    $_.EASMaxConcurrency + "`t" + `
	    $_.EASPercentTimeInAD + "`t" + `
	    $_.EASPercentTimeInCAS + "`t" + `
	    $_.EASPercentTimeInMailboxRPC + "`t" + `
	    $_.EASMaxDevices + "`t" + `
	    $_.EASMaxDeviceDeletesPerMonth + "`t" + `
	    $_.IMAPMaxConcurrency + "`t" + `
	    $_.IMAPPercentTimeInAD + "`t" + `
	    $_.IMAPPercentTimeInCAS + "`t" + `
	    $_.IMAPPercentTimeInMailboxRPC + "`t" + `
	    $_.OWAMaxConcurrency + "`t" + `
	    $_.OWAPercentTimeInAD + "`t" + `
	    $_.OWAPercentTimeInCAS + "`t" + `
	    $_.OWAPercentTimeInMailboxRPC + "`t" + `
	    $_.POPMaxConcurrency + "`t" + `
	    $_.POPPercentTimeInAD + "`t" + `
	    $_.POPPercentTimeInCAS + "`t" + `
	    $_.POPPercentTimeInMailboxRPC + "`t" + `
	    $_.PowerShellMaxConcurrency + "`t" + `
	    $_.PowerShellMaxTenantConcurrency + "`t" + `
	    $_.PowerShellMaxCmdlets + "`t" + `
	    $_.PowerShellMaxCmdletsTimePeriod + "`t" + `
	    $_.ExchangeMaxCmdlets + "`t" + `
	    $_.PowerShellMaxCmdletQueueDepth + "`t" + `
	    $_.PowerShellMaxDestructiveCmdlets + "`t" + `
	    $_.PowerShellMaxDestructiveCmdletsTimePeriod + "`t" + `
	    $_.MessageRateLimit + "`t" + `
	    $_.RecipientRateLimit + "`t" + `
	    $_.ForwardeeLimit + "`t" + `
	    $_.CPUStartPercent + "`t" + `
	    $_.DiscoveryMaxConcurrency + "`t" + `
	    $_.DiscoveryMaxMailboxes + "`t" + `
	    $_.WhenCreatedUTC + "`t" + `
	    $_.WhenChangedUTC 
	}
	else
	{
		$output_ExOrg_GetThrottlingPolicy = $_.Identity + "`t" + `
		$_.Name + "`t" + `
		$_.IsDefault + "`t" + `
	    $_.EWSMaxConcurrency + "`t" + `
	    $_.EWSPercentTimeInAD + "`t" + `
	    $_.EWSPercentTimeInCAS + "`t" + `
	    $_.EWSPercentTimeInMailboxRPC + "`t" + `
	    $_.EWSMaxSubscriptions + "`t" + `
	    $_.EWSFastSearchTimeoutInSeconds + "`t" + `
	    $_.EWSFindCountLimit + "`t" + `
	    $_.RCAMaxConcurrency + "`t" + `
	    $_.RCAPercentTimeInAD + "`t" + `
	    $_.RCAPercentTimeInCAS + "`t" + `
	    $_.RCAPercentTimeInMailboxRPC + "`t" + `
	    $_.CPAMaxConcurrency + "`t" + `
	    $_.CPAPercentTimeInCAS + "`t" + `
	    $_.CPAPercentTimeInMailboxRPC + "`t" + `
	    $_.AnonymousMaxConcurrency  + "`t" + `
	    $_.AnonymousPercentTimeInAD + "`t" + `
	    $_.AnonymousPercentTimeInCAS + "`t" + `
	    $_.AnonymousPercentTimeInMailboxRPC + "`t" + `
	    $_.EASMaxConcurrency + "`t" + `
	    $_.EASPercentTimeInAD + "`t" + `
	    $_.EASPercentTimeInCAS + "`t" + `
	    $_.EASPercentTimeInMailboxRPC + "`t" + `
	    $_.EASMaxDevices + "`t" + `
	    $_.EASMaxDeviceDeletesPerMonth + "`t" + `
	    $_.IMAPMaxConcurrency + "`t" + `
	    $_.IMAPPercentTimeInAD + "`t" + `
	    $_.IMAPPercentTimeInCAS + "`t" + `
	    $_.IMAPPercentTimeInMailboxRPC + "`t" + `
	    $_.OWAMaxConcurrency + "`t" + `
	    $_.OWAPercentTimeInAD + "`t" + `
	    $_.OWAPercentTimeInCAS + "`t" + `
	    $_.OWAPercentTimeInMailboxRPC + "`t" + `
	    $_.POPMaxConcurrency + "`t" + `
	    $_.POPPercentTimeInAD + "`t" + `
	    $_.POPPercentTimeInCAS + "`t" + `
	    $_.POPPercentTimeInMailboxRPC + "`t" + `
	    $_.PowerShellMaxConcurrency + "`t" + `
	    $_.PowerShellMaxTenantConcurrency + "`t" + `
	    $_.PowerShellMaxCmdlets + "`t" + `
	    $_.PowerShellMaxCmdletsTimePeriod + "`t" + `
	    $_.ExchangeMaxCmdlets + "`t" + `
	    $_.PowerShellMaxCmdletQueueDepth + "`t" + `
	    $_.PowerShellMaxDestructiveCmdlets + "`t" + `
	    $_.PowerShellMaxDestructiveCmdletsTimePeriod + "`t" + `
	    $_.MessageRateLimit + "`t" + `
	    $_.RecipientRateLimit + "`t" + `
	    $_.ForwardeeLimit + "`t" + `
	    $_.CPUStartPercent + "`t" + `
	    $_.DiscoveryMaxConcurrency + "`t" + `
	    $_.DiscoveryMaxMailboxes + "`t" + `
	    $_.WhenCreatedUTC + "`t" + `
	    $_.WhenChangedUTC 
	}
	$output_ExOrg_GetThrottlingPolicy | Out-File -FilePath $ExOrg_GetThrottlingPolicy_outputfile -append 
}

$a = Get-Process -pid $PID

$EventText = "ExOrg_GetThrottlingPolicy " + "`n" + $server + "`n"
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
