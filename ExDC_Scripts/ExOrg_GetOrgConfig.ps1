#############################################################################
#                     	ExOrg_GetOrgConfig.ps1								#
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
$ErrorText = "ExOrg_GetOrgConfig " + "`n" + $server + "`n"
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
	Import-PSSession -Session $session -AllowClobber -CommandName Get-OrganizationConfig,Set-AdServerSettings
    Set-AdServerSettings -ViewEntireForest $true
}
else
{
	Add-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.Admin
    ([Microsoft.Exchange.Data.Directory.AdminSessionADSettings]::Instance).ViewEntireForest = $true
}

$ExOrg_GetOrgConfig_outputfile = $output_location + "\ExOrg_GetOrgConfig.txt"

@(Get-OrganizationConfig) | ForEach-Object `
{
	$output_ExOrg_GetOrgConfig = $_.Name + "`t" + `
		$_.DefaultPublicFolderDatabase + "`t" + `
		$_.IssueWarningQuota + "`t" + `
		$_.PublicFolderContentReplicationDisabled + "`t" + `
		$_.PublicFoldersLockedForMigration + "`t" + `
		$_.PublicFolderMigrationComplete + "`t" + `
		$_.PublicFoldersEnabled + "`t" + `
		$_.DefaultPublicFolderAgeLimit + "`t" + `
		$_.DefaultPublicFolderIssueWarningQuota + "`t" + `
		$_.DefaultPublicFolderProhibitPostQuota + "`t" + `
		$_.DefaultPublicFolderMaxItemSize + "`t" + `
		$_.DefaultPublicFolderDeletedItemRetention + "`t" + `
		$_.DefaultPublicFolderMovedItemRetention + "`t" + `
		$_.PublicFolderMailboxesLockedForNewConnections + "`t" + `
		$_.PublicFolderMailboxesMigrationComplete + "`t" + `
		$_.IsMixedMode + "`t" + `
		$_.SCLJunkThreshold + "`t" + `
		$_.Industry + "`t" + `
		$_.CustomerFeedbackEnabled + "`t" + `
		$_.OrganizationSummary + "`t" + `
		$_.MailTipsExternalRecipientsTipsEnabled + "`t" + `
		$_.MailTipsLargeAudienceThreshold + "`t" + `
		$_.MailTipsMailboxSourcedTipsEnabled + "`t" + `
		$_.MailTipsGroupMetricsEnabled + "`t" + `
		$_.MailTipsAllTipsEnabled + "`t" + `
		$_.ReadTrackingEnabled + "`t" + `
		$_.DistributionGroupDefaultOU + "`t" + `
		$_.DistributionGroupNameBlockedWordsList + "`t" + `
		$_.DistributionGroupNamingPolicy + "`t" + `
		$_.EwsEnabled + "`t" + `
		$_.EwsAllowOutlook + "`t" + `
		$_.EwsAllowMacOutlook + "`t" + `
		$_.EwsAllowEntourage + "`t" + `
		$_.EwsApplicationAccessPolicy + "`t" + `
		$_.EwsAllowList + "`t" + `
		$_.EwsBlockList + "`t" + `
		$_.ActivityBasedAuthenticationTimeoutInterval + "`t" + `
		$_.ActivityBasedAuthenticationTimeoutEnabled + "`t" + `
		$_.ActivityBasedAuthenticationTimeoutWithSingleSignOnEnabled + "`t" + `
		$_.IPListBlocked + "`t" + `
		$_.AutoExpandingArchiveEnabled + "`t" + `
		$_.MaxConcurrentMigrations + "`t" + `
		$_.IntuneManagedStatus + "`t" + `
		$_.AzurePremiumSubscriptionStatus + "`t" + `
		$_.HybridConfigurationStatus + "`t" + `
		$_.UnblockUnsafeSenderPromptEnabled


	$output_ExOrg_GetOrgConfig | Out-File -FilePath $ExOrg_GetOrgConfig_outputfile -append 
}

$a = Get-Process -pid $PID

$EventText = "ExOrg_GetOrgConfig " + "`n" + $server + "`n"
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
