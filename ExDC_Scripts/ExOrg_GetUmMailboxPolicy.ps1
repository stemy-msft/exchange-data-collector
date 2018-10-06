#############################################################################
#                    ExOrg_GetUmMailboxPolicy.ps1		 					#
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
$ErrorText = "ExOrg_GetUmMailboxPolicy " + "`n" + $server + "`n"
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

if ($PSSession -ne $null)
{
	$cxnUri = "http://" + $PSSession + "/powershell"
	$session = New-PSSession -configurationName Microsoft.Exchange -ConnectionUri $cxnUri -authentication Kerberos
	Import-PSSession -Session $session -AllowClobber -CommandName Get-UmMailboxPolicy,Set-AdServerSettings
    Set-AdServerSettings -ViewEntireForest $true
}
else
{
	Add-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.Admin
    ([Microsoft.Exchange.Data.Directory.AdminSessionADSettings]::Instance).ViewEntireForest = $true
}

$ExOrg_GetUmMailboxPolicy_outputfile = $output_location + "\ExOrg_GetUmMailboxPolicy.txt"

@(Get-UmMailboxPolicy) | ForEach-Object `
{
	$output_ExOrg_GetUmMailboxPolicy = $_.name + "`t" + `
		$_.Identity + "`t" + `
		$_.MaxGreetingDuration + "`t" + `
		$_.MaxLogonAttempts + "`t" + `
		$_.AllowCommonPatterns + "`t" + `
		$_.PINLifetime + "`t" + `
		$_.PINHistoryCount + "`t" + `
		$_.AllowSMSNotification + "`t" + `
		$_.ProtectUnauthenticatedVoiceMail + "`t" + `
		$_.ProtectAuthenticatedVoiceMail + "`t" + `
		$_.ProtectedVoiceMailText + "`t" + `
		$_.RequireProtectedPlayOnPhone + "`t" + `
		$_.MinPINLength + "`t" + `
		$_.FaxMessageText + "`t" + `
		$_.UMEnabledText + "`t" + `
		$_.ResetPINText + "`t" + `
		$_.SourceForestPolicyNames + "`t" + `
		$_.VoiceMailText + "`t" + `
		$_.UMDialPlan + "`t" + `
		$_.FaxServerURI + "`t" + `
		$_.AllowedInCountryOrRegionGroups + "`t" + `
		$_.AllowedInternationalGroups + "`t" + `
		$_.AllowDialPlanSubscribers + "`t" + `
		$_.AllowExtensions + "`t" + `
		$_.LogonFailuresBeforePINReset + "`t" + `
		$_.AllowMissedCallNotifications + "`t" + `
		$_.AllowFax + "`t" + `
		$_.AllowTUIAccessToCalendar + "`t" + `
		$_.AllowTUIAccessToEmail + "`t" + `
		$_.AllowSubscriberAccess + "`t" + `
		$_.AllowTUIAccessToDirectory + "`t" + `
		$_.AllowTUIAccessToPersonalContacts + "`t" + `
		$_.AllowAutomaticSpeechRecognition + "`t" + `
		$_.AllowPlayOnPhone + "`t" + `
		$_.AllowVoiceMailPreview + "`t" + `
		$_.AllowCallAnsweringRules + "`t" + `
		$_.AllowMessageWaitingIndicator + "`t" + `
		$_.AllowPinlessVoiceMailAccess + "`t" + `
		$_.AllowVoiceResponseToOtherMessageTypes + "`t" + `
		$_.AllowVoiceMailAnalysis + "`t" + `
		$_.AllowVoiceNotification + "`t" + `
		$_.InformCallerOfVoiceMailAnalysis + "`t" + `
		$_.VoiceMailPreviewPartnerAddress + "`t" + `
		$_.VoiceMailPreviewPartnerAssignedID + "`t" + `
		$_.VoiceMailPreviewPartnerMaxMessageDuration + "`t" + `
		$_.VoiceMailPreviewPartnerMaxDeliveryDelay + "`t" + `
		$_.IsDefault + "`t" + `
		$_.IsValid
	$output_ExOrg_GetUmMailboxPolicy | Out-File -FilePath $ExOrg_GetUmMailboxPolicy_outputfile -append 
}

$a = Get-Process -pid $PID

$EventText = "ExOrg_GetUmMailboxPolicy " + "`n" + $server + "`n"
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
