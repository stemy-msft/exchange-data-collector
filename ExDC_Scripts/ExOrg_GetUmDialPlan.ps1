#############################################################################
#                       ExOrg_GetUmDialPlan.ps1		 						#
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
$ErrorText = "ExOrg_GetUmDialPlan " + "`n" + $server + "`n"
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
	Import-PSSession -Session $session -AllowClobber -CommandName Get-UmDialPlan,Set-AdServerSettings
    Set-AdServerSettings -ViewEntireForest $true
}
else
{
	Add-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.Admin
    ([Microsoft.Exchange.Data.Directory.AdminSessionADSettings]::Instance).ViewEntireForest = $true
}

$ExOrg_GetUmDialPlan_outputfile = $output_location + "\ExOrg_GetUmDialPlan.txt"

@(Get-UmDialPlan) | ForEach-Object `
{
	$output_ExOrg_GetUmDialPlan = $_.name + "`t" + `
		$_.Identity + "`t" + `
		$_.NumberOfDigitsInExtension + "`t" + `
		$_.LogonFailuresBeforeDisconnect + "`t" + `
		$_.AccessTelephoneNumbers + "`t" + `
		$_.FaxEnabled + "`t" + `
		$_.InputFailuresBeforeDisconnect + "`t" + `
		$_.OutsideLineAccessCode + "`t" + `
		$_.DialByNamePrimary + "`t" + `
		$_.DialByNameSecondary + "`t" + `
		$_.AudioCodec + "`t" + `
		$_.AvailableLanguages + "`t" + `
		$_.DefaultLanguage + "`t" + `
		$_.VoIPSecurity + "`t" + `
		$_.MaxCallDuration + "`t" + `
		$_.MaxRecordingDuration + "`t" + `
		$_.RecordingIdleTimeout + "`t" + `
		$_.PilotIdentifierList + "`t" + `
		$_.UMServers + "`t" + `
		$_.UMMailboxPolicies + "`t" + `
		$_.UMAutoAttendants + "`t" + `
		$_.WelcomeGreetingEnabled + "`t" + `
		$_.AutomaticSpeechRecognitionEnabled + "`t" + `
		$_.PhoneContext + "`t" + `
		$_.WelcomeGreetingFilename + "`t" + `
		$_.InfoAnnouncementFilename + "`t" + `
		$_.OperatorExtension + "`t" + `
		$_.DefaultOutboundCallingLineId + "`t" + `
		$_.Extension + "`t" + `
		$_.MatchedNameSelectionMethod + "`t" + `
		$_.InfoAnnouncementEnabled + "`t" + `
		$_.InternationalAccessCode + "`t" + `
		$_.NationalNumberPrefix + "`t" + `
		$_.InCountryOrRegionNumberFormat + "`t" + `
		$_.InternationalNumberFormat + "`t" + `
		$_.CallSomeoneEnabled + "`t" + `
		$_.ContactScope + "`t" + `
		$_.ContactAddressList + "`t" + `
		$_.SendVoiceMsgEnabled + "`t" + `
		$_.UMAutoAttendant + "`t" + `
		$_.AllowDialPlanSubscribers + "`t" + `
		$_.AllowExtensions + "`t" + `
		$_.AllowedInCountryOrRegionGroups + "`t" + `
		$_.AllowedInternationalGroups + "`t" + `
		$_.ConfiguredInCountryOrRegionGroups + "`t" + `
		$_.LegacyPromptPublishingPoint + "`t" + `
		$_.ConfiguredInternationalGroups + "`t" + `
		$_.UMIPGateway + "`t" + `
		$_.URIType + "`t" + `
		$_.SubscriberType + "`t" + `
		$_.GlobalCallRoutingScheme + "`t" + `
		$_.TUIPromptEditingEnabled + "`t" + `
		$_.CallAnsweringRulesEnabled + "`t" + `
		$_.SipResourceIdentifierRequired + "`t" + `
		$_.FDSPollingInterval + "`t" + `
		$_.EquivalentDialPlanPhoneContexts + "`t" + `
		$_.NumberingPlanFormats + "`t" + `
		$_.AllowHeuristicADCallingLineIdResolution + "`t" + `
		$_.CountryOrRegionCode + "`t" + `
		$_.ExchangeVersion + "`t" + `
		$_.IsValid
	$output_ExOrg_GetUmDialPlan | Out-File -FilePath $ExOrg_GetUmDialPlan_outputfile -append 
}

$a = Get-Process -pid $PID

$EventText = "ExOrg_GetUmDialPlan " + "`n" + $server + "`n"
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
