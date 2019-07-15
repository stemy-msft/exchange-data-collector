#############################################################################
#                          ExOrg_GetUser.ps1 								#
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
$ErrorText = "ExOrg_GetUser " + "`n" + $server + "`n"
$ErrorText += $_

$ErrorLog = New-Object -TypeName System.Diagnostics.EventLog -ArgumentList Application
$ErrorLog.MachineName = "."
$ErrorLog.Source = "ExDC"
$ErrorLog.WriteEntry($ErrorText,"Error", 100)
}

set-location -LiteralPath $location
$output_location = $location + "\output\ExOrg\GetUser"

if ((Test-Path -LiteralPath $output_location) -eq $false)
    {New-Item -Path $output_location -ItemType directory -Force}

if ($PSSession -ne $null)
{
	$cxnUri = "http://" + $PSSession + "/powershell"
	$session = New-PSSession -configurationName Microsoft.Exchange -ConnectionUri $cxnUri -authentication Kerberos
	Import-PSSession -Session $session -AllowClobber -CommandName Set-AdServerSettings,Get-User
    Set-AdServerSettings -ViewEntireForest $true
}
else
{
	Add-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.Admin
    ([Microsoft.Exchange.Data.Directory.AdminSessionADSettings]::Instance).ViewEntireForest = $true
}

@(Get-Content -Path ".\CheckedMailbox.Set$i.txt") | ForEach-Object `
{
	$ExOrg_GetUser_outputfile = $output_location + "\\Set$i~~GetUser.txt"
    $mailbox = $_
	@(Get-User -identity $mailbox) | ForEach-Object `
	{
		$output_ExOrg_GetUser = $mailbox + "`t" + `
			$_.SamAccountName + "`t" + `
			$_.Sid + "`t" + `
			$_.SidHistory + "`t" + `
			$_.UserPrincipalName + "`t" + `
			$_.ResetPasswordOnNextLogon + "`t" + `
			$_.CertificateSubject + "`t" + `
			$_.RemotePowerShellEnabled + "`t" + `
			$_.WindowsLiveID + "`t" + `
			$_.MicrosoftOnlineServicesID + "`t" + `
			$_.NetID + "`t" + `
			$_.ConsumerNetID + "`t" + `
			$_.UserAccountControl + "`t" + `
			$_.OrganizationalUnit + "`t" + `
			$_.IsLinked + "`t" + `
			$_.LinkedMasterAccount + "`t" + `
			$_.ExternalDirectoryObjectId + "`t" + `
			$_.SKUAssigned + "`t" + `
			$_.IsSoftDeletedByRemove + "`t" + `
			$_.IsSoftDeletedByDisable + "`t" + `
			$_.WhenSoftDeleted + "`t" + `
			$_.PreviousRecipientTypeDetails + "`t" + `
			$_.UpgradeRequest + "`t" + `
			$_.UpgradeStatus + "`t" + `
			$_.UpgradeDetails + "`t" + `
			$_.UpgradeMessage + "`t" + `
			$_.UpgradeStage + "`t" + `
			$_.UpgradeStageTimeStamp + "`t" + `
			$_.MailboxRegion + "`t" + `
			$_.MailboxRegionLastUpdateTime + "`t" + `
			$_.MailboxProvisioningConstraint + "`t" + `
			$_.MailboxProvisioningPreferences + "`t" + `
			$_.LegacyExchangeDN + "`t" + `
			$_.InPlaceHoldsRaw + "`t" + `
			$_.MailboxRelease + "`t" + `
			$_.ArchiveRelease + "`t" + `
			$_.AccountDisabled + "`t" + `
			$_.AuthenticationPolicy + "`t" + `
			$_.StsRefreshTokensValidFrom + "`t" + `
			$_.MailboxLocations + "`t" + `
			$_.AdministrativeUnits + "`t" + `
			$_.AssistantName + "`t" + `
			$_.City + "`t" + `
			$_.Company + "`t" + `
			$_.CountryOrRegion + "`t" + `
			$_.Department + "`t" + `
			$_.DirectReports + "`t" + `
			$_.DisplayName + "`t" + `
			$_.Fax + "`t" + `
			$_.FirstName + "`t" + `
			$_.GeoCoordinates + "`t" + `
			$_.HomePhone + "`t" + `
			$_.Initials + "`t" + `
			$_.IsDirSynced + "`t" + `
			$_.LastName + "`t" + `
			$_.Manager + "`t" + `
			$_.MobilePhone + "`t" + `
			$_.Notes + "`t" + `
			$_.Office + "`t" + `
			$_.OtherFax + "`t" + `
			$_.OtherHomePhone + "`t" + `
			$_.OtherTelephone + "`t" + `
			$_.Pager + "`t" + `
			$_.Phone + "`t" + `
			$_.PhoneticDisplayName + "`t" + `
			$_.PostalCode + "`t" + `
			$_.PostOfficeBox + "`t" + `
			$_.RecipientType + "`t" + `
			$_.RecipientTypeDetails + "`t" + `
			$_.SimpleDisplayName + "`t" + `
			$_.StateOrProvince + "`t" + `
			$_.StreetAddress + "`t" + `
			$_.Title + "`t" + `
			$_.UMDialPlan + "`t" + `
			$_.UMDtmfMap + "`t" + `
			$_.AllowUMCallsFromNonUsers + "`t" + `
			$_.WebPage + "`t" + `
			$_.TelephoneAssistant + "`t" + `
			$_.WindowsEmailAddress + "`t" + `
			$_.UMCallingLineIds + "`t" + `
			$_.SeniorityIndex + "`t" + `
			$_.VoiceMailSettings + "`t" + `
			$_.Identity + "`t" + `
			$_.IsValid + "`t" + `
			$_.ExchangeVersion + "`t" + `
			$_.Name + "`t" + `
			$_.DistinguishedName + "`t" + `
			$_.Guid + "`t" + `
			$_.WhenChangedUTC + "`t" + `
			$_.WhenCreatedUTC

		$output_ExOrg_GetUser | Out-File -FilePath $ExOrg_GetUser_outputfile -append
	}
}

$a = Get-Process -pid $PID

$EventText = "ExOrg_GetUser " + "`n" + $server + "`n"
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
