#############################################################################
#                  ExOrg_GetDynamicDistributionGroup.ps1					#
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
$ErrorText = "ExOrg_GetDynamicDistributionGroup " + "`n" + $server + "`n"
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
	Import-PSSession -Session $session -AllowClobber -CommandName Get-DynamicDistributionGroup,Set-AdServerSettings
    Set-AdServerSettings -ViewEntireForest $true
}
else
{
	Add-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.Admin
    ([Microsoft.Exchange.Data.Directory.AdminSessionADSettings]::Instance).ViewEntireForest = $true
}

$ExOrg_GetDDG_outputfile = $output_location + "\ExOrg_GetDynamicDistributionGroup.txt"

@(Get-DynamicDistributionGroup -resultsize unlimited) | ForEach-Object `
{	
	$output_ExOrg_GetDDG = $_.Name + "`t" + `
		$_.ExpansionServer + "`t" + `
		$_.AcceptMessagesOnlyFrom + "`t" + `
		$_.AcceptMessagesOnlyFromDLMembers + "`t" + `
		$_.Alias + "`t" + `
		$_.GrantSendOnBehalfTo + "`t" + `
		$_.HiddenFromAddressListsEnabled + "`t" + `
		$_.MaxSendSize + "`t" + `
		$_.MaxReceiveSize + "`t" + `
		$_.RejectMessagesFrom + "`t" + `
		$_.RejectMessagesFromDLMembers + "`t" + `
		$_.RequireSenderAuthenticationEnabled + "`t" + `
		$_.ManagedBy + "`t" + `
		$_.OrganizationalUnit + "`t" + `
		$_.RecipientContainer + "`t" + `
		$_.RecipientFilter + "`t" + `
		$_.LdapRecipientFilter + "`t" + `
		$_.IncludedRecipients + "`t" + `
		$_.ReportToManagerEnabled + "`t" + `
		$_.ReportToOriginatorEnabled + "`t" + `
		$_.SendOofMessageToOriginatorEnabled + "`t" + `
		$_.AcceptMessagesOnlyFromSendersOrMembers + "`t" + `
		$_.RejectMessagesFromSendersOrMembers + "`t" + `
		$_.ModeratedBy + "`t" + `
		$_.ModerationEnabled + "`t" + `
		$_.PrimarySmtpAddress + "`t" + `
		$_.RecipientType + "`t" + `
		$_.RecipientTypeDetails + "`t" + `
		$_.WhenCreatedUTC + "`t" + `
		$_.WhenChangedUTC + "`t" + `
		$_.IsValid
	$output_ExOrg_GetDDG | Out-File -FilePath $ExOrg_GetDDG_outputfile -append 
}

$a = Get-Process -pid $PID

$EventText = "ExOrg_GetDynamicDistributionGroup " + "`n" + $server + "`n"
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
