#############################################################################
#                ExOrg_GetManagementRoleAssignment.ps1		 				#
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
$ErrorText = "ExOrg_GetManagementRoleAssignment " + "`n" + $server + "`n"
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

$ExOrg_GetManagementRoleAssignment_outputfile = $output_location + "\ExOrg_GetManagementRoleAssignment.txt"

if ($session_0 -ne $null)
{
	$cxnUri = "http://" + $session_0 + "/powershell"
	$session = New-PSSession -configurationName Microsoft.Exchange -ConnectionUri $cxnUri -authentication Kerberos
	Import-PSSession -Session $session -AllowClobber -CommandName Get-RoleGroup,Get-ManagementRoleAssignment,Set-AdServerSettings
    Set-AdServerSettings -ViewEntireForest $true
}
else
{
	Add-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.Admin
    ([Microsoft.Exchange.Data.Directory.AdminSessionADSettings]::Instance).ViewEntireForest = $true
}

#$ManagementRoleAssignment = @(Get-ManagementRoleAssignment -GetEffectiveUsers)
#$AdminRoles = $ManagementRoleAssignment.RoleAssigneeName | Sort-Object -Unique

$RoleGroups = @(Get-RoleGroup)
$RoleGroupCount = $RoleGroups.count
$DataArray = New-Object 'object[,]' -ArgumentList ($RoleGroupCount+1),3
$Row = 0
Foreach ($Group in $RoleGroups)
{
	$Members = @($Group.members)
	$Roles = @($Group.roles)
	$MemberArray = @()
	$RolesArray = @()
	Foreach ($member in $Members)
	{
		$MemberArray += $member.name	
	}
	Foreach ($role in $Roles)
	{
		$RolesArray += $role.name
	}
	
	$DataArray[$Row,0] = $Group.name
	$DataArray[$Row,1] = $MemberArray
	$DataArray[$Row,2] =  $RolesArray
	$Row++
}


@(Get-ManagementRoleAssignment -GetEffectiveUsers) | ForEach-Object `
{
	$output_ExOrg_ManagementRoleAssignment = $_.Name + "`t" + `
		$_.RoleAssigneeName + "`t" + `
		$_.MessageClass + "`t" + `
		$_.RetentionEnabled + "`t" + `
		$_.RetentionAction + "`t" + `
		$_.AgeLimitForRetention + "`t" + `
		$_.MoveToDestinationFolder + "`t" + `
		$_.TriggerForRetention + "`t" + `
		$_.Type + "`t" + `
		$_.WhenCreatedUTC + "`t" + `
		$_.WhenChangedUTC 
	$output_ExOrg_ManagementRoleAssignment | Out-File -FilePath $ExOrg_GetManagementRoleAssignment_outputfile -append 
}

$a = Get-Process -pid $PID

$EventText = "ExOrg_GetManagementRoleAssignment " + "`n" + $server + "`n"
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
