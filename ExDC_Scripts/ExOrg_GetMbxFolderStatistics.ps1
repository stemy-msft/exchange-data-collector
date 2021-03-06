#############################################################################
#                   ExOrg_GetMbxFolderStatistics.ps1 						#
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
#     Modify $ItemCountThreshold variable to change the data collected		#
#############################################################################
Param($location,$i,$Session_0)

$ItemCountThreshold = 1000

Write-Output -InputObject $PID
Write-Output -InputObject "ExOrg"

$ErrorActionPreference = "Stop"
Trap {
$ErrorText = "ExOrg_GetMbxFolderStatistics " + "`n" + $server + "`n"
$ErrorText += $_

$ErrorLog = New-Object -TypeName System.Diagnostics.EventLog -ArgumentList Application
$ErrorLog.MachineName = "."
$ErrorLog.Source = "ExDC"
$ErrorLog.WriteEntry($ErrorText,"Error", 100)
}

set-location -LiteralPath $location
$output_location = $location + "\output\ExOrg\GetMbxFolderStatistics"

if ((Test-Path -LiteralPath $output_location) -eq $false)
    {New-Item -Path $output_location -ItemType directory -Force}

# Initial connection
if ($session_0 -ne $null)
{
	$cxnUri = "http://" + $session_0 + "/powershell"
	$session = New-PSSession -configurationName Microsoft.Exchange -ConnectionUri $cxnUri -authentication Kerberos
	Import-PSSession -Session $session -AllowClobber -CommandName Get-ExchangeServer,Get-MailboxFolderStatistics,Set-AdServerSettings
    Set-AdServerSettings -ViewEntireForest $true
}
else
{
	Add-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.Admin
    ([Microsoft.Exchange.Data.Directory.AdminSessionADSettings]::Instance).ViewEntireForest = $true
}

$ServerList = Get-ExchangeServer
$Ex2013or2016 = $null
$Ex2010 = $null
$Ex2013or2016List = @()
$Ex2010List = @()
ForEach ($Server in $ServerList)
{
    If ($Server.AdminDisplayVersion -match "Version 14")
    {
        $Ex2010 = $true
        $Ex2010List += $Server.Name
    }
    If ($Server.AdminDisplayVersion -match "Version 15")
    {
        $Ex2013or2016 = $true
        $Ex2013or2016List += $Server.Name
    }
}

#If environment contains Ex2010, then Powershell needs to use Ex2010 Powershell for Get-MailboxFolderStatistics
If (($Ex2010 -eq $true) -and ($Ex2013or2016 -eq $true))
{
    # Let's connect to Ex2010 powershell instead
    # Max value for int... it only gets better than this.
    [int]$PingAverageMin = 2147483647
    [string]$session_Ex2010 = $null
    Foreach ($Ex2010Server in $Ex2010List)
    {
        # Find server with lowest ping time
        $PingAverage = (Test-Connection $Ex2010Server -Count 2 | Measure-Object responsetime -Average).Average
        If ($PingAverage -lt $PingAverageMin)
        {
            $PingAverageMin = $PingAverage
            $session_Ex2010 = $Ex2010Server
        }
    }
	$cxnUri = "http://" + $session_Ex2010 + "/powershell"
	$session = New-PSSession -configurationName Microsoft.Exchange -ConnectionUri $cxnUri -authentication Kerberos
	$ImportedSession = Import-PSSession -Session $session -AllowClobber -CommandName Get-MailboxFolderStatistics
}

$Mailboxes = @(Get-Content -Path ".\CheckedMailbox.Set$i.txt") 
Foreach ($mailbox in $Mailboxes)
{
	$ExOrg_GetMbxFolderStatistics_outputfile = $output_location + "\\Set$i~~GetMbxFolderStatistics.txt"
    #$mailbox = $_
	$Stats = @(Get-MailboxFolderStatistics -identity $mailbox -ErrorAction continue) 
    foreach ($Folder in $stats)
	{	
		if ($Folder.ItemsInFolder -ge $ItemCountThreshold)
		{
			[string]$FolderSize = [string]$Folder.Foldersize
			$FolderSizeBytesLeft = $FolderSize.split("(")
			$FolderSizeBytesRight = $FolderSizeBytesLeft[1].split(" bytes)")
			$FolderSizeBytes = [long]$FolderSizeBytesRight[0]
			$FolderSizeMB = [Math]::Round($FolderSizeBytes/1048576,2)

			$output_ExOrg_GetMbxFolderStatistics = $mailbox + "`t" + `
				$Folder.name + "`t" + `
				$Folder.FolderType + "`t" + `
				$Folder.identity + "`t" + `
				$Folder.ItemsInFolder + "`t" + `
				$Folder.FolderSize + "`t" + `
				$FolderSizeMB
			$output_ExOrg_GetMbxFolderStatistics | Out-File -FilePath $ExOrg_GetMbxFolderStatistics_outputfile -append 
		}
	}
}

$a = Get-Process -pid $PID

$EventText = "ExOrg_GetMbxFolderStatistics " + "`n" + $server + "`n"
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
