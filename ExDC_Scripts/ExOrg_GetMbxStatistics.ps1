#############################################################################
#                    ExOrg_GetMbxStatistics.ps1 							#
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
Param($location,$i,$Session_0)

Write-Output -InputObject $PID
Write-Output -InputObject "ExOrg"

Trap {
$ErrorText = "ExOrg_GetMbxStatistics " + "`n" + $server + "`n"
$ErrorText += $_

$ErrorLog = New-Object -TypeName System.Diagnostics.EventLog -ArgumentList Application
$ErrorLog.MachineName = "."
$ErrorLog.Source = "ExDC"
$ErrorLog.WriteEntry($ErrorText,"Error", 100)
}

set-location -LiteralPath $location
$output_location = $location + "\output\ExOrg\GetMbxStatistics"

if ((Test-Path -LiteralPath $output_location) -eq $false)
    {New-Item -Path $output_location -ItemType directory -Force}

if ($session_0 -ne $null)
{
	$cxnUri = "http://" + $session_0 + "/powershell"
	$session = New-PSSession -configurationName Microsoft.Exchange -ConnectionUri $cxnUri -authentication Kerberos
	Import-PSSession -Session $session -AllowClobber -CommandName Get-MailboxStatistics,Set-AdServerSettings
    Set-AdServerSettings -ViewEntireForest $true
	#if (Test-Path -Path HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup)
    #	{$exinstall = (Get-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup).MsiInstallPath}
    #else
    #	{$exinstall = (Get-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup).MsiInstallPath}
	#$dll = $exinstall + "\Bin\Microsoft.Exchange.Data.dll"
	#[Reflection.Assembly]::LoadFile($dll)


}
else
{
	Add-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.Admin
    ([Microsoft.Exchange.Data.Directory.AdminSessionADSettings]::Instance).ViewEntireForest = $true
}

@(Get-Content -Path ".\CheckedMailbox.Set$i.txt") | ForEach-Object `
{
	$ExOrg_GetMbxStatistics_outputfile = $output_location + "\\Set$i~~GetMbxStatistics.txt"
    $mailbox = $_
	Get-MailboxStatistics -identity $mailbox | ForEach-Object `
    {
       	If ($_.totalitemsize -ne $null)
        {	
			$MailboxSize = [string]$_.totalitemsize
			$MailboxSizeBytesLeft = $MailboxSize.split("(")
			$MailboxSizeBytesRight = $MailboxSizeBytesLeft[1].split(" bytes)")
			$MailboxSizeBytes = [long]$MailboxSizeBytesRight[0]
			$MailboxSizeMB = [Math]::Round($MailboxSizeBytes/1048576,2)
			#$MailboxSize = [Microsoft.Exchange.Data.ByteQuantifiedSize]$_.totalitemsize
            #$MailboxSizeMB = $MailboxSize.tomb()
            #$MailboxSizeMB = ([Microsoft.Exchange.Data.ByteQuantifiedSize]$_.totalitemsize.value).tomb()
            #$MailboxSizeMB = $MailboxSize/1mb
            #$MailboxSizeMB = ([Microsoft.Exchange.Data.ByteQuantifiedSize]$_.totalitemsize.value)/1mb
        }
        If ($_.TotalDeletedItemSize -ne $null)
        {
            $DumpsterSize = [string]$_.TotalDeletedItemSize
			$DumpsterSizeBytesLeft = $DumpsterSize.split("(")
			$DumpsterSizeBytesRight = $DumpsterSizeBytesLeft[1].split(" bytes)")
			$DumpsterSizeBytes = [long]$DumpsterSizeBytesRight[0]
			$DumpsterSizeMB = [Math]::Round($DumpsterSizeBytes/1048576,2)
			#$DumpsterSize = [Microsoft.Exchange.Data.ByteQuantifiedSize]$_.TotalDeletedItemSize
            #$DumpsterSizeMB = $DumpsterSize.tomb()
            #$DumpsterSizeMB = ([Microsoft.Exchange.Data.ByteQuantifiedSize]$_.TotalDeletedItemSize.value).tomb()
            #$DumpsterSizeMB = $DumpsterSize/1mb
            #$DumpsterSizeMB = ([Microsoft.Exchange.Data.ByteQuantifiedSize]$_.TotalDeletedItemSize.value)/1mb
        }

    	$output_ExOrg_GetMbxStatistics = $mailbox + "`t" + `
    		$_.DisplayName + "`t" + `
    		$_.ServerName + "`t" + `
    	    $_.Database + "`t" + `
    	    $_.ItemCount + "`t" + `
    	    $_.TotalItemSize + "`t" + `
    		$MailboxSizeMB + "`t" + `
    	    $_.TotalDeletedItemSize + "`t" + `
    		$DumpsterSizeMB
    	$output_ExOrg_GetMbxStatistics | Out-File -FilePath $ExOrg_GetMbxStatistics_outputfile -append 
    }
}

$a = Get-Process -pid $PID

$EventText = "ExOrg_GetMbxStatistics " + "`n" + $server + "`n"
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
