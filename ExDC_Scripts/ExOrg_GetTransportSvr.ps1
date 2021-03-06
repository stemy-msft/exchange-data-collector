#############################################################################
#                       ExOrg_GetTransportSvr.ps1							#
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
$ErrorText = "ExOrg_GetTransportSvr " + "`n" + $server + "`n"
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
	# First try with preferred cmdlets for newer servers
	$Legacy = $false
	$ImportedSession = Import-PSSession -Session $session -AllowClobber -CommandName Get-TransportService,Set-AdServerSettings
	# Then go back and load the old ones if necessary
	If ($ImportedSession.ExportedCommands.keys -notcontains "Get-TransportService")
	{
		$Legacy=$true
		Import-PSSession -Session $session -AllowClobber -CommandName Get-TransportServer,Set-AdServerSettings
	}
	
	Set-AdServerSettings -ViewEntireForest $true
}
else
{
	Add-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.Admin
    ([Microsoft.Exchange.Data.Directory.AdminSessionADSettings]::Instance).ViewEntireForest = $true
}

$ExOrg_GetTSvr_outputfile = $output_location + "\ExOrg_GetTransportSvr.txt"

If ($Legacy -eq $false)
{
	@(Get-TransportService -ErrorAction Continue) | ForEach-Object `
	{
		$MDCS_Remote_Path = $null
		$MDCS_Value = $null
		try
		{
			# DatabaseMaxCacheSize on all Transport servers
			$MDCS_Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$_.Name)
			$MDCS_Reg_Key = $MDCS_Reg.opensubkey("SYSTEM\\CurrentControlSet\\Services\\MSExchangeTransport")
			$MDCS_LocalPath = $MDCS_reg_key.getvalue("ImagePath")
			$MDCS_LocalPathLength = $MDCS_LocalPath.length 
			$MDCS_DriveLetter = $MDCS_LocalPath.Substring(1,1)
			$MDCS_PathMinusDriveLetter = $MDCS_LocalPath.Substring(3,$MDCS_LocalPathLength-27)
			$MDCS_Remote_Path = "\\" + $_.name + "\" + $mdcs_driveletter + "$" + $MDCS_PathMinusDriveLetter + "EdgeTransport.exe.config"
			$EdgeTransport = Get-Content -Path $MDCS_Remote_Path -ErrorAction Continue 
			$MDCS_Value = $EdgeTransport | Where-Object {$_ -match "DatabaseMaxCacheSize"}
			$SRPE_Value = $EdgeTransport | Where-Object {$_ -match "ShadowRedundancyPromotionEnabled"}
			$QDP_Value = $EdgeTransport | Where-Object {$_ -match "QueueDatabasePath"}
			$QDLP_Value = $EdgeTransport | Where-Object {$_ -match "QueueDatabaseLoggingPath"}
		}
		catch [System.Management.Automation.MethodInvocationException]
		{
			$ErrorText = "ExOrg_GetTransportSvr `n" + $_ + "`n" 
			$ErrorText += "Error encountered: [System.Management.Automation.MethodInvocationException]"
			$ErrorLog = New-Object -TypeName System.Diagnostics.EventLog -ArgumentList Application
			$ErrorLog.MachineName = "."
			$ErrorLog.Source = "ExDC"
			$ErrorLog.WriteEntry($ErrorText,"Error", 100)
		}
		finally
		{
			$output_ExOrg_GetTSvr = $_.Name + "`t" + `
				$_.AntispamAgentsEnabled + "`t" + `
				$_.AntispamUpdatesEnabled + "`t" + `
				$_.DelayNotificationTimeout + "`t" + `
				$_.MessageExpirationTimeout + "`t" + `
				$_.MessageRetryInterval + "`t" + `
				$_.ExternalDNSAdapterEnabled + "`t" + `
				$_.ExternalDNSServers + "`t" + `
				$_.InternalDNSAdapterEnabled + "`t" + `
				$_.InternalDNSServers + "`t" + `
				$_.MaxOutboundConnections + "`t" + `
				$_.MaxPerDomainOutboundConnections + "`t" + `
	   			$_.TransientFailureRetryCount + "`t" + `
	   			$_.ConnectivityLogEnabled + "`t" + `
				$_.MessageTrackingLogEnabled + "`t" + `
				$_.MessageTrackingLogSubjectLoggingEnabled + "`t" + `
				$MDCS_Remote_Path + "`t" + `
				$MDCS_Value + "`t" + `
				$SRPE_Value + "`t" + `
				$QDP_Value + "`t" + `
				$QDLP_Value + "`t" + `
				$_.ActiveUserStatisticsLogPath + "`t" + `
				$_.ConnectivityLogPath + "`t" + `
				$_.MessageTrackingLogPath + "`t" + `
				$_.ReceiveProtocolLogPath + "`t" + `
				$_.RoutingTableLogPath + "`t" + `
				$_.SendProtocolLogPath + "`t" + `
				$_.ServerStatisticsLogPath		
			
			$output_ExOrg_GetTSvr | Out-File -FilePath $ExOrg_GetTSvr_outputfile -append 
		}
	}
}
If ($Legacy -eq $true)
{
	@(Get-TransportServer -ErrorAction Continue) | ForEach-Object `
	{
		$MDCS_Remote_Path = $null
		$MDCS_Value = $null
		try
		{
			# DatabaseMaxCacheSize on all Transport servers
			$MDCS_Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine',$_.Name)
			$MDCS_Reg_Key = $MDCS_Reg.opensubkey("SYSTEM\\CurrentControlSet\\Services\\MSExchangeTransport")
			$MDCS_LocalPath = $MDCS_reg_key.getvalue("ImagePath")
			$MDCS_LocalPathLength = $MDCS_LocalPath.length 
			$MDCS_DriveLetter = $MDCS_LocalPath.Substring(1,1)
			$MDCS_PathMinusDriveLetter = $MDCS_LocalPath.Substring(3,$MDCS_LocalPathLength-27)
			$MDCS_Remote_Path = "\\" + $_.name + "\" + $mdcs_driveletter + "$" + $MDCS_PathMinusDriveLetter + "EdgeTransport.exe.config"
			$EdgeTransport = Get-Content -Path $MDCS_Remote_Path -ErrorAction Continue 
			$MDCS_Value = $EdgeTransport | Where-Object {$_ -match "DatabaseMaxCacheSize"}
			$SRPE_Value = $EdgeTransport | Where-Object {$_ -match "ShadowRedundancyPromotionEnabled"}
			$QDP_Value = $EdgeTransport | Where-Object {$_ -match "QueueDatabasePath"}
			$QDLP_Value = $EdgeTransport | Where-Object {$_ -match "QueueDatabaseLoggingPath"}
		}
		catch [System.Management.Automation.MethodInvocationException]
		{
			$ErrorText = "ExOrg_GetTransportSvr `n" + $_ + "`n" 
			$ErrorText += "Error encountered: [System.Management.Automation.MethodInvocationException]"
			$ErrorLog = New-Object -TypeName System.Diagnostics.EventLog -ArgumentList Application
			$ErrorLog.MachineName = "."
			$ErrorLog.Source = "ExDC"
			$ErrorLog.WriteEntry($ErrorText,"Error", 100)
		}
		finally
		{
			$output_ExOrg_GetTSvr = $_.Name + "`t" + `
				$_.AntispamAgentsEnabled + "`t" + `
				$_.AntispamUpdatesEnabled + "`t" + `
				$_.DelayNotificationTimeout + "`t" + `
				$_.MessageExpirationTimeout + "`t" + `
				$_.MessageRetryInterval + "`t" + `
				$_.ExternalDNSAdapterEnabled + "`t" + `
				$_.ExternalDNSServers + "`t" + `
				$_.InternalDNSAdapterEnabled + "`t" + `
				$_.InternalDNSServers + "`t" + `
				$_.MaxOutboundConnections + "`t" + `
				$_.MaxPerDomainOutboundConnections + "`t" + `
	   			$_.TransientFailureRetryCount + "`t" + `
	   			$_.ConnectivityLogEnabled + "`t" + `
				$_.MessageTrackingLogEnabled + "`t" + `
				$_.MessageTrackingLogSubjectLoggingEnabled + "`t" + `
				$MDCS_Remote_Path + "`t" + `
				$MDCS_Value + "`t" + `
				$SRPE_Value + "`t" + `
				$QDP_Value + "`t" + `
				$QDLP_Value + "`t" + `
				$_.ActiveUserStatisticsLogPath + "`t" + `
				$_.ConnectivityLogPath + "`t" + `
				$_.MessageTrackingLogPath + "`t" + `
				$_.ReceiveProtocolLogPath + "`t" + `
				$_.RoutingTableLogPath + "`t" + `
				$_.SendProtocolLogPath + "`t" + `
				$_.ServerStatisticsLogPath		
			
			$output_ExOrg_GetTSvr | Out-File -FilePath $ExOrg_GetTSvr_outputfile -append 
		}
	}
}

$a = Get-Process -pid $PID

$EventText = "ExOrg_GetTransportSvr " + "`n" + $server + "`n"
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
