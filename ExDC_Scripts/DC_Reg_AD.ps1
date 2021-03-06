#############################################################################
#                             DC_Reg_AD.ps1		 							#
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
Param($server,$location,$append)

Write-Output -InputObject $PID
Write-Output -InputObject "WMI"

$ErrorActionPreference = "Stop"
Trap {
$ErrorText = "DC_Reg_AD " + "`n" + $server + "`n"
$ErrorText += $_

$ErrorLog = New-Object -TypeName System.Diagnostics.EventLog -ArgumentList Application
$ErrorLog.MachineName = "."
$ErrorLog.Source = "ExDC"
$ErrorLog.WriteEntry($ErrorText,"Error", 100)
}

Set-Location -LiteralPath $location
$output_location = $location + "\output\DC_Reg_AD"

if ((Test-Path -LiteralPath $output_location) -eq $false)
    {New-Item -Path $output_location -ItemType directory -Force}

$DC_Reg_AD_outputfile = $output_location + "\" +$server + "~~DC_Reg_AD.txt"

# AD Queries
$Forest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
$ForestName = $Forest.tostring()
$ForestMode = $Forest.ForestMode
foreach ($GC in $forest.GlobalCatalogs)
{
	if ($GC.name -eq $server)
	{
		$isGC = "True"
	}
}
Foreach ($domain in $forest.domains)
{
	Foreach ($dc in $domain.domaincontrollers)
	{
		If ($dc.name -eq $server)
		{
			$DomainName = $domain.ToString()
			$DomainMode = $domain.DomainMode
			$Roles = $dc.Roles
			$SiteName = $dc.sitename	
		}
	}
}

# NTDS Queries
$regNtdsKey = "SYSTEM\CurrentControlSet\Services\NTDS\Parameters"
$regNtdsKeyType = [Microsoft.Win32.RegistryHive]::LocalMachine
$regNtdsRemoteBase=[Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($regNtdsKeyType,$server)
$regNtdsRegKey = $regNtdsRemoteBase.OpenSubKey($regNtdsKey)
$regNtdsTypeValue = $regNtdsRegKey.getvalue("Type")
$regNtdsTcpIpPortValue = $regNtdsRegKey.getvalue("TCP/IP Port")

# NTFRS Queries
$regNtfrsKey = "SYSTEM\CurrentControlSet\Services\NTFRS\Parameters"
$regNtfrsKeyType = [Microsoft.Win32.RegistryHive]::LocalMachine
$regNtfrsRemoteBase=[Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($regNtfrsKeyType,$server)
$regNtfrsRegKey = $regNtfrsRemoteBase.OpenSubKey($regNtfrsKey)
$regNtfrsTypeValue = $regNtfrsRegKey.getvalue("Type")
$regNtfrsRpcTcpIpPortValue = $regNtfrsRegKey.getvalue("RPC TCP/IP Port Assignment")					

# Netlogon Queries
$regNetlogonKey = "SYSTEM\CurrentControlSet\Services\Netlogon\Parameters"
$regNetlogonKeyType = [Microsoft.Win32.RegistryHive]::LocalMachine
$regNetlogonRemoteBase=[Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($regNetlogonKeyType,$server)
$regNetlogonRegKey = $regNetlogonRemoteBase.OpenSubKey($regNetlogonKey)
$regNetlogonTypeValue = $regNetlogonRegKey.getvalue("Type")
$regNetlogonDCTcpipPortValue = $regNetlogonRegKey.getvalue("DCTcpipPort")		
$regNetlogonMaxConcurrentApi = $regNetlogonRegKey.getvalue("MaxConcurrentApi")		

$output_DC_Reg_AD = $server + "`t" + `
	$ForestName + "`t" + `
	$DomainName + "`t" + `
	$SiteName + "`t" + `
	$IsGC + "`t" + `
	$Roles + "`t" + `
	$ForestMode + "`t" + `
	$DomainMode + "`t" + `
	$regNtdsTcpIpPortValue + "`t" + `
	$regNtfrsRpcTcpIpPortValue + "`t" + `
	$regNetlogonDCTcpipPortValue + "`t" + `
    $regNetlogonMaxConcurrentApi

$output_DC_Reg_AD | Out-File -FilePath $DC_Reg_AD_outputfile -append 

$a = Get-Process -pid $PID

$EventText = "DC_Reg_AD " + "`n" + $server + "`n"
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


