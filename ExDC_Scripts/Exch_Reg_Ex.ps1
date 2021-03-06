#############################################################################
#                             Exch_Reg_Ex.ps1		 						#
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
$ErrorText = "Exch_Reg_Ex " + "`n" + $server + "`n"
$ErrorText += $_

$ErrorLog = New-Object -TypeName System.Diagnostics.EventLog -ArgumentList Application
$ErrorLog.MachineName = "."
$ErrorLog.Source = "ExDC"
$ErrorLog.WriteEntry($ErrorText,"Error", 100)
}

set-location -LiteralPath $location
$output_location = $location + "\output\Exch_Reg_Ex"

if ((Test-Path -LiteralPath $output_location) -eq $false)
    {New-Item -Path $output_location -ItemType directory -Force}

$Exch_Reg_Ex_outputfile = $output_location + "\" + $server + "~~Exch_Reg_Ex.txt"

# MSExchangeDSAccess Queries
$regMUDKey = "SYSTEM\CurrentControlSet\Services\MSExchangeDSAccess\Profiles\Default"
$regMUDKeyType = [Microsoft.Win32.RegistryHive]::LocalMachine
$regMUDRemoteBase=[Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($regMUDKeyType,$server)
$regMUDRegKey = $regMUDRemoteBase.OpenSubKey($regMUDKey)
if ($regMUDRegKey -ne $null)
{
	$regMinUserDC = $regMUDRegKey.getvalue("MinUserDC")
}

# MSExchangeSA Queries
$regSAKey = "SYSTEM\CurrentControlSet\Services\MSExchangeSA\Parameters"
$regSAKeyType = [Microsoft.Win32.RegistryHive]::LocalMachine
$regSARemoteBase=[Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($regSAKeyType,$server)
$regSARegKey = $regSARemoteBase.OpenSubKey($regSAKey)
if ($regSARegKey -ne $null)
{
	$regSATcpIpPort = $regSARegKey.getvalue("TCP/IP Port")
	$regSATcpIpNspiPort = $regSARegKey.getvalue("TCP/IP NSPI Port")
	$regSANspiTargetServerPort = $regSARegKey.getvalue("NSPI Target Server")
	$regSARfrTargetServerPort = $regSARegKey.getvalue("RFR Target Server")
}

# MSExchangeIS Queries
$regISKey = "SYSTEM\CurrentControlSet\Services\MSExchangeIS\ParametersSystem"
$regISKeyType = [Microsoft.Win32.RegistryHive]::LocalMachine
$regISRemoteBase=[Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($regISKeyType,$server)
$regISRegKey = $regISRemoteBase.OpenSubKey($regISKey)
if ($regISRegKey -ne $null)
{
	$regISTcpIpPort = $regISRegKey.getvalue("TCP/IP Port")
	$regISPolling = $regISRegKey.getvalue("Maximum Polling Interval")
    $regISOlmChecksum = $regISRegKey.getvalue("Online Maintenance Checksum")
}

$regVirusKey = "SYSTEM\CurrentControlSet\Services\MSExchangeIS\VirusScan"
$regVirusKeyType = [Microsoft.Win32.RegistryHive]::LocalMachine
$regVirusRemoteBase=[Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($regVirusKeyType,$server)
$regVirusRegKey = $regVirusRemoteBase.OpenSubKey($regVirusKey)
if ($regVirusRegKey -ne $null)
{
	$regVirusEnabled = $regVirusRegKey.getvalue("Enabled")
	$regVirusVendor = $regVirusRegKey.getvalue("Vendor")
    $regVirusVersion = $regVirusRegKey.getvalue("Version")
}

# MSExchangeRPC Queries
$regRPCKey = "SYSTEM\CurrentControlSet\Services\MSExchangeRPC\ParametersSystem"
$regRPCKeyType = [Microsoft.Win32.RegistryHive]::LocalMachine
$regRPCRemoteBase=[Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($regRPCKeyType,$server)
$regRPCRegKey = $regRPCRemoteBase.OpenSubKey($regRPCKey)
if ($regRPCRegKey -ne $null)
{
	$regRPCTcpIpPort = $regRPCRegKey.getvalue("TCP/IP Port")
	$regRPCPush = $regRPCRegKey.getvalue("EnablePushNotifications")
}

# MSExchangeAB Queries
$regABKey = "SYSTEM\CurrentControlSet\Services\MSExchangeAB\Parameters"
$regABKeyType = [Microsoft.Win32.RegistryHive]::LocalMachine
$regABRemoteBase=[Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($regABKeyType,$server)
$regABRegKey = $regABRemoteBase.OpenSubKey($regABKey)
if ($regABRegKey -ne $null)
{
	$regABRpcTcpPort = $regABRegKey.getvalue("RpcTcpPort")
}

# RESvc Queries
$regREKey = "SYSTEM\CurrentControlSet\Services\RESvc\Parameters"
$regREKeyType = [Microsoft.Win32.RegistryHive]::LocalMachine
$regRERemoteBase=[Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($regREKeyType,$server)
$regRERegKey = $regRERemoteBase.OpenSubKey($regREKey)
if ($regRERegKey -ne $null)
{
	$regRESuppressStateChanges = $regRERegKey.getvalue("SuppressStateChanges")
}

$output_Exch_Reg_Ex = $server + "`t" + `
	$regMinUserDC + "`t" + `
	$regVirusEnabled  + "`t" + `
	$regVirusVendor  + "`t" + `
    $regVirusVersion  + "`t" + `
	$regSATcpIpPort + "`t" + `
	$regSATcpIpNspiPort + "`t" + `
	$regSANspiTargetServerPort + "`t" + `
	$regSARfrTargetServerPort + "`t" + `
	$regISTcpIpPort + "`t" + `
	$regISPolling + "`t" + `
    $regISOlmChecksum + "`t" + `
	$regRPCTcpIpPort + "`t" + `
	$regRPCPush + "`t" + `
	$regABRpcTcpPort + "`t" + `
	$regRESuppressStateChanges
$output_Exch_Reg_Ex | Out-File -FilePath $Exch_Reg_Ex_outputfile -append 

$a = Get-Process -pid $PID

$EventText = "Exch_Reg_Ex " + "`n" + $server + "`n"
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
