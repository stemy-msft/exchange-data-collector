#############################################################################
#                             DC_Reg_OS.ps1		 							#
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
$ErrorText = "DC_Reg_OS " + "`n" + $server + "`n"
$ErrorText += $_

$ErrorLog = New-Object -TypeName System.Diagnostics.EventLog -ArgumentList Application
$ErrorLog.MachineName = "."
$ErrorLog.Source = "ExDC"
$ErrorLog.WriteEntry($ErrorText,"Error", 100)
}

set-location -LiteralPath $location
$output_location = $location + "\output\DC_Reg_OS"

if ((Test-Path -LiteralPath $output_location) -eq $false)
    {New-Item -Path $output_location -ItemType directory -Force}

$DC_Reg_OS_outputfile = $output_location + "\" +$server + "~~DC_Reg_OS.txt"

# Time Queries
$regW32TimeKey = "SYSTEM\CurrentControlSet\Services\W32Time\Parameters"
$regW32TimeKeyType = [Microsoft.Win32.RegistryHive]::LocalMachine
$regW32TimeRemoteBase=[Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($regW32TimeKeyType,$server)
$regW32TimeRegKey = $regW32TimeRemoteBase.OpenSubKey($regW32TimeKey)
$regW32TimeTypeValue = $regW32TimeRegKey.getvalue("Type")
$regW32TimeNTPServerValue = $regW32TimeRegKey.getvalue("NTPServer")

$regW32TimeConfigKey = "SYSTEM\CurrentControlSet\Services\W32Time\Config"
$regW32TimeConfigKeyType = [Microsoft.Win32.RegistryHive]::LocalMachine
$regW32TimeConfigRemoteBase=[Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($regW32TimeConfigKeyType,$server)
$regW32TimeConfigRegKey = $regW32TimeConfigRemoteBase.OpenSubKey($regW32TimeConfigKey)
$regW32TimeConfigTypeValue = $regW32TimeConfigRegKey.getvalue("Type")
$regW32MaxNegPhase = $regW32TimeConfigRegKey.getvalue("MaxNegPhaseCorrection")
$regW32MaxPosPhase = $regW32TimeConfigRegKey.getvalue("MaxPosPhaseCorrection")
					
# Tcpip Queries
$regTcpipKey = "SYSTEM\CurrentControlSet\Services\Tcpip\Parameters"
$regTcpipKeyType = [Microsoft.Win32.RegistryHive]::LocalMachine
$regTcpipRemoteBase=[Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($regTcpipKeyType,$server)
$regTcpipRegKey = $regTcpipRemoteBase.OpenSubKey($regTcpipKey)
$regTcpipRSSValue = $regTcpipRegKey.getvalue("EnableRSS")
$regTcpipTCPAValue = $regTcpipRegKey.getvalue("EnableTCPA")
$regTcpipTCPChimneyValue = $regTcpipRegKey.getvalue("EnableTCPChimney")

# Tcpip v6 Queries
$regTcpip6Key = "SYSTEM\CurrentControlSet\Services\Tcpip6\Parameters"
$regTcpip6KeyType = [Microsoft.Win32.RegistryHive]::LocalMachine
$regTcpip6RemoteBase=[Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($regTcpip6KeyType,$server)
$regTcpip6RegKey = $regTcpip6RemoteBase.OpenSubKey($regTcpip6Key)
if ($regTcpip6RegKey -ne $null)
{
	$regTcpip6DisabledValue = $regTcpip6RegKey.getvalue("DisabledComponents")
}
	# 0 = enable all IPv6 components (default)
	# 0xffffffff = disable all IPv6 except IPv6 loopback and prefer IPv4 over IPv6
	# 0x20 = prefer IPv4 over IPv6
	# 0x10 = disable IPv6 on all nontunnel interfaces
	# 0x01 = disable IPv6 on all tunnel interfaces
	# 0x11 = disable all IPv6 except IPv6 loopback
	
# Environmental Variables Queries
$regEnvironmentKey = "SYSTEM\CurrentControlSet\Control\Session Manager\Environment"
$regEnvironmentKeyType = [Microsoft.Win32.RegistryHive]::LocalMachine
$regEnvironmentRemoteBase=[Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($regEnvironmentKeyType,$server)
$regEnvironmentRegKey = $regEnvironmentRemoteBase.OpenSubKey($regEnvironmentKey)
$regEnvironmentTEMPValue = $regEnvironmentRegKey.getvalue("TEMP")
$regEnvironmentTMPValue = $regEnvironmentRegKey.getvalue("TMP")

# LSA Queries
$regLsaKey = "SYSTEM\CurrentControlSet\Control\LSA"
$regLsaKeyType = [Microsoft.Win32.RegistryHive]::LocalMachine
$regLsaRemoteBase=[Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($regLsaKeyType,$server)
$regLsaRegKey = $regLsaRemoteBase.OpenSubKey($regLsaKey)
if ($regLsaRegKey -ne $null)
{
	$regRestrictAnonName = $regLsaRegKey.getvalue("RestrictAnonymous")
	$regRestrictAnonSamName = $regLsaRegKey.getvalue("RestrictAnonymousSam")
	$regLmCompatName = $regLsaRegKey.getvalue("LMCompatibilityLevel")
}

# NTLM Compatability Queries
$regNtlmKey = "SYSTEM\CurrentControlSet\Control\LSA\MSV1_0"
$regNtlmKeyType = [Microsoft.Win32.RegistryHive]::LocalMachine
$regNtlmRemoteBase=[Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($regNtlmKeyType,$server)
$regNtlmRegKey = $regNtlmRemoteBase.OpenSubKey($regNtlmKey)
if ($regNtlmRegKey -ne $null)
{
	$regNtlmMinClientName = $regNtlmRegKey.getvalue("NtlmMinClientSec")
	$regNtlmMinClientName_Hex = [String]::Format("{0:X}", $regNtlmMinClientName) 
	$regNtlmMinServerName = $regNtlmRegKey.getvalue("NtlmMinServerSec")
	$regNtlmMinServerName_Hex = [String]::Format("{0:X}", $regNtlmMinServerName) 
}
	# 0x00000010 - Message Integrity
	# 0x00000020 - Message Confidentiality
	# 0x00080000 - NTLM v2 Session Security
	# 0x20000000 - Confidentiality 128-bit
	# 0x80000000 - Confidentiality 56-bit

$output_DC_Reg_OS = $server + "`t" + `
	$regW32TimeTypeValue + "`t" + `
	$regW32TimeNTPServerValue + "`t" + `
	$regW32MaxNegPhase + "`t" + `
	$regW32MaxPosPhase + "`t" + `
	$regTcpipRSSValue + "`t" + `
	$regTcpipTCPAValue + "`t" + `
	$regTcpipTCPChimneyValue + "`t" + `
	$regTcpip6DisabledValue + "`t" + `
	$regEnvironmentTEMPValue + "`t" + `
	$regEnvironmentTMPValue + "`t" + `
	$regRestrictAnonName + "`t" + `
	$regRestrictAnonSamName + "`t" + `
	$regNtlmMinClientName_Hex + "`t" + `
	$regNtlmMinServerName_Hex
$output_DC_Reg_OS | Out-File -FilePath  $DC_Reg_OS_outputfile -append 

$a = Get-Process -pid $PID

$EventText = "DC_Reg_OS " + "`n" + $server + "`n"
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

