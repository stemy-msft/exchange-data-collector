#############################################################################
#                          Core_Build_Nodes		 							#
#                                     			 							#
#                               4.0.2    		 							#
#                                     			 							#
#############################################################################
# don't forget...                     			 							#
#   set-executionpolicy unrestricted  			 							#
#                                     			 							#
# Requires:                           			 							#
# 	Exchange Management Shell (or PowerShell)  	 							#
# See ExDC_Instructions.txt for more info									#
#                                     			 							#
# Issues, comments, or suggestions    			 							#
#   mail stemy@microsoft.com          			 							#
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

Param($location,$session_0)

Write-Output -InputObject $PID
Write-Output -InputObject "Core"

set-location -LiteralPath $location

$ErrorActionPreference = "Stop"
Trap {
$ErrorText = "Core_BuildFiles " + "`n" + $server + "`n"
$ErrorText += $_

$ErrorLog = New-Object -TypeName System.Diagnostics.EventLog -ArgumentList Application
$ErrorLog.MachineName = "."
$ErrorLog.Source = "ExDC"
$ErrorLog.WriteEntry($ErrorText,"Error", 100)
}

$error.clear()

###########################
# Output Filenames
###########################
$Servers_outputfile = ".\ClusterNodes.txt"
	
###########################
# Add Exchange Snap-in
###########################

$RegisteredSnapins = Get-PSSnapin -Registered
foreach ($Snapin in $RegisteredSnapins)
{
	if ($Snapin.name -like "Microsoft.Exchange.Management.PowerShell.Admin")
	{
		# Ignoring $session_0, if present, to use Exchange 2007 Powershell
		Add-PSSnapin -Name Microsoft.Exchange.Management.PowerShell.Admin
        ([Microsoft.Exchange.Data.Directory.AdminSessionADSettings]::Instance).ViewEntireForest = $true
		$session_0 = $null
	}
	elseif ($session_0 -ne $null)
	{
		$cxnUri = "http://" + $session_0 + "/powershell"
		$session = New-PSSession -configurationName Microsoft.Exchange -ConnectionUri $cxnUri -authentication Kerberos
		Import-PSSession -Session $session -AllowClobber -CommandName Get-ExchangeServer,Get-MailboxServer,Set-AdServerSettings
        Set-AdServerSettings -ViewEntireForest $true
	}
	
}

###########################
# Initialize Arrays
###########################
[array]$arrayClusterNodes = @()

###########################
# Populate Exchange Nodes in Forest
###########################
$objExchDomain = New-Object -TypeName System.DirectoryServices.DirectoryEntry -ArgumentList ("LDAP://RootDSE")
$objExchSchemaVersion = New-Object -TypeName System.DirectoryServices.DirectoryEntry -ArgumentList ("LDAP://CN=ms-Exch-Schema-Version-Pt," + $objExchDomain.schemaNamingContext)
$error.clear()

If ($objExchSchemaVersion.rangeUpper -ge 10628)
{
#	Write-Host -Object "E2k7 or later environment detected"
#	Write-Host -Object "ms-Exch-Schema-Version is " $objExchSchemaVersion.rangeUpper
	$EventLog = New-Object -TypeNameSystem.Diagnostics.EventLog -ArgumentList Application
	$EventLog.MachineName = "."
	$EventLog.Source = "ExDC"
	$EventText = "ms-Exch-Schema-Version is " + $objExchSchemaVersion.rangeUpper
	$EventLog.WriteEntry($EventText,"Information", 12)
	
	$colGES = @(Get-ExchangeServer)
	foreach ($objGES in $colGES)
	{
		if ($objGES.IsExchange2007orLater -eq $false)
		{
			$ExchangeServer = $objGES.name
			$strExchFilter = "(&(objectclass=msExchExchangeServer)((cn=$ExchangeServer))(!(objectclass=msExchExchangeServerPolicy))(!(objectclass=msExchClientAccessArray)))"
			$objExchDomain = New-Object -TypeName System.DirectoryServices.DirectoryEntry -ArgumentList ("LDAP://RootDSE")
			$objExchConfig = New-Object -TypeName System.DirectoryServices.DirectoryEntry -ArgumentList ("LDAP://CN=Microsoft Exchange,CN=Services," + $objExchDomain.configurationNamingContext)
			$objExchSearcher = New-Object -TypeName System.DirectoryServices.DirectorySearcher
			$objExchSearcher.SearchRoot = $objExchConfig
			$objExchSearcher.PageSize = 1000
			$objExchSearcher.Filter = $strExchFilter
			$objExchSearcher.SearchScope = "Subtree"
			$colExchProplist = "cn", "heuristics"
			foreach ($objExchPropList in $colExchProplist)
			{
				$objExchSearcher.PropertiesToLoad.Add($objExchPropList)
			}
			$colExchResults = $objExchSearcher.FindAll()
			foreach ($objExchResult in $colExchResults)
			{
				if ($objExchResult.properties.heuristics -eq 1579012)
				{
					# Exchange 2003 is clustered
					###########################
					# WMI Queries
					# MSCluster_Node
					###########################
					[string]$strExchCN = $objExchResult.Properties.cn
					$col_LegacyExch = get-wmiobject -class "MSCluster_node" -namespace "root/MSCluster" -comp $strExchCN -ErrorAction SilentlyContinue
					foreach ($obj_LegacyExch in $col_LegacyExch)
					{
						if ($obj_LegacyExch.name -ne $null)
						#Exchange server is clustered
						{
							$ClusterAndNodeName = $strExchCN + "~~" + $obj_LegacyExch.Name
							$arrayClusterNodes += $ClusterAndNodeName 
						}
					}	
				}
			}
		}
		elseif ($session_0 -eq $null) # This is an Exchange 2007 Powershell
		{
			if (($objGES.IsMemberOfCluster -eq "yes"))
			#This should catch all Clustered E2k7 servers
			#Get-MailboxServer should list all nodes as RedundantMachines
			{
				$objGMBS = Get-MailboxServer $objGES.name
				foreach ($machine in $objGMBS.RedundantMachines)
				{
					$ClusterAndNodeName = $objGES.name + "~~" + $machine
					$arrayClusterNodes += $ClusterAndNodeName
				}
			}
		}
	}
}
else
{
	$EventLog = New-Object -TypeNameSystem.Diagnostics.EventLog -ArgumentList Application
	$EventLog.MachineName = "."
	$EventLog.Source = "ExDC"
	$EventText = "ms-Exch-Schema-Version is " + $objExchSchemaVersion.rangeUpper + "`nPre-Exchange 2007 environments do not support the Exchange cmdlets."
	$EventLog.WriteEntry($EventText,"Information", 12)

	###########################
	# Find Exchange servers in Config
	###########################
	$strExchFilter = "(&(objectclass=msExchExchangeServer)(!(objectclass=msExchExchangeServerPolicy))(!(objectclass=msExchClientAccessArray)))"
	$objExchDomain = New-Object -TypeName System.DirectoryServices.DirectoryEntry -ArgumentList ("LDAP://RootDSE")
	$objExchConfig = New-Object -TypeName System.DirectoryServices.DirectoryEntry -ArgumentList ("LDAP://CN=Microsoft Exchange,CN=Services," + $objExchDomain.configurationNamingContext)
	$objExchSearcher = New-Object -TypeNameSystem.DirectoryServices.DirectorySearcher
	$objExchSearcher.SearchRoot = $objExchConfig
	$objExchSearcher.PageSize = 1000
	$objExchSearcher.Filter = $strExchFilter
	$objExchSearcher.SearchScope = "Subtree"
	$colExchProplist = "cn", "heuristics"
	foreach ($objExchPropList in $colExchProplist)
	{
		$objExchSearcher.PropertiesToLoad.Add($objExchPropList)
	}
	
	$colExchResults = $objExchSearcher.FindAll()

	foreach ($objExchResult in $colExchResults)
	{
		if ($objExchResult.properties.heuristics -eq 1579012)
		{
			# Exchange 2003 is clustered
			###########################
			# WMI Queries
			# MSCluster_Node
			###########################
			[string]$strExchCN = $objExchResult.Properties.cn
			$col_LegacyExch = get-wmiobject -class "MSCluster_node" -namespace "root/MSCluster" -comp $strExchCN -ErrorAction SilentlyContinue
			foreach ($obj_LegacyExch in $col_LegacyExch)
			{
				if ($obj_LegacyExch.name -ne $null)
				#Exchange server is clustered
				{
					$ClusterAndNodeName = $strExchCN + "~~" + $obj_LegacyExch.Name 
					$arrayClusterNodes += $ClusterAndNodeName
				}
			}	
		}
	}

}

# Output each array to text files with dupes removed
foreach ($server in $arrayClusterNodes | Sort-Object -u) 
	{
		$server | Out-File -FilePath $Servers_outputfile -append
	}
if ($arrayClusterNodes.count -lt 1)
{
	"" | Out-File -filepath $Servers_outputfile -append
}

	
if ($session_0 -ne $null)
{
	Remove-PSSession -Session $session
}
