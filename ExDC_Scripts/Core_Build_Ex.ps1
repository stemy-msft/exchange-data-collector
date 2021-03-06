#############################################################################
#                          Core_Build_Ex		 							#
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

Param($location)

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
$Exchange_outputfile = $location.path + "\Exchange.txt"
$Schema_output_location = $location.path + "\output\ExOrg"

if ((Test-Path -LiteralPath $Schema_output_location) -eq $false)
    {New-Item -Path $Schema_output_location -ItemType directory -Force}

$Schema_outputfile = $Schema_output_location + "\ExOrg_GetSchema.txt"

###########################
# Populate Exchange servers in Forest
###########################
$status_Step1.Text = "Step 1 Status: Running - Collecting list of Exchange servers"

$objExchDomain = New-Object -TypeName System.DirectoryServices.DirectoryEntry -ArgumentList ("LDAP://RootDSE")
$objExchSchemaVersion = New-Object -TypeName system.DirectoryServices.DirectoryEntry -ArgumentList ("LDAP://CN=ms-Exch-Schema-Version-Pt," + $objExchDomain.schemaNamingContext)
$RangeUpper = $objExchSchemaVersion.rangeUpper
if ((Test-Path -LiteralPath $Schema_outputfile) -eq $false)
{
	$RangeUpper | Out-File -FilePath $Schema_outputfile -Force
}
$error.clear()

if ((Test-Path -LiteralPath $Exchange_outputfile) -eq $true)
{
	Write-Host -Object "Exchange.txt is already present in this folder." -ForegroundColor Red
	Write-Host -Object "Loading values from text file that is present." -ForegroundColor Red
    $status_Step1.Text = "Step 1 Status: Failed - Exchange.txt is already present. Loading values from existing file."
	exit
}
	
###########################
# Initialize Arrays
###########################
$arrayExchange = @()



# rangeUpper values
# Ex2016 RTM = 15317
# Ex2013 RTM = 15137
# Ex2010 RTM = 14622
# Ex2007 RTM = 10628
# Ex2003 RTM = 6870

If ($rangeUpper -lt 10628)
{
    Write-Host -Object "Pre-Exchange 2007 environments do not support the Exchange cmdlets."
    $status_Step1.Text = "Step 1 Status: Running with Warning - Pre-Exchange 2007 environments do not support the Exchange cmdlets."
}
elseIf (($rangeUpper -ge 10628) -and ($RangeUpper -lt 14622))
{
	Write-Host -Object "Ex2007 environment detected"
	Write-Host -Object "ms-Exch-Schema-Version is $rangeUpper "
}
elseIf (($rangeUpper -ge 14622) -and ($RangeUpper -lt 15137))
{
	Write-Host -Object "Ex2010 environment detected"
	Write-Host -Object "ms-Exch-Schema-Version is $rangeUpper "
}
elseIf (($rangeUpper -ge 15137) -and ($RangeUpper -lt 15317))
{
	Write-Host -Object "Ex2013 environment detected"
	Write-Host -Object "ms-Exch-Schema-Version is $rangeUpper "
}
elseIf ($rangeUpper -ge 15317)
{
	Write-Host -Object "Ex2016 or later environment detected"
	Write-Host -Object "ms-Exch-Schema-Version is $rangeUpper "
}
$status_Step1.Text = "Step 1 Status: Running - Collecting list of Exchange servers.  Schema version $rangeUpper detected."


# We're going to look for all msExchExchangeServer objects in the config container
$strExchFilter = "(&(objectclass=msExchExchangeServer)(!(objectclass=msExchExchangeTransportServer))(!(objectclass=msExchExchangeServerPolicy))(!(objectclass=msExchClientAccessArray)))"
$objExchDomain = New-Object -TypeName System.DirectoryServices.DirectoryEntry -ArgumentList ("LDAP://RootDSE")
$objExchConfig = New-Object -TypeName System.DirectoryServices.DirectoryEntry -ArgumentList ("LDAP://CN=Microsoft Exchange,CN=Services," + $objExchDomain.configurationNamingContext)
$objExchSearcher = New-Object -TypeName System.DirectoryServices.DirectorySearcher
$objExchSearcher.SearchRoot = $objExchConfig
$objExchSearcher.PageSize = 1000
$objExchSearcher.Filter = $strExchFilter
$objExchSearcher.SearchScope = "Subtree"

### Minimize Property list
$colProplist = "networkAddress"
foreach ($objPropList in $colPropList)
{
	$objExchSearcher.PropertiesToLoad.Add($objPropList) | Out-Null
}

$colResults = $objExchSearcher.FindAll()
$arrayExchange = @()
foreach ($objResult in $colResults)
{
	$objItem = $objResult.Properties
	foreach( $na in $objItem.networkaddress )
    {
        if( $na -match 'ncacn_ip_tcp:(.+)' )
        {
            $arrayExchange +=  $Matches.item(1)
            break
        }
    }
}


# Output each array to text files with dupes removed
$status_Step1.Text = "Step 1 Status: Running - Sorting arrays and writing output files."	
foreach ($Exchange in $arrayExchange | Sort-Object -u) 
	{
		$Exchange | Out-File -FilePath $Exchange_outputfile -append
	}
    
$status_Step1.Text = "Step 1 Status: Idle"
