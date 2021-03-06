#############################################################################
#                   ExOrg_GetExchangeCertificate.ps1		 				#
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
Param($server,$location,$append,$session_0)

Write-Output -InputObject $PID
Write-Output -InputObject "ExOrg"

$ErrorActionPreference = "Stop"
Trap {
$ErrorText = "ExOrg_GetExchangeCertificate " + "`n" + $server + "`n"
$ErrorText += $_ 
$ErrorText += $i

$ErrorLog = New-Object -TypeName System.Diagnostics.EventLog -ArgumentList Application
$ErrorLog.MachineName = "."
$ErrorLog.Source = "ExDC"
$ErrorLog.WriteEntry($ErrorText,"Error", 100)
}

set-location -LiteralPath $location
$output_location = $location + "\output\ExOrg\GetExchCert"

if ((Test-Path -LiteralPath $output_location) -eq $false)
    {New-Item -Path $output_location -ItemType directory -Force}

function GetTypeListFromXmlFile( [string] $typeFileName ) 
{
	$xmldata = [xml](Get-Content $typeFileName)
	$returnList = $xmldata.Types.Type | where { ($_.Name.StartsWith("Microsoft.Exchange") -and !$_.Name.Contains("[[")) } | foreach { $_.Name }
	return $returnList
}

if ($session_0 -ne $null)
{
	$cxnUri = "http://" + $session_0 + "/powershell"
	$session = New-PSSession -configurationName Microsoft.Exchange -ConnectionUri $cxnUri -authentication Kerberos
	Import-PSSession -Session $session -AllowClobber -CommandName Get-ExchangeCertificate,Set-AdServerSettings
    Set-AdServerSettings -ViewEntireForest $true

    if (Test-Path -Path HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup)
    {
        $global:exbin = (Get-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup).MsiInstallPath + "bin\"

        ## LOAD CONNECTION FUNCTIONS #################################################

        # ConnectFunctions.ps1 uses some of the Exchange types. PowerShell does some type binding at the 
        # time of loading the scripts, so we'd rather load the scripts before we reference those types.
        "Microsoft.Exchange.Data.dll", "Microsoft.Exchange.Configuration.ObjectModel.dll" `
          | ForEach { [System.Reflection.Assembly]::LoadFrom((join-path $global:exbin $_)) } `
          | Out-Null

        ## LOAD EXCHANGE EXTENDED TYPE INFORMATION ###################################

        $FormatEnumerationLimit = 16

        # loads powershell types file, parses out just the type names and returns an array of string
        # it skips all template types as template parameter types individually are defined in types file

        # Check if every single type from from Exchange.Types.ps1xml can be successfully loaded
        $typeFilePath = join-path $global:exbin "exchange.types.ps1xml"
        $typeListToCheck = GetTypeListFromXmlFile $typeFilePath
        # Load all management cmdlet related types.
        $assemblyNames = [Microsoft.Exchange.Configuration.Tasks.CmdletAssemblyHelper]::ManagementCmdletAssemblyNames
        $typeLoadResult = [Microsoft.Exchange.Configuration.Tasks.CmdletAssemblyHelper]::EnsureTargetTypesLoaded($assemblyNames, $typeListToCheck)
        # $typeListToCheck is a big list, release it to free up some memory
        $typeListToCheck = $null

        $SupportPath = join-path $global:exbin "Microsoft.Exchange.Management.Powershell.Support.dll"
        [Microsoft.Exchange.Configuration.Tasks.TaskHelper]::LoadExchangeAssemblyAndReferences($SupportPath) > $null

        if (Get-ItemProperty HKLM:\Software\microsoft\ExchangeServer\v15\CentralAdmin -ea silentlycontinue)
        {
            $CentralAdminPath = join-path $global:exbin "Microsoft.Exchange.Management.Powershell.CentralAdmin.dll"
            [Microsoft.Exchange.Configuration.Tasks.TaskHelper]::LoadExchangeAssemblyAndReferences($CentralAdminPath) > $null
        }

        # Register Assembly Resolver to handle generic types
        [Microsoft.Exchange.Data.SerializationTypeConverter]::RegisterAssemblyResolver()

        # Finally, load the types information
        # We will load type information only if every single type from Exchange.Types.ps1xml can be successfully loaded
        if ($typeLoadResult)
        {
	        Update-TypeData -PrependPath $typeFilePath
        }
        else
        {
	        write-error $RemoteExchange_LocalizedStrings.res_types_file_not_loaded
        }

        #load partial types
        $partialTypeFile = join-path $global:exbin "Exchange.partial.Types.ps1xml"
        Update-TypeData -PrependPath $partialTypeFile 

    }
    else
    {
        ## EXCHANGE VARIABLEs ########################################################

        $global:exbin = (get-itemproperty HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup).MsiInstallPath + "bin\"

        ## LOAD EXCHANGE EXTENDED TYPE INFORMATION ###################################

        # loads powershell types file, parses out just the type names and returns an array of string
        # it skips all template types as template parameter types individually are defined in types file

        $ConfigurationPath = join-path $global:exbin "Microsoft.Exchange.Configuration.ObjectModel.dll"
        [System.Reflection.Assembly]::LoadFrom($ConfigurationPath) > $null

        # Check if every single type from from Exchange.Types.ps1xml can be successfully loaded
        $ManagementPath = join-path $global:exbin "Microsoft.Exchange.Management.dll"
        $typeFilePath = join-path $global:exbin "exchange.types.ps1xml"
        $typeListToCheck = GetTypeListFromXmlFile $typeFilePath
        $typeLoadResult = [Microsoft.Exchange.Configuration.Tasks.TaskHelper]::TryLoadExchangeTypes($ManagementPath, $typeListToCheck)
        # $typeListToCheck is a big list, release it to free up some memory
        $typeListToCheck = $null

        # Register Assembly Resolver to handle generic types
        [Microsoft.Exchange.Data.SerializationTypeConverter]::RegisterAssemblyResolver()

        #load partial types
        $partialTypeFile = join-path $global:exbin "Exchange.partial.Types.ps1xml"
        Update-TypeData -PrependPath $partialTypeFile 
    }
    $ExOrg_GetExchangeCertificate_outputfile = $output_location + "\" + $server + "~~GetExchCert.txt"

    @(Get-ExchangeCertificate -server $server) | foreach `
    {
	    $output_ExOrg_GetExchangeCertificate = $server + "`t" + `
		    $_.Status + "`t" + `
		    $_.IsSelfSigned + "`t" + `
		    $_.RootCAType + "`t" + `
		    $_.PublicKeySize + "`t" + `
		    $_.Thumbprint + "`t" + `
		    $_.Services + "`t" + `
		    $_.Subject + "`t" + `
		    $_.Issuer + "`t" + `
		    $_.NotBefore + "`t" + `
		    $_.NotAfter + "`t" + `
		    $_.CertificateDomains 

	    $output_ExOrg_GetExchangeCertificate | out-file $ExOrg_GetExchangeCertificate_outputfile -append 
    }
}

$a = Get-Process -pid $PID

$EventText = "ExOrg_GetExchangeCertificate " + "`n" + $server + "`n"
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
