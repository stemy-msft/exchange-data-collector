#############################################################################
#                    Core_Assemble_ExOrg_Excel.ps1		 					#
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
Param($RunLocation)

$ErrorActionPreference = "Stop"
Trap {
$ErrorText = "Core_Assemble_ExOrg_Excel " + "`n" + $server + "`n"
$ErrorText += $_

$ErrorLog = New-Object -TypeName System.Diagnostics.EventLog -ArgumentList Application
$ErrorLog.MachineName = "."
$ErrorLog.Source = "ExDC"
#$ErrorLog.WriteEntry($ErrorText,"Error", 100)
}

function Process-Datafile{
    param ([int]$NumberOfColumns, `
			[array]$DataFromFile, `
			$Wsheet, `
			[int]$ExcelVersion) 
		$RowCount = $DataFromFile.Count
        $ArrayRow = 0
        $BadArrayValue = @()
        $DataArray = New-Object 'object[,]' -ArgumentList $RowCount,$NumberOfColumns
		Foreach ($DataRow in $DataFromFile)
        {
            $DataField = $DataRow.Split("`t")
            for ($ArrayColumn = 0 ; $ArrayColumn -lt $NumberOfColumns ; $ArrayColumn++)
            {
                # Excel chokes if field starts with = so we'll try to prepend the ' to the string if it does
                Try{If ($DataField[$ArrayColumn].substring(0,1) -eq "=") {$DataField[$ArrayColumn] = "'"+$DataField[$ArrayColumn]}}
				Catch{}
                # Excel 2003 limit of 1823 characters
                if ($DataField[$ArrayColumn].length -lt 1823) 
                    {$DataArray[$ArrayRow,$ArrayColumn] = $DataField[$ArrayColumn]}
                # Excel 2007 limit of 8203 characters
                elseif (($ExcelVersion -ge 12) -and ($DataField[$ArrayColumn].length -lt 8203)) 
                    {$DataArray[$ArrayRow,$ArrayColumn] = $DataField[$ArrayColumn]} 
                # No known Excel 2010 limit
                elseif ($ExcelVersion -ge 14)
                    {$DataArray[$ArrayRow,$ArrayColumn] = $DataField[$ArrayColumn]}
                else
                {
                    Write-Host -Object "Number of characters in array member exceeds the version of limitations of this version of Excel" -ForegroundColor Yellow
                    Write-Host -Object "-- Writing value to temp variable" -ForegroundColor Yellow
                    $DataArray[$ArrayRow,$ArrayColumn] = $DataField[$ArrayColumn]
                    $BadArrayValue += "$ArrayRow,$ArrayColumn"
                }
            }
            $ArrayRow++
        }
   
        # Replace big values in $DataArray
        $BadArrayValue_count = $BadArrayValue.count
        $BadArrayValue_Temp = @()
        for ($i = 0 ; $i -lt $BadArrayValue_count ; $i++)
        {
            $BadArray_Split = $badarrayvalue[$i].Split(",")
            $BadArrayValue_Temp += $DataArray[$BadArray_Split[0],$BadArray_Split[1]]
            $DataArray[$BadArray_Split[0],$BadArray_Split[1]] = "**TEMP**"
            Write-Host -Object "-- Replacing long value with **TEMP**" -ForegroundColor Yellow
        }
    
        $EndCellRow = ($RowCount+1)
        $Data_range = $Wsheet.Range("a2","$EndCellColumn$EndCellRow")
        $Data_range.Value2 = $DataArray
    
        # Paste big values back into the spreadsheet
        for ($i = 0 ; $i -lt $BadArrayValue_count ; $i++)
        {
            $BadArray_Split = $badarrayvalue[$i].Split(",")
            # Adjust for header and $i=0
            $CellRow = [int]$BadArray_Split[0] + 2
            # Adjust for $i=0
            $CellColumn = [int]$BadArray_Split[1] + 1
               
            $Range = $Wsheet.cells.item($CellRow,$CellColumn) 
            $Range.Value2 = $BadArrayValue_Temp[$i]
            Write-Host -Object "-- Pasting long value back in spreadsheet" -ForegroundColor Yellow
        }    
    }

set-location -LiteralPath $RunLocation

	$EventLog = New-Object -TypeName System.Diagnostics.EventLog -ArgumentList Application
	$EventLog.MachineName = "."
	$EventLog.Source = "ExDC"
	#$EventLog.WriteEntry("Starting Core_Assemble_ExOrg_Excel","Information", 42)	

Write-Host -Object "---- Starting to create com object for Excel"
$Excel_ExOrg = New-Object -ComObject excel.application
Write-Host -Object "---- Hiding Excel"
$Excel_ExOrg.visible = $false
Write-Host -Object "---- Setting ShowStartupDialog to false" 
$Excel_ExOrg.ShowStartupDialog = $false 
Write-Host -Object "---- Setting DefaultFilePath"
$Excel_ExOrg.DefaultFilePath = $RunLocation + "\output"
Write-Host -Object "---- Setting SheetsInNewWorkbook" 
$Excel_ExOrg.SheetsInNewWorkbook = 64
Write-Host -Object "---- Checking Excel version"
$Excel_Version = $Excel_ExOrg.version
if ($Excel_version -ge 12)
{
	$Excel_ExOrg.DefaultSaveFormat = 51
	$excel_Extension = ".xlsx"
}
else
{
	$Excel_ExOrg.DefaultSaveFormat = 56
	$excel_Extension = ".xls"
}
Write-Host -Object "---- Excel version $Excel_version and DefaultSaveFormat $Excel_extension"

# Create new Excel workbook
Write-Host -Object "---- Adding workbook"
$Excel_ExOrg_workbook = $Excel_ExOrg.workbooks.add()
Write-Host -Object "---- Setting output file"
$ExDC_ExOrg_XLS = $RunLocation + "\output\ExDC_ExOrg" + $excel_Extension

Write-Host -Object "---- Setting workbook properties"
$Excel_ExOrg_workbook.author = "Exchange Data Collector v4 (ExDC v4)"
$Excel_ExOrg_workbook.title = "ExDC v4 - Exchange Organization"
$Excel_ExOrg_workbook.comments = "ExDC v4.0.1"

$intSheetCount = 1
$intColorIndex_ClientAccess = 45
$intColorIndex_Global = 11
$intColorIndex_Recipient = 45
$intColorIndex_Transport = 11
$intColorIndex_Um = 45
$intColorIndex_Misc = 11


# Client Access
#Region Get-ActiveSyncDevice sheet
Write-Host -Object "---- Starting Get-ActiveSyncDevice"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "ActiveSyncDevice"
	$Worksheet.Tab.ColorIndex = $intColorIndex_ClientAccess
	$row = 1
	$header = @()
	$header +=  "FriendlyName"
	$header +=  "DeviceMobileOperator"
	$header +=  "DeviceOS"
	$header +=  "DeviceTelephoneNumber"
	$header +=  "DeviceType"
	$header +=  "DeviceUserAgent"
	$header +=  "DeviceModel"
	$header +=  "FirstSyncTime"		# Column H
	$header +=  "UserDisplayName"
	$header +=  "DeviceAccessState"
	$header +=  "DeviceAccessStateReason"
	$header +=  "DeviceActiveSyncVersion"
	$header +=  "Name"
	$header +=  "Identity"
    $a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\ExOrg_GetActiveSyncDevice.txt") -eq $true)
{	
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\ExOrg_GetActiveSyncDevice.txt") 
	
	# Send the data to the function to process and add to the Excel worksheet
	Process-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}	

# Format time/date columns
$EndRow = $DataFile.count + 1
# FirstSyncTime
$Column_Range = $Worksheet.Range("H1","H$EndRow")
$Column_Range.cells.NumberFormat = "mm/dd/yy hh:mm:ss"

	#EndRegion Get-ActiveSyncDevice sheet
		
#Region Get-ActiveSyncMailboxPolicy sheet
Write-Host -Object "---- Starting Get-ActiveSyncMailboxPolicy"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "ActiveSyncMailboxPolicy"
	$Worksheet.Tab.ColorIndex = $intColorIndex_ClientAccess
	$row = 1
	$header = @()
	$header +=  "Name"
	$header +=  "AllowNonProvisionableDevices"
	$header +=  "AlphanumericPasswordRequired"
	$header +=  "AttachmentsEnabled"
	$header +=  "DeviceEncryptionEnabled"
	$header +=  "RequireStorageCardEncryption"
	$header +=  "DevicePasswordEnabled"
	$header +=  "PasswordRecoveryEnabled"
	$header +=  "DevicePolicyRefreshInterval"
	$header +=  "AllowSimpleDevicePassword"
	$header +=  "MaxAttachmentSize"
	$header +=  "WSSAccessEnabled"
	$header +=  "UNCAccessEnabled"
	$header +=  "MinDevicePasswordLength"
	$header +=  "MaxInactivityTimeDeviceLock"			# Column O
	$header +=  "MaxDevicePasswordFailedAttempts"
	$header +=  "DevicePasswordExpiration"
	$header +=  "DevicePasswordHistory"
	$header +=  "IsDefaultPolicy"
	$header +=  "AllowApplePushNotifications"
	$header +=  "AllowMicrosoftPushNotifications"
	$header +=  "AllowStorageCard"
	$header +=  "AllowCamera"
	$header +=  "RequireDeviceEncryption"
	$header +=  "AllowUnsignedApplications"
	$header +=  "AllowUnsignedInstallationPackages"
	$header +=  "AllowWiFi"
	$header +=  "AllowTextMessaging"
	$header +=  "AllowPOPIMAPEmail"
	$header +=  "AllowIrDA"
	$header +=  "RequireManualSyncWhenRoaming"
	$header +=  "AllowDesktopSync"
	$header +=  "AllowHTMLEmail"
	$header +=  "RequireSignedSMIMEMessages"
	$header +=  "RequireEncryptedSMIMEMessages"
	$header +=  "AllowSMIMESoftCerts"
	$header +=  "AllowBrowser"
	$header +=  "AllowConsumerEmail"
	$header +=  "AllowRemoteDesktop"
	$header +=  "AllowInternetSharing"
	$header +=  "AllowBluetooth"
	$header +=  "MaxCalendarAgeFilter"
	$header +=  "MaxEmailAgeFilter"
	$header +=  "RequireSignedSMIMEAlgorithm"
	$header +=  "RequireEncryptionSMIMEAlgorithm"
	$header +=  "AllowSMIMEEncryptionAlgorithmNegotiation"
	$header +=  "MinDevicePasswordComplexCharacters"
	$header +=  "MaxEmailBodyTruncationSize"
	$header +=  "MaxEmailHTMLBodyTruncationSize"
	$header +=  "UnapprovedInROMApplicationList"
	$header +=  "ApprovedApplicationList"
	$header +=  "AllowExternalDeviceManagement"
	$header +=  "MobileOTAUpdateMode"
	$header +=  "AllowMobileOTAUpdate"
	$header +=  "IrmEnabled"
	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\ExOrg_GetActiveSyncMbxPolicy.txt") -eq $true)
{	
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\ExOrg_GetActiveSyncMbxPolicy.txt") 
	# Send the data to the function to process and add to the Excel worksheet
	Process-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

# Format time/date columns
$EndRow = $DataFile.count + 1
# MaxinActivityTimeDeviceLock
$Column_Range = $Worksheet.Range("O1","O$EndRow")
$Column_Range.cells.NumberFormat = "hh:mm:ss"

	#EndRegion Get-ActiveSyncMailboxPolicy sheet
	
#Region Get-ActiveSyncVirtualDirectory sheet
Write-Host -Object "---- Starting Get-ActiveSyncVirtualDirectory"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "ActiveSyncVirtualDirectory"
	$Worksheet.Tab.ColorIndex = $intColorIndex_ClientAccess
	$row = 1
	$header = @()
	$header +=  "Server"
	$header +=  "Name"
	$header +=  "MetabasePath"
	$header +=  "Path"
	$header +=  "BasicAuthEnabled"
	$header +=  "WindowsAuthEnabled"
	$header +=  "CompressionEnabled"
	$header +=  "ClientCertAuth"
	$header +=  "WebsiteSSLEnabled"
	$header +=  "InternalUrl"
	$header +=  "InternalAuthenticationMethods"
	$header +=  "ExternalURL"
	$header +=  "ExternalAuthenticationMethods"
	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\GetActiveSyncVirtualDirectory") -eq $true)
{	
	foreach ($file in (Get-ChildItem -LiteralPath "$RunLocation\output\ExOrg\GetActiveSyncVirtualDirectory" | Where-Object {$_.name -match "~~GetActiveSyncVirtualDirectory"}))
	{
		$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\GetActiveSyncVirtualDirectory\" + $file) 
	}

	# Send the data to the function to process and add to the Excel worksheet
	Process-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}
	#EndRegion Get-ActiveSyncVirtualDirectory sheet

#Region Get-AutoDiscoverVirtualDirectory sheet
Write-Host -Object "---- Starting Get-AutoDiscoverVirtualDirectory"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "AutoDiscoverVirtualDirectory"
	$Worksheet.Tab.ColorIndex = $intColorIndex_ClientAccess
	$row = 1
	$header = @()
	$header +=  "Server"
	$header +=  "Name"
	$header +=  "MetabasePath"
	$header +=  "Path"
	$header +=  "BasicAuthentication"
	$header +=  "DigestAuthentication"
	$header +=  "WindowsAuthentication"
	$header +=  "InternalUrl"
	$header +=  "InternalAuthenticationMethods"
	$header +=  "ExternalURL"
	$header +=  "ExternalAuthenticationMethods"
	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\GetAutodiscoverVirtualDirectory") -eq $true)
{	
	foreach ($file in (Get-ChildItem -LiteralPath "$RunLocation\output\ExOrg\GetAutodiscoverVirtualDirectory" | Where-Object {$_.name -match "~~GetAutodiscoverVirtualDirectory"}))
	{
		$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\GetAutodiscoverVirtualDirectory\" + $file) 
	}

	# Send the data to the function to process and add to the Excel worksheet
	Process-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
} 
	#EndRegion Get-AutoDiscoverVirtualDirectory sheet

#Region Get-AvailabilityAddressSpace sheet
Write-Host -Object "---- Starting Get-AvailabilityAddressSpace"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "AvailabilityAddressSpace"
	$Worksheet.Tab.ColorIndex = $intColorIndex_ClientAccess
	$row = 1
	$header = @()
	$header +=  "Name"
	$header +=  "ForestName"
	$header +=  "UserName"
	$header +=  "UseServiceAccount"
	$header +=  "AccessMethod"
	$header +=  "ProxyUrl"
	$header +=  "TargetAutodiscoverEpr"

	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\ExOrg_GetAvailabilityAddressSpace.txt") -eq $true)
{	
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\ExOrg_GetAvailabilityAddressSpace.txt") 

	# Send the data to the function to process and add to the Excel worksheet
	Process-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}
	#EndRegion Get-AvailabilityAddressSpace sheet

#Region Get-ClientAccessArray sheet
Write-Host -Object "---- Starting Get-ClientAccessArray"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "ClientAccessArray"
	$Worksheet.Tab.ColorIndex = $intColorIndex_ClientAccess
	$row = 1
	$header = @()
	$header +=  "Name"
	$header +=  "Fqdn"
	$header +=  "Site"
	$header +=  "SiteName"
	$header +=  "Members"
	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\ExOrg_GetClientAccessArray.txt") -eq $true)
{	
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\ExOrg_GetClientAccessArray.txt")

	# Send the data to the function to process and add to the Excel worksheet
	Process-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}	
#EndRegion Get-ClientAccessArray sheet

#Region Get-ClientAccessServer sheet
Write-Host -Object "---- Starting Get-ClientAccessServer"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "ClientAccessServer"
	$Worksheet.Tab.ColorIndex = $intColorIndex_ClientAccess
	$row = 1
	$header = @()
	$header +=  "Name"
	$header +=  "Fqdn"
	$header +=  "OutlookAnywhereEnabled"
	$header +=  "AutoDiscoverServiceCN"
	$header +=  "AutoDiscoverServiceInternalUri"
	$header +=  "AutoDiscoverSiteScope"
	$header +=  "WhenCreated"
	$header +=  "IsValid"
	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\ExOrg_GetClientAccessSvr.txt") -eq $true)
{	
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\ExOrg_GetClientAccessSvr.txt")

	# Send the data to the function to process and add to the Excel worksheet
	Process-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}
	#EndRegion Get-ClientAccessServer sheet

#Region Get-ECPVirtualDirectory sheet
Write-Host -Object "---- Starting Get-ECPVirtualDirectory"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "ECPVirtualDirectory"
	$Worksheet.Tab.ColorIndex = $intColorIndex_ClientAccess
	$row = 1
	$header = @()
	$header +=  "Server"
	$header +=  "Name"
	$header +=  "MetabasePath"
	$header +=  "Path"
	$header +=  "BasicAuthentication"
	$header +=  "DigestAuthentication"
	$header +=  "FormsAuthentication"
	$header +=  "WindowsAuthentication"
	$header +=  "GzipLevel"
	$header +=  "InternalUrl"
	$header +=  "InternalAuthenticationMethods"
	$header +=  "ExternalURL"
	$header +=  "ExternalAuthenticationMethods"
	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\GetEcpVirtualDirectory") -eq $true)
{	
	foreach ($file in (Get-ChildItem -LiteralPath "$RunLocation\output\ExOrg\GetEcpVirtualDirectory" | Where-Object {$_.name -match "~~GetEcpVirtualDirectory"}))
	{
		$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\GetEcpVirtualDirectory\" + $file) 
	}

	# Send the data to the function to process and add to the Excel worksheet
	Process-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}
	#EndRegion Get-ECPVirtualDirectory sheet

#Region Get-OABVirtualDirectory sheet
Write-Host -Object "---- Starting Get-OABVirtualDirectory"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "OABVirtualDirectory"
	$Worksheet.Tab.ColorIndex = $intColorIndex_ClientAccess
	$row = 1
	$header = @()
	$header +=  "Server"
	$header +=  "Name"
	$header +=  "MetabasePath"
	$header +=  "Path"
	$header +=  "RequireSSL"
	$header +=  "BasicAuthentication"
	$header +=  "WindowsAuthentication"
	$header +=  "InternalURL"
	$header +=  "InternalAuthenticationMethods"
	$header +=  "ExternalURL"
	$header +=  "ExternalAuthenticationMethods"
	$header +=  "OfflineAddressBooks"
	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\GetOabVirtualDirectory") -eq $true)
{	
	foreach ($file in (Get-ChildItem -LiteralPath "$RunLocation\output\ExOrg\GetOabVirtualDirectory" | Where-Object {$_.name -match "~~GetOabVirtualDirectory"}))
	{
		$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\GetOabVirtualDirectory\" + $file) 
	}

	# Send the data to the function to process and add to the Excel worksheet
	Process-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}
#EndRegion Get-OABVirtualDirectory sheet

#Region Get-OutlookAnywhere sheet
Write-Host -Object "---- Starting Get-OutlookAnywhere"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "OutlookAnywhere"
	$Worksheet.Tab.ColorIndex = $intColorIndex_ClientAccess
	$row = 1
	$header = @()
	$header +=  "ServerName"
	$header +=  "SSLOffloading"
	$header +=  "ExternalHostname"
	$header +=  "InternalHostname"
	$header +=  "ClientAuthenticationMethod (Ex2010)"
	$header +=  "InternalClientAuthenticationMethod"
	$header +=  "ExternalClientAuthenticationMethod"
	$header +=  "IISAuthenticationMethods"
	$header +=  "MetabasePath"
	$header +=  "ExchangeVersion"
	$header +=  "Name"
	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\ExOrg_GetOutlookAnywhere.txt") -eq $true)
{	
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\ExOrg_GetOutlookAnywhere.txt") 

	# Send the data to the function to process and add to the Excel worksheet
	Process-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}
	#EndRegion Get-OutlookAnywhere sheet

#Region Get-OwaMailboxPolicy sheet
Write-Host -Object "---- Starting Get-OwaMailboxPolicy"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "OwaMailboxPolicy"
	$Worksheet.Tab.ColorIndex = $intColorIndex_ClientAccess
	$row = 1
	$header = @()
	$header +=  "Name"
	$header +=  "DirectFileAccessOnPublicComputersEnabled"
	$header +=  "DirectFileAccessOnPrivateComputersEnabled"
	$header +=  "WebReadyDocumentViewingOnPublicComputersEnabled"
	$header +=  "WebReadyDocumentViewingOnPrivateComputersEnabled"
	$header +=  "ForceWebReadyDocumentViewingFirstOnPublicComputers"
	$header +=  "ForceWebReadyDocumentViewingFirstOnPrivateComputers"
	$header +=  "ActionForUnknownFileAndMIMETypes"
	$header +=  "WebReadyDocumentViewingForAllSupportedTypes"
	$header +=  "PhoneticSupportEnabled"
	$header +=  "DefaultTheme"
	$header +=  "DefaultClientLanguage"
	$header +=  "LogonAndErrorLanguage"
	$header +=  "UseGB18030"
	$header +=  "UseISO885915"
	$header +=  "OutboundCharset"
	$header +=  "GlobalAddressListEnabled"
	$header +=  "OrganizationEnabled"
	$header +=  "ExplicitLogonEnabled"
	$header +=  "OWALightEnabled"
	$header +=  "OWAMiniEnabled (Ex2010)"
	$header +=  "DelegateAccessEnabled"
	$header +=  "IRMEnabled"
	$header +=  "CalendarEnabled"
	$header +=  "ContactsEnabled"
	$header +=  "TasksEnabled"
	$header +=  "JournalEnabled"
	$header +=  "NotesEnabled"
	$header +=  "RemindersAndNotificationsEnabled"
	$header +=  "PremiumClientEnabled"
	$header +=  "SpellCheckerEnabled"
	$header +=  "SearchFoldersEnabled"
	$header +=  "SignaturesEnabled"
	$header +=  "ThemeSelectionEnabled"
	$header +=  "JunkEmailEnabled"
	$header +=  "UMIntegrationEnabled"
	$header +=  "WSSAccessOnPublicComputersEnabled"
	$header +=  "WSSAccessOnPrivateComputersEnabled"
	$header +=  "ChangePasswordEnabled"
	$header +=  "UNCAccessOnPublicComputersEnabled"
	$header +=  "UNCAccessOnPrivateComputersEnabled"
	$header +=  "ActiveSyncIntegrationEnabled"
	$header +=  "AllAddressListsEnabled"
	$header +=  "RulesEnabled"
	$header +=  "PublicFoldersEnabled"
	$header +=  "SMimeEnabled"
	$header +=  "RecoverDeletedItemsEnabled"
	$header +=  "InstantMessagingEnabled"
	$header +=  "TextMessagingEnabled"
	$header +=  "ForceSaveAttachmentFilteringEnabled"
	$header +=  "SilverlightEnabled"
	$header +=  "InstantMessagingType"
	#$a = [int][char]'a' -1
	#if ($header.GetLength(0) -gt 26) 
	${$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	$else 
	${$EndCellColumn = [char]($header.GetLength(0) + $a)}

    $HeaderCount = $header.getlength(0)
      if ($HeaderCount -gt 26) 
    {
        $m = ((($HeaderCount-1)%26)+65)
        $d = [int][math]::floor((($HeaderCount-1)/26)+64)
        $EndCellColumn = [char]$d + [char]$m
    }
    else
    {
        $EndCellColumn = [char]($HeaderCount+64)
    }

	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1
if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\ExOrg_GetOwaMailboxPolicy.txt") -eq $true)
{	
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\ExOrg_GetOwaMailboxPolicy.txt") 

	# Send the data to the function to process and add to the Excel worksheet
	Process-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}
	#EndRegion Get-ActiveSyncMailboxPolicy sheet

#Region Get-OWAVirtualDirectory sheet
Write-Host -Object "---- Starting Get-OWAVirtualDirectory"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "OWAVirtualDirectory"
	$Worksheet.Tab.ColorIndex = $intColorIndex_ClientAccess
	$row = 1
	$header = @()
	$header +=  "Server"
	$header +=  "Name"
	$header +=  "MetabasePath"
	$header +=  "Path"
	$header +=  "BasicAuthentication"
	$header +=  "DigestAuthentication"
	$header +=  "WindowsAuthentication"
	$header +=  "FormsAuthentication"
	$header +=  "LiveIdAuthentication"
	$header +=  "InternalUrl"
	$header +=  "InternalAuthenticationMethods"
	$header +=  "ExternalURL"
	$header +=  "ExternalAuthenticationMethods"
	$header +=  "GzipLevel"
	$header +=  "OwaVersion"
	$header +=  "ChangePasswordEnabled"
	$header +=  "FailBackUrl"
	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\GetOwaVirtualDirectory") -eq $true)
{	
	foreach ($file in (Get-ChildItem -LiteralPath "$RunLocation\output\ExOrg\GetOwaVirtualDirectory" | Where-Object {$_.name -match "~~GetOwaVirtualDirectory"}))
	{
		$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\GetOwaVirtualDirectory\" + $file) 
	}

	# Send the data to the function to process and add to the Excel worksheet
	Process-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}
	#EndRegion Get-OWAVirtualDirectory sheet

#Region Get-PowershellVirtualDirectory sheet
Write-Host -Object "---- Starting Get-PowershellVirtualDirectory"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "PowershellVirtualDirectory"
	$Worksheet.Tab.ColorIndex = $intColorIndex_ClientAccess
	$row = 1
	$header = @()
	$header +=  "Server"
	$header +=  "Name"
	$header +=  "RequireSSL"
	$header +=  "MetabasePath"
	$header +=  "Path"
	$header +=  "BasicAuthentication"
	$header +=  "DigestAuthentication"
	$header +=  "WindowsAuthentication"
	$header +=  "InternalUrl"
	$header +=  "InternalAuthenticationMethods"
	$header +=  "ExternalUrl"
	$header +=  "ExternalAuthenticationMethods"
	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\GetPowershellVirtualDirectory") -eq $true)
{	
	foreach ($file in (Get-ChildItem -LiteralPath "$RunLocation\output\ExOrg\GetPowershellVirtualDirectory" | Where-Object {$_.name -match "~~GetPowershellVirtualDirectory"}))
	{
		$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\GetPowershellVirtualDirectory\" + $file) 
	}

	# Send the data to the function to process and add to the Excel worksheet
	Process-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}
	#EndRegion Get-PowershellVirtualDirectory sheet

#Region Get-RPCClientAccess sheet
Write-Host -Object "---- Starting Get-RPCClientAccess"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "RPCClientAccess"
	$Worksheet.Tab.ColorIndex = $intColorIndex_ClientAccess
	$row = 1
	$header = @()
	$header +=  "Server"
	$header +=  "MaximumConnections"
	$header +=  "EncryptionRequired"
	$header +=  "BlockedClientVersions"
	$header +=  "Responsibility"
	$header +=  "Name"
	$header +=  "IsValid"
	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\ExOrg_GetRPCClientAccess.txt") -eq $true)
{	
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\ExOrg_GetRPCClientAccess.txt")

	# Send the data to the function to process and add to the Excel worksheet
	Process-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}
	#EndRegion Get-RPCClientAccess sheet

#Region Get-ThrottlingPolicy sheet
Write-Host -Object "---- Starting Get-ThrottlingPolicy"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "ThrottlingPolicy"
	$Worksheet.Tab.ColorIndex = $intColorIndex_ClientAccess
	$row = 1
	$header = @()
	$header +=  "Identity"
	$header +=  "Name"
	$header +=  "IsDefault"
	$header +=  "EWSMaxConcurrency"
	$header +=  "EWSPercentTimeInAD (Ex2010)"
	$header +=  "EWSPercentTimeInCAS (Ex2010)"
	$header +=  "EWSPercentTimeInMailboxRPC (Ex2010)"
	$header +=  "EWSMaxSubscriptions"
	$header +=  "EWSFastSearchTimeoutInSeconds (Ex2010)"
	$header +=  "EWSFindCountLimit (Ex2010)"
	$header +=  "RCAMaxConcurrency"
	$header +=  "RCAPercentTimeInAD (Ex2010)"
	$header +=  "RCAPercentTimeInCAS (Ex2010)"
	$header +=  "RCAPercentTimeInMailboxRPC (Ex2010)"
	$header +=  "CPAMaxConcurrency"
	$header +=  "CPAPercentTimeInCAS (Ex2010)"
	$header +=  "CPAPercentTimeInMailboxRPC (Ex2010)"
	$header +=  "AnonymousMaxConcurrency"
	$header +=  "AnonymousPercentTimeInAD (Ex2010)"
	$header +=  "AnonymousPercentTimeInCAS (Ex2010)"
	$header +=  "AnonymousPercentTimeInMailboxRPC (Ex2010)"
	$header +=  "EASMaxConcurrency"
	$header +=  "EASPercentTimeInAD (Ex2010)"
	$header +=  "EASPercentTimeInCAS (Ex2010)"
	$header +=  "EASPercentTimeInMailboxRPC (Ex2010)"
	$header +=  "EASMaxDevices"
	$header +=  "EASMaxDeviceDeletesPerMonth"
	$header +=  "IMAPMaxConcurrency"
	$header +=  "IMAPPercentTimeInAD (Ex2010)"
	$header +=  "IMAPPercentTimeInCAS (Ex2010)"
	$header +=  "IMAPPercentTimeInMailboxRPC (Ex2010)"
	$header +=  "OWAMaxConcurrency"
	$header +=  "OWAPercentTimeInAD (Ex2010)"
	$header +=  "OWAPercentTimeInCAS (Ex2010)"
	$header +=  "OWAPercentTimeInMailboxRPC (Ex2010)"
	$header +=  "POPMaxConcurrency"
	$header +=  "POPPercentTimeInAD (Ex2010)"
	$header +=  "POPPercentTimeInCAS (Ex2010)"
	$header +=  "POPPercentTimeInMailboxRPC (Ex2010)"
	$header +=  "PowerShellMaxConcurrency"
	$header +=  "PowerShellMaxTenantConcurrency"
	$header +=  "PowerShellMaxCmdlets"
	$header +=  "PowerShellMaxCmdletsTimePeriod"
	$header +=  "ExchangeMaxCmdlets"
	$header +=  "PowerShellMaxCmdletQueueDepth"
	$header +=  "PowerShellMaxDestructiveCmdlets"
	$header +=  "PowerShellMaxDestructiveCmdletsTimePeriod"
	$header +=  "MessageRateLimit"
	$header +=  "RecipientRateLimit"
	$header +=  "ForwardeeLimit"
	$header +=  "CPUStartPercent (Ex2010)"
	$header +=  "DiscoveryMaxConcurrency"
	$header +=  "DiscoveryMaxMailboxes"
	$header +=  "WhenCreatedUTC"
	$header +=  "WhenChangedUTC "
	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\ExOrg_GetThrottlingPolicy.txt") -eq $true)
{	
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\ExOrg_GetThrottlingPolicy.txt")

	# Send the data to the function to process and add to the Excel worksheet
	Process-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

# Format time/date columns
$EndRow = $DataFile.count + 1
# WhenCreatedUTC
$Column_Range = $Worksheet.Range("BB1","BB$EndRow")
$Column_Range.cells.NumberFormat = "mm/dd/yy hh:mm:ss"
# WhenChangedUTC
$Column_Range = $Worksheet.Range("BC1","BC$EndRow")
$Column_Range.cells.NumberFormat = "mm/dd/yy hh:mm:ss"
	#EndRegion Get-ThrottlingPolicy sheet

#Region Get-WebServicesVirtualDirectory sheet
Write-Host -Object "---- Starting Get-WebServicesVirtualDirectory"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "WebServicesVirtualDirectory"
	$Worksheet.Tab.ColorIndex = $intColorIndex_ClientAccess
	$row = 1
	$header = @()
	$header +=  "Server"
	$header +=  "Name"
	$header +=  "MetabasePath"
	$header +=  "Path"
	$header +=  "BasicAuthentication"
	$header +=  "DigestAuthentication"
	$header +=  "WindowsAuthentication"
	$header +=  "InternalUrl"
	$header +=  "InternalAuthenticationMethods"
	$header +=  "ExternalURL"
	$header +=  "ExternalAuthenticationMethods"
	$header +=  "InternalNLBBypassUrl"
	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\GetWebServicesVirtualDirectory") -eq $true)
{	
	foreach ($file in (Get-ChildItem -LiteralPath "$RunLocation\output\ExOrg\GetWebServicesVirtualDirectory" | Where-Object {$_.name -match "~~GetWebServicesVirtualDirectory"}))
	{
		$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\GetWebServicesVirtualDirectory\" + $file) 
	}

	# Send the data to the function to process and add to the Excel worksheet
	Process-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}
	#EndRegion Get-WebServicesVirtualDirectory sheet


# Global
#Region Get-AddressBookPolicy sheet
Write-Host -Object "---- Starting Get-AddressBookPolicy"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "AddressBookPolicy"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Global
	$row = 1
	$header = @()
	$header +=  "Name"
	$header +=  "AddressLists"
	$header +=  "GlobalAddressList"
	$header +=  "RoomList"
	$header +=  "OfflineAddressBook"
	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\ExOrg_GetAddressBookPolicy.txt") -eq $true)
{	
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\ExOrg_GetAddressBookPolicy.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Process-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}
	#EndRegion Get-AddressBookPolicy sheet

#Region Get-AddressList sheet
Write-Host -Object "---- Starting Get-AddressList"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "AddressList"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Global
	$row = 1
	$header = @()
	$header +=  "DisplayName"
	$header +=  "Path"
	$header +=  "RecipientFilter"
	$header +=  "WhenCreatedUTC"
	$header +=  "WhenChangedUTC"
	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\ExOrg_GetAddressList.txt") -eq $true)
{	
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\ExOrg_GetAddressList.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Process-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

# Format time/date columns
$EndRow = $DataFile.count + 1
# WhenCreatedUTC
$Column_Range = $Worksheet.Range("D1","D$EndRow")
$Column_Range.cells.NumberFormat = "mm/dd/yy hh:mm:ss"
# WhenChangedUTC
$Column_Range = $Worksheet.Range("E1","E$EndRow")
$Column_Range.cells.NumberFormat = "mm/dd/yy hh:mm:ss"

	#EndRegion Get-AddressList sheet

#Region Get-DatabaseAvailabilityGroup sheet
Write-Host -Object "---- Starting Get-DatabaseAvailabilityGroup"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "DatabaseAvailabilityGroup"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Global
	$row = 1
	$header = @()
	$header +=  "Name"
	$header +=  "Identity"
	$header +=  "Servers"
	$header +=  "WitnessServer"
	$header +=  "WitnessDirectory"
	$header +=  "AlternateWitnessServer"
	$header +=  "AlternateWitnessDirectory"
	$header +=  "NetworkCompression"
	$header +=  "NetworkEncryption"
	$header +=  "DatacenterActivationMode"
	$header +=  "AutoDagDatabaseCopiesPerVolume"
	$header +=  "AutoDagDatabaseCopiesPerDatabase"
	$header +=  "AutoDagDatabasesRootFolderPath"
	$header +=  "AutoDagVolumesRootFolderPath"
	$header +=  "StoppedMailboxServers"
	$header +=  "StartedMailboxServer"
	$header +=  "DatabaseAvailabilityGroupIpv4Addresses"
	$header +=  "DatabaseAvailabilityGroupIpAddresses"
	$header +=  "AllowCrossSiteRpcClientAccess"
	$header +=  "OperationalServers"
	$header +=  "PrimaryActiveManager"
	$header +=  "ServersInMaintenance"
	$header +=  "ThirdPartyReplication"
	$header +=  "NetworkNames"
	$header +=  "ReplicationPort"
	$header +=  "WitnessShareInUse"
	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\ExOrg_GetDAG.txt") -eq $true)
{	
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\ExOrg_GetDAG.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Process-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}
	#EndRegion Get-DatabaseAvailabilityGroup sheet

#Region Get-DatabaseAvailabilityGroupNetwork sheet
Write-Host -Object "---- Starting Get-DatabaseAvailabilityGroupNetwork"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "DAGNetwork"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Global
	$row = 1
	$header = @()
	$header +=  "Name"
	$header +=  "Subnets"
	$header +=  "Interfaces"
	$header +=  "MapiAccessEnabled"
	$header +=  "ReplicationEnabled"
	$header +=  "IgnoreNetwork"
	$header +=  "Identity"
	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\ExOrg_GetDatabaseAvailabilityGroupNetwork.txt") -eq $true)
{	
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\ExOrg_GetDatabaseAvailabilityGroupNetwork.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Process-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}
	#EndRegion Get-DatabaseAvailabilityGroupNetwork sheet

#Region Get-EmailAddressPolicy sheet
Write-Host -Object "---- Starting Get-EmailAddressPolicy"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "EmailAddressPolicy"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Global
	$row = 1
	$header = @()
	$header +=  "Name"
	$header +=  "IsValid"
	$header +=  "RecipientFilter"
	$header +=  "LdapRecipientFilter"
	$header +=  "LastUpdatedRecipientFilter"
	$header +=  "RecipientFilterApplied"
	$header +=  "IncludedRecipients"
	$header +=  "ConditionalDepartment"
	$header +=  "ConditionalCompany"
	$header +=  "ConditionalStateOrProvince"
	$header +=  "ConditionalCustomAttribute1"
	$header +=  "ConditionalCustomAttribute2"
	$header +=  "ConditionalCustomAttribute3"
	$header +=  "ConditionalCustomAttribute4"
	$header +=  "ConditionalCustomAttribute5"
	$header +=  "ConditionalCustomAttribute6"
	$header +=  "ConditionalCustomAttribute7"
	$header +=  "ConditionalCustomAttribute8"
	$header +=  "ConditionalCustomAttribute9"
	$header +=  "ConditionalCustomAttribute10"
	$header +=  "ConditionalCustomAttribute11"
	$header +=  "ConditionalCustomAttribute12"
	$header +=  "ConditionalCustomAttribute13"
	$header +=  "ConditionalCustomAttribute14"
	$header +=  "ConditionalCustomAttribute15"
	$header +=  "RecipientContainer"
	$header +=  "RecipientFilterType"
	$header +=  "Priority"
	$header +=  "EnabledPrimarySMTPAddressTemplate"
	$header +=  "EnabledEmailAddressTemplates"
	$header +=  "DisabledEmailAddressTemplates"
	$header +=  "HasEmailAddressSetting"
	$header +=  "HasMailboxManagerSetting"
	$header +=  "NonAuthoritativeDomains"
	$header +=  "ExchangeVersion"
	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\ExOrg_GetEmailAddressPolicy.txt") -eq $true)
{	
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\ExOrg_GetEmailAddressPolicy.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Process-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}
	#EndRegion Get-EmailAddressPolicy sheet

#Region Get-ExchangeCertificate sheet
Write-Host -Object "---- Starting Get-ExchangeCertificate"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "ExchangeCertificate"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Global 
	$row = 1
	$header = @()
	$header +=  "Server"
	$header +=  "Status"
	$header +=  "IsSelfSigned"
	$header +=  "RootCAType"
	$header +=  "PublicKeySize"
	$header +=  "Thumbprint"
	$header +=  "Services"
	$header +=  "Subject"
	$header +=  "Issuer"
	$header +=  "NotBefore"		# Column J
	$header +=  "NotAfter"		# Column K
	$header +=  "CertificateDomains"
	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\GetExchCert") -eq $true)
{
	foreach ($file in (Get-ChildItem -LiteralPath "$RunLocation\output\ExOrg\GetExchCert" | Where-Object {$_.name -match "~~GetExchCert"}))
	{
		$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\GetExchCert\" + $file)
	}
		$RowCount = $DataFile.Count
	# Send the data to the function to process and add to the Excel worksheet
	Process-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

# Format time/date columns
$EndRow = $DataFile.count + 1
# NotBefore
$Column_Range = $Worksheet.Range("J1","J$EndRow")
$Column_Range.cells.NumberFormat = "mm/dd/yy hh:mm:ss"
# NotAfter
$Column_Range = $Worksheet.Range("K1","K$EndRow")
$Column_Range.cells.NumberFormat = "mm/dd/yy hh:mm:ss"

	#EndRegion Get-ExchangeCertificate sheet

#Region Get-ExchangeServer sheet
Write-Host -Object "---- Starting Get-ExchangeServer"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "ExchangeServer"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Global
	$row = 1
	$header = @()
	$header +=  "Name"
	$header +=  "Fqdn"
	$header +=  "Edition"
	$header +=  "ErrorReportingEnabled"
	$header +=  "CustomerFeedbackEnabled"
	$header +=  "ServerRole"
	$header +=  "AdminDisplayVersion"
	$header +=  "StaticDomainControllers"
	$header +=  "StaticGlobalCatalogs"
	$header +=  "StaticConfigDomainController"
	$header +=  "StaticExcludedDomainControllers"
	$header +=  "Site"
	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\ExOrg_GetExchangeSvr.txt") -eq $true)
{	
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\ExOrg_GetExchangeSvr.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Process-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}
	#EndRegion Get-ExchangeServer sheet

#Region Get-MailboxDatabase sheet
Write-Host -Object "---- Starting Get-MailboxDatabase"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "MailboxDatabase"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Global
	$row = 1
	$header = @()
	$header +=  "Server"
	$header +=  "Name"
	$header +=  "Recovery"
	$header +=  "EdbFilePath"
	$header +=  "LogFolderPath"
	$header +=  "MountAtStartup"
	$header +=  "OfflineAddressBook"
	$header +=  "PublicFolderDatabase"
	$header +=  "IssueWarningQuota"
	$header +=  "ProhibitSendQuota"
	$header +=  "ProhibitSendReceiveQuota"
	$header +=  "DatabaseCopies"
	$header +=  "ReplayLagTimes"
	$header +=  "TruncationLagTimes"
	$header +=  "RPCClientAccessServer"
	$header +=  "MasterServerOrAvailabilityGroup"
	$header +=  "MasterType"
	$header +=  "DataMoveReplicationConstraint"
	$header +=  "ActivationPreference"
	$header +=  "MailboxRetention"
	$header +=  "DeletedItemRetention"
	$header +=  "BackgroundDatabaseMaintenance"
	$header +=  "Servers"
	$header +=  "MountedOnServer"
	$header +=  "CircularLoggingEnabled"
	$header +=  "Database Size"
	$header +=  "Database Size (MB)"
	$header +=  "Database Size (GB)"
	$header +=  "AvailableNewMailboxSpace (MB)"
	$header +=  "AvailableNewMailboxSpace (%)"
	$header +=  "Mailbox Count"
	$header +=  "AllowFileRestore"
	$header +=  "RetainDeletedItemsUntilBackup"
	$header +=  "SnapshotLastFullBackup"			# Column AH
	$header +=  "SnapshotLastIncrementalBackup"		# Column AI
	$header +=  "SnapshotLastDifferentialBackup"	# Column AJ
	$header +=  "LastFullBackup"					# Column AK
	$header +=  "LastIncrementalBackup"				# Column AL
	$header +=  "LastDifferentialBackup"			# Column AM
	$header +=  "VolumeName"
	$header +=  "VolumeCapacity (MB)"
	$header +=  "VolumeFreespace (MB)"
	$header +=  "VolumeFreespace (%)"
	$header +=  "JournalRecipient"
	$header +=  "MaintenanceSchedule"
	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\GetMbxDb") -eq $true)
{
	foreach ($file in (Get-ChildItem -LiteralPath "$RunLocation\output\ExOrg\GetMbxDb" | Where-Object {$_.name -match "~~GetMbxDb"}))
	{
		$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\GetMbxDb\" + $file)
	}
	# Send the data to the function to process and add to the Excel worksheet
	Process-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

# Format time/date columns
$EndRow = $DataFile.count + 1
# SnapshotLastFullBackup
$Column_Range = $Worksheet.Range("AH1","AH$EndRow")
$Column_Range.cells.NumberFormat = "mm/dd/yy hh:mm:ss"
# SnapshotLastIncrementalBackup
$Column_Range = $Worksheet.Range("AI1","AI$EndRow")
$Column_Range.cells.NumberFormat = "mm/dd/yy hh:mm:ss"
# SnapshotLastDifferentialBackup
$Column_Range = $Worksheet.Range("AJ1","AJ$EndRow")
$Column_Range.cells.NumberFormat = "mm/dd/yy hh:mm:ss"
# LastFullBackup
$Column_Range = $Worksheet.Range("AK1","AK$EndRow")
$Column_Range.cells.NumberFormat = "mm/dd/yy hh:mm:ss"
# LastIncrementalBackup
$Column_Range = $Worksheet.Range("AL1","AL$EndRow")
$Column_Range.cells.NumberFormat = "mm/dd/yy hh:mm:ss"
# LastDifferentialBackup
$Column_Range = $Worksheet.Range("AM1","AM$EndRow")
$Column_Range.cells.NumberFormat = "mm/dd/yy hh:mm:ss"

	#EndRegion Get-MailboxDatabase sheet

#Region Get-MailboxDatabaseCopyStatus sheet
Write-Host -Object "---- Starting Get-MailboxDatabaseCopyStatus"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "MailboxDatabaseCopyStatus"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Global
	$row = 1
	$header = @()
	$header +=  "MailboxServer"
	$header +=  "DatabaseName"
	$header +=  "Identity"
	$header +=  "Status"
	$header +=  "ActiveDatabaseCopy"
	$header +=  "ActivationSuspended"
	$header +=  "CopyQueueLength"
	$header +=  "ReplayQueueLength"
	$header +=  "LogCopyQueueIncreasing"
	$header +=  "LogReplayQueueIncreasing"
	$header +=  "ContentIndexState"
	$header +=  "LastInspectedLogTime"		# Column L
	$header +=  "ActiveCopy"
	$header +=  "IncomingLogCopyingNetwork"
	$header +=  "SeedingNetwork"
	$header +=  "ActivationPreference"
	$header +=  "StatusRetrievalTime"
	$header +=  "AutoActivationPolicy"
	$header +=  "DatabaseVolumeMountPoint"
	$header +=  "LogVolumeMountPoint"

	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\GetMbxDatabaseCopyStatus") -eq $true)
{	
	foreach ($file in (Get-ChildItem -LiteralPath "$RunLocation\output\ExOrg\GetMbxDatabaseCopyStatus" | Where-Object {$_.name -match "~~GetMbxDatabaseCopyStatus"}))
	{
		$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\GetMbxDatabaseCopyStatus\" + $file) 
	}
	# Send the data to the function to process and add to the Excel worksheet
	Process-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

# Format time/date columns
$EndRow = $DataFile.count + 1
# LastInspectedLogTime
$Column_Range = $Worksheet.Range("L1","L$EndRow")
$Column_Range.cells.NumberFormat = "mm/dd/yy hh:mm:ss"

	#EndRegion Get-MailboxDatabaseCopyStatus sheet

#Region Get-MailboxServer sheet
Write-Host -Object "---- Starting Get-MailboxServer"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "MailboxServer"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Global
	$row = 1
	$header = @()
	$header +=  "Name"
	$header +=  "MessageTrackingLogEnabled (Ex2010)"
	$header +=  "MessageTrackingLogSubjectLoggingEnabled (Ex2010)"
	$header +=  "AutodatabaseMountDial"
	$header +=  "DatabaseCopyAutoActivationPolicy"
	$header +=  "MAPIEncryptionRequired"
	$header +=  "DatabaseAvailabilityGroup"
	$header +=  "ServerRole"
	$header +=  "DataPath"
	$header +=  "CalendarRepairMode"
	$header +=  "DatabaseCopyActivationDisabledAndMoveNow"
	$header +=  "MaximumActiveDatabases"
	$header +=  "MaximumPreferredActiveDatabases"

	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\ExOrg_GetMbxSvr.txt") -eq $true)
{	
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\ExOrg_GetMbxSvr.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Process-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}
	#EndRegion Get-MailboxServer sheet

#Region Get-OfflineAddressBook sheet
Write-Host -Object "---- Starting Get-OfflineAddressBook"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "OfflineAddressBook"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Global
	$row = 1
	$header = @()
	$header +=  "Name"
	$header +=  "Identity"
	$header +=  "Server"
	$header +=  "AddressLists"
	$header +=  "Versions"
	$header +=  "IsDefault"
	$header +=  "PublicFolderDatabase"
	$header +=  "PublicFolderDistributionEnabled"
	$header +=  "WebDistributionEnabled"
	$header +=  "VirtualDirectories"
	$header +=  "Schedule"
	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\ExOrg_GetOfflineAddressBook.txt") -eq $true)
{	
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\ExOrg_GetOfflineAddressBook.txt") 
	# Send the data to the function to process and add to the Excel worksheet
	Process-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}
	#EndRegion Get-OfflineAddressBook sheet

#Region Get-OrgConfig sheet
Write-Host -Object "---- Starting Get-OrgConfig"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "OrganizationConfig"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Global
	$row = 1
	$header = @()
	$header +=  "Name"
	$header +=  "DefaultPublicFolderDatabase (Ex2010 only)"
	$header +=  "IssueWarningQuota (Ex2010 only)"
	$header +=  "PublicFolderContentReplicationDisabled (Ex2010 only)"
	$header +=  "PublicFoldersLockedForMigration"
	$header +=  "PublicFolderMigrationComplete"
	$header +=  "PublicFoldersEnabled"
	$header +=  "DefaultPublicFolderAgeLimit"
	$header +=  "DefaultPublicFolderIssueWarningQuota"
	$header +=  "DefaultPublicFolderProhibitPostQuota"
	$header +=  "DefaultPublicFolderMaxItemSize"
	$header +=  "DefaultPublicFolderDeletedItemRetention"
	$header +=  "DefaultPublicFolderMovedItemRetention"
	$header +=  "PublicFolderMailboxesLockedForNewConnections"
	$header +=  "PublicFolderMailboxesMigrationComplete"
	$header +=  "IsMixedMode"
	$header +=  "SCLJunkThreshold"
	$header +=  "Industry"
	$header +=  "CustomerFeedbackEnabled"
	$header +=  "OrganizationSummary"
	$header +=  "MailTipsExternalRecipientsTipsEnabled"
	$header +=  "MailTipsLargeAudienceThreshold"
	$header +=  "MailTipsMailboxSourcedTipsEnabled"
	$header +=  "MailTipsGroupMetricsEnabled"
	$header +=  "MailTipsAllTipsEnabled"
	$header +=  "ReadTrackingEnabled"
	$header +=  "DistributionGroupDefaultOU"
	$header +=  "DistributionGroupNameBlockedWordsList"
	$header +=  "DistributionGroupNamingPolicy"
	$header +=  "EwsEnabled"
	$header +=  "EwsAllowOutlook"
	$header +=  "EwsAllowMacOutlook"
	$header +=  "EwsAllowEntourage"
	$header +=  "EwsApplicationAccessPolicy"
	$header +=  "EwsAllowList"
	$header +=  "EwsBlockList"
	$header +=  "ActivityBasedAuthenticationTimeoutInterval"
	$header +=  "ActivityBasedAuthenticationTimeoutEnabled"
	$header +=  "ActivityBasedAuthenticationTimeoutWithSingleSignOnEnabled"
	$header +=  "IPListBlocked"
	$header +=  "AutoExpandingArchiveEnabled"
	$header +=  "MaxConcurrentMigrations"
	$header +=  "IntuneManagedStatus"
	$header +=  "AzurePremiumSubscriptionStatus"
	$header +=  "HybridConfigurationStatus"
	$header +=  "UnblockUnsafeSenderPromptEnabled"
	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\ExOrg_GetOrgConfig.txt") -eq $true)
{	
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\ExOrg_GetOrgConfig.txt") 
	# Send the data to the function to process and add to the Excel worksheet
	Process-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}
# Format time/date columns
$EndRow = $DataFile.count + 1
# DefaultPublicFolderDeletedItemRetention
$Column_Range = $Worksheet.Range("L1","L$EndRow")
$Column_Range.cells.NumberFormat = "dd:hh:mm:ss"
# DefaultPublicFolderMovedItemRetention
$Column_Range = $Worksheet.Range("M1","M$EndRow")
$Column_Range.cells.NumberFormat = "dd:hh:mm:ss"
# ActivityBasedAuthenticationTimeoutInterval
$Column_Range = $Worksheet.Range("AK1","AK$EndRow")
$Column_Range.cells.NumberFormat = "hh:mm:ss"
	#EndRegion Get-OrgConfig sheet

#Region Get-PublicFolderDatabase sheet
Write-Host -Object "---- Starting Get-PublicFolderDatabase"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "PublicFolderDatabase"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Global
	$row = 1
	$header = @()
	$header +=  "Server"
	$header +=  "Alias"
	$header +=  "EdbFilePath"
	$header +=  "LogFolderPath"
	$header +=  "MountAtStartup"
	$header +=  "FirstInstance"
	$header +=  "MaxItemSize"
	$header +=  "ItemRetentionPeriod"
	$header +=  "ProhibitPostQuota"
	$header +=  "ReplicationSchedule"
	$header +=  "IssueWarningQuota"
	$header +=  "Name"
	$header +=  "Database Size"
	$header +=  "Database Size (MB)"
	$header +=  "Database Size (GB)"
	$header +=  "AvailableNewMailboxSpace (MB)"
	$header +=  "AllowFileRestore"
	$header +=  "RetainDeletedItemsUntilBackup"
	$header +=  "SnapshotLastFullBackup"			# Column S
	$header +=  "SnapshotLastIncrementalBackup"		# Column T
	$header +=  "SnapshotLastDifferentialBackup"	# Column U
	$header +=  "LastFullBackup"					# Column V
	$header +=  "LastIncrementalBackup"				# Column W
	$header +=  "LastDifferentialBackup"			# Column X
	$header +=  "MaintenanceSchedule"
	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\GetPublicFolderDb") -eq $true)
{
	foreach ($file in (Get-ChildItem -LiteralPath "$RunLocation\output\ExOrg\GetPublicFolderDb" | Where-Object {$_.name -match "~~GetPublicFolderDb"}))
	{
		$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\GetPublicFolderDb\" + $file)
	}
	# Send the data to the function to process and add to the Excel worksheet
	Process-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}
# Format time/date columns
$EndRow = $DataFile.count + 1
# SnapshotLastFullBackup
$Column_Range = $Worksheet.Range("s1","s$EndRow")
$Column_Range.cells.NumberFormat = "mm/dd/yy hh:mm:ss"
# SnapshotLastIncrementalBackup
$Column_Range = $Worksheet.Range("t1","t$EndRow")
$Column_Range.cells.NumberFormat = "mm/dd/yy hh:mm:ss"
# SnapshotLastDifferentialBackup
$Column_Range = $Worksheet.Range("u1","u$EndRow")
$Column_Range.cells.NumberFormat = "mm/dd/yy hh:mm:ss"
# LastFullBackup
$Column_Range = $Worksheet.Range("v1","v$EndRow")
$Column_Range.cells.NumberFormat = "mm/dd/yy hh:mm:ss"
# LastIncrementalBackup
$Column_Range = $Worksheet.Range("w1","w$EndRow")
$Column_Range.cells.NumberFormat = "mm/dd/yy hh:mm:ss"
# LastDifferentialBackup
$Column_Range = $Worksheet.Range("x1","x$EndRow")
$Column_Range.cells.NumberFormat = "mm/dd/yy hh:mm:ss"

	#EndRegion Get-PublicFolderDatabase sheet

#Region Get-Rbac sheet
Write-Host -Object "---- Starting Get-Rbac"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "Rbac"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Transport
	$row = 1
	$header = @()
	$header +=  "Name"
	$header +=  "Members"
	$header +=  "Roles"
	
	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\ExOrg_GetRbac.xml") -eq $true)
{	
	$DataFile = Import-Clixml "$RunLocation\output\ExOrg\ExOrg_GetRbac.xml"
	$RowCount = $DataFile.Count
	$ArrayRow = 0
	$BadArrayValue = @()
	$DataArray = New-Object 'object[,]' -ArgumentList $RowCount,$ColumnCount
	
	Foreach ($DataRow in $DataFile)
	{	
		for ($ArrayColumn = 0 ; $ArrayColumn -lt $ColumnCount ; $ArrayColumn++)
        {
            $DataField = $([string]$DataRow.($header[($ArrayColumn)]))

			# Excel 2003 limit of 1823 characters
            if ($DataField.length -lt 1823) 
                {$DataArray[$ArrayRow,$ArrayColumn] = $DataField}
			# Excel 2007 limit of 8203 characters
            elseif (($Excel_ExOrg.version -ge 12) -and ($DataField.length -lt 8203)) 
                {$DataArray[$ArrayRow,$ArrayColumn] = $DataField} 
			# No known Excel 2010 limit
            elseif ($Excel_ExOrg.version -ge 14)
                {$DataArray[$ArrayRow,$ArrayColumn] = $DataField}
            else
            {
                Write-Host -Object "Number of characters in array member exceeds the version of limitations of this version of Excel" -ForegroundColor Yellow
				Write-Host -Object "-- Writing value to temp variable" -ForegroundColor Yellow
                $DataArray[$ArrayRow,$ArrayColumn] = $DataField
                $BadArrayValue += "$ArrayRow,$ArrayColumn"
            }
        }
		$ArrayRow++
	}

    # Replace big values in $DataArray
    $BadArrayValue_count = $BadArrayValue.count
    $BadArrayValue_Temp = @()
    for ($i = 0 ; $i -lt $BadArrayValue_count ; $i++)
    {
        $BadArray_Split = $badarrayvalue[$i].Split(",")
        $BadArrayValue_Temp += $DataArray[$BadArray_Split[0],$BadArray_Split[1]]
        $DataArray[$BadArray_Split[0],$BadArray_Split[1]] = "**TEMP**"
		Write-Host -Object "-- Replacing long value with **TEMP**" -ForegroundColor Yellow
    }

	$EndCellRow = ($RowCount+1)
	$Data_range = $Worksheet.Range("a2","$EndCellColumn$EndCellRow")
	$Data_range.Value2 = $DataArray

    # Paste big values back into the spreadsheet
    for ($i = 0 ; $i -lt $BadArrayValue_count ; $i++)
    {
        $BadArray_Split = $badarrayvalue[$i].Split(",")
        # Adjust for header and $i=0
        $CellRow = [int]$BadArray_Split[0] + 2
        # Adjust for $i=0
        $CellColumn = [int]$BadArray_Split[1] + 1
           
        $Range = $Worksheet.cells.item($CellRow,$CellColumn) 
        $Range.Value2 = $BadArrayValue_Temp[$i]
		Write-Host -Object "-- Pasting long value back in spreadsheet" -ForegroundColor Yellow
    }    
}
	
	#EndRegion Get-Rbac sheet

#Region Get-RetentionPolicy sheet
Write-Host -Object "---- Starting Get-RetentionPolicy"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "RetentionPolicy"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Global
	$row = 1
	$header = @()
	$header +=  "Name"
	$header +=  "RetentionPolicyTagLinks"
	$header +=  "WhenCreatedUTC"
	$header +=  "WhenChangedUTC"
	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\ExOrg_GetRetentionPolicy.txt") -eq $true)
{	
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\ExOrg_GetRetentionPolicy.txt") 
	# Send the data to the function to process and add to the Excel worksheet
	Process-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

# Format time/date columns
$EndRow = $DataFile.count + 1
# WhenCreatedUTC
$Column_Range = $Worksheet.Range("C1","C$EndRow")
$Column_Range.cells.NumberFormat = "mm/dd/yy hh:mm:ss"
# WhenChangedUTC
$Column_Range = $Worksheet.Range("D1","D$EndRow")
$Column_Range.cells.NumberFormat = "mm/dd/yy hh:mm:ss"
	#EndRegion Get-RetentionPolicy sheet

#Region Get-RetentionPolicyTag sheet
Write-Host -Object "---- Starting Get-RetentionPolicyTag"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "RetentionPolicyTag"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Global
	$row = 1
	$header = @()
	$header +=  "Name"
	$header +=  "MessageClassDisplayName"
	$header +=  "MessageClass"
	$header +=  "RetentionEnabled"
	$header +=  "RetentionAction"
	$header +=  "AgeLimitForRetention"
	$header +=  "MoveToDestinationFolder"
	$header +=  "TriggerForRetention"
	$header +=  "Type"
	$header +=  "WhenCreatedUTC"
	$header +=  "WhenChangedUTC"
	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\ExOrg_GetRetentionPolicyTag.txt") -eq $true)
{	
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\ExOrg_GetRetentionPolicyTag.txt") 
	# Send the data to the function to process and add to the Excel worksheet
	Process-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

# Format time/date columns
$EndRow = $DataFile.count + 1
# AgeLimitForRetention
$Column_Range = $Worksheet.Range("F1","F$EndRow")
$Column_Range.cells.NumberFormat = "dd:hh:mm:ss"
# WhenCreatedUTC
$Column_Range = $Worksheet.Range("J1","J$EndRow")
$Column_Range.cells.NumberFormat = "mm/dd/yy hh:mm:ss"
# WhenChangedUTC
$Column_Range = $Worksheet.Range("K1","K$EndRow")
$Column_Range.cells.NumberFormat = "mm/dd/yy hh:mm:ss"
	#EndRegion Get-RetentionPolicyTag sheet

#Region Get-StorageGroup sheet
Write-Host -Object "---- Starting Get-StorageGroup"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "StorageGroup"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Global
	$row = 1
	$header = @()
	$header +=  "Server"
	$header +=  "Identity"
	$header +=  "CircularLoggingEnabled"
	$header +=  "ZeroDatabasePages"
	$header +=  "Replicated"
	$header +=  "HasLocalCopy"
	$header +=  "StandbyMachines"
	$header +=  "LogFolderPath"
	$header +=  "SystemFolderPath"
	$header +=  "LogFilePrefix"
	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\GetStorageGroup") -eq $true)
{
	foreach ($file in (Get-ChildItem -LiteralPath "$RunLocation\output\ExOrg\GetStorageGroup" | Where-Object {$_.name -match "~~GetStorageGroup"}))
	{
		$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\GetStorageGroup\" + $file)  
	}
	# Send the data to the function to process and add to the Excel worksheet
	Process-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}
	#EndRegion Get-StorageGroup sheet


# Receipient
#Region Get-AdPermission sheet
Write-Host -Object "---- Starting Get-AdPermission"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "AdPermission"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Recipient
	$row = 1
	$header = @()
	$header +=  "Identity"
	$header +=  "User (ACL'ed on Mbx)"
	$header +=  "ExtendedRights"
	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\GetAdPermission") -eq $true)
{
	foreach ($file in (Get-ChildItem -LiteralPath "$RunLocation\output\ExOrg\GetAdPermission" | Where-Object {$_.name -match "~~GetAdPerm"}))
	{
		$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\GetAdPermission\" + $file)
	}
	# Send the data to the function to process and add to the Excel worksheet
	Process-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}
	#EndRegion Get-AdPermission sheet

#Region Get-CalendarProcessing sheet
Write-Host -Object "---- Starting Get-CalendarProcessing"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CalendarProcessing"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Recipient
	$row = 1
	$header = @()
	$header +=  "Mailbox"
	$header +=  "Identity"
	$header +=  "MailboxOwnerId"
	$header +=  "AutomateProcessing"
	$header +=  "AllowConflicts"
	$header +=  "AllowRecurringMeetings"
	$header +=  "ConflictPercentageAllowed"
	$header +=  "MaximumConflictInstances"
	$header +=  "ForwardRequestsToDelegates"
	$header +=  "DeleteAttachments"
	$header +=  "DeleteComments"
	$header +=  "RemovePrivateProperty"
	$header +=  "DeleteSubject"
	$header +=  "DeleteNonCalendarItems"
	$header +=  "TentativePendingApproval"
	$header +=  "ResourceDelegates"
	$header +=  "RequestOutOfPolicy"
	$header +=  "AllRequestOutOfPolicy"
	$header +=  "BookInPolicy"
	$header +=  "AllBookInPolicy"
	$header +=  "RequestInPolicy"
	$header +=  "AllRequestInPolicy"
	$header +=  "AddNewRequestsTentatively"
	$header +=  "ProcessExternalMeetingMessages"
	$header +=  "RemoveForwardedMeetingNotifications"
	$header +=  "IsValid"
	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\GetCalendarProcessing") -eq $true)
{
	foreach ($file in (Get-ChildItem -LiteralPath "$RunLocation\output\ExOrg\GetCalendarProcessing" | Where-Object {$_.name -match "~~GetCalendarProcessing"}))
	{
		$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\GetCalendarProcessing\" + $file) 
	}
	# Send the data to the function to process and add to the Excel worksheet
	Process-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}
	#EndRegion Get-CalendarProcessing sheet	

#Region Get-CASMailbox sheet
Write-Host -Object "---- Starting Get-CASMailbox"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "CASMailbox"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Recipient
	$row = 1
	$header = @()
	$header +=  "Name"
	$header +=  "Identity"
	$header +=  "ServerName"
	$header +=  "ActiveSyncMailboxPolicy"
	$header +=  "ActiveSyncEnabled"
	$header +=  "HasActiveSyncDevicePartnership"
	$header +=  "OwaMailboxPolicy"
	$header +=  "OWAEnabled"
	$header +=  "ECPEnabled"
	$header +=  "EmwsEnabled (Ex2010)"
	$header +=  "PopEnabled"
	$header +=  "ImapEnabled"
	$header +=  "MAPIEnabled"
	$header +=  "MAPIBlockOutlookNonCachedMode"
	$header +=  "MAPIBlockOutlookVersions"
	$header +=  "MAPIBlockOutlookRpcHttp"
	$header +=  "EwsEnabled"
	$header +=  "EwsAllowOutlook"
	$header +=  "EwsAllowMacOutlook"
	$header +=  "EwsAllowEntourage"
	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\GetCASMailbox") -eq $true)
{	
	foreach ($file in (Get-ChildItem -LiteralPath "$RunLocation\output\ExOrg\GetCASMailbox" | Where-Object {$_.name -match "~~GetCASMailbox"}))
	{
		$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\GetCASMailbox\" + $file)
	}
	# Send the data to the function to process and add to the Excel worksheet
	Process-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}
	#EndRegion Get-CASMailbox sheet	

#Region Get-DistributionGroup sheet
Write-Host -Object "---- Starting Get-DistributionGroup"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "DistributionGroup"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Recipient
	$row = 1
	$header = @()
	$header +=  "Name"
	$header +=  "GroupType"
	$header +=  "Member Count"
	$header +=  "ExpansionServer"
	$header +=  "AcceptMessagesOnlyFrom"
	$header +=  "AcceptMessagesOnlyFromDLMembers"
	$header +=  "Alias"
	$header +=  "GrantSendOnBehalfTo"
	$header +=  "HiddenFromAddressListsEnabled"
	$header +=  "MaxSendSize"
	$header +=  "MaxReceiveSize"
	$header +=  "RejectMessagesFrom"
	$header +=  "RejectMessagesFromDLMembers"
	$header +=  "RequireSenderAuthenticationEnabled"
	$header +=  "ManagedBy"
	$header +=  "OrganizationalUnit"
	$header +=  "MemberJoinRestriction"
	$header +=  "MemberDepartRestriction"
	$header +=  "ReportToManagerEnabled"
	$header +=  "ReportToOriginatorEnabled"
	$header +=  "SendOofMessageToOriginatorEnabled"
	$header +=  "AcceptMessagesOnlyFromSendersOrMembers"
	$header +=  "ModeratedBy"
	$header +=  "ModerationEnabled"
	$header +=  "PrimarySmtpAddress"
	$header +=  "RecipientType"
	$header +=  "RecipientTypeDetails"
	$header +=  "RejectMessagesFromSendersOrMembers"
	$header +=  "WhenCreatedUTC"
	$header +=  "WhenChangedUTC"
	$header +=  "IsValid"
	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\ExOrg_GetDistributionGroup.txt") -eq $true)
{	
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\ExOrg_GetDistributionGroup.txt") 
	# Send the data to the function to process and add to the Excel worksheet
	Process-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

# Format time/date columns
$EndRow = $DataFile.count + 1
# WhenCreatedUTC
$Column_Range = $Worksheet.Range("AC1","AC$EndRow")
$Column_Range.cells.NumberFormat = "mm/dd/yy hh:mm:ss"
# WhenChangedUTC
$Column_Range = $Worksheet.Range("AD1","AD$EndRow")
$Column_Range.cells.NumberFormat = "mm/dd/yy hh:mm:ss"
	#EndRegion Get-DistributionGroup sheet

#Region Get-DynamicDistributionGroup sheet
Write-Host -Object "---- Starting Get-DynamicDistributionGroup"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "DynamicDistributionGroup"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Recipient
	$row = 1
	$header = @()
	$header +=  "Name"
	$header +=  "ExpansionServer"
	$header +=  "AcceptMessagesOnlyFrom"
	$header +=  "AcceptMessagesOnlyFromDLMembers"
	$header +=  "Alias"
	$header +=  "GrantSendOnBehalfTo"
	$header +=  "HiddenFromAddressListsEnabled"
	$header +=  "MaxSendSize"
	$header +=  "MaxReceiveSize"
	$header +=  "RejectMessagesFrom"
	$header +=  "RejectMessagesFromDLMembers"
	$header +=  "RequireSenderAuthenticationEnabled"
	$header +=  "ManagedBy"
	$header +=  "OrganizationalUnit"
	$header +=  "RecipientContainer"
	$header +=  "RecipientFilter"
	$header +=  "LdapRecipientFilter"
	$header +=  "IncludedRecipients"
	$header +=  "ReportToManagerEnabled"
	$header +=  "ReportToOriginatorEnabled"
	$header +=  "SendOofMessageToOriginatorEnabled"
	$header +=  "AcceptMessagesOnlyFromSendersOrMembers"
	$header +=  "RejectMessagesFromSendersOrMembers"
	$header +=  "ModeratedBy"
	$header +=  "ModerationEnabled"
	$header +=  "PrimarySmtpAddress"
	$header +=  "RecipientType"
	$header +=  "RecipientTypeDetails"
	$header +=  "WhenCreatedUTC"
	$header +=  "WhenChangedUTC"
	$header +=  "IsValid"
	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\ExOrg_GetDynamicDistributionGroup.txt") -eq $true)
{	
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\ExOrg_GetDynamicDistributionGroup.txt") 
	# Send the data to the function to process and add to the Excel worksheet
	Process-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

# Format time/date columns
$EndRow = $DataFile.count + 1
# WhenCreatedUTC
$Column_Range = $Worksheet.Range("AC1","AC$EndRow")
$Column_Range.cells.NumberFormat = "mm/dd/yy hh:mm:ss"
# WhenChangedUTC
$Column_Range = $Worksheet.Range("AD1","AD$EndRow")
$Column_Range.cells.NumberFormat = "mm/dd/yy hh:mm:ss"
	#EndRegion Get-DynamicDistributionGroup sheet
		
#Region Get-Mailbox sheet
Write-Host -Object "---- Starting Get-Mailbox"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "Mailbox"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Recipient
	$row = 1
	$header = @()
	$header +=  "Mailbox"
	$header +=  "Identity"
	$header +=  "Alias"
	$header +=  "Database"
	$header +=  "RecipientType"
	$header +=  "RecipientTypeDetails"
	$header +=  "PrimarySmtpAddress"
	$header +=  "EmailAddresses"
	$header +=  "HiddenFromAddressListsEnabled"
	$header +=  "GrantSendOnBehalfTo"
	$header +=  "ForwardingSmtpAddress"
	$header +=  "IsMailboxEnabled"
	$header +=  "MailboxMoveStatus"
	$header +=  "SingleItemRecoveryEnabled"
	$header +=  "CalendarVersionStoreDisabled"
	$header +=  "LitigationHoldEnabled"
	$header +=  "LitigationHoldDate"
	$header +=  "LitigationHoldOwner"
	$header +=  "LitigationHoldDuration"
	$header +=  "ThrottlingPolicy"
	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\GetMbx") -eq $true)
{
	foreach ($file in (Get-ChildItem -LiteralPath "$RunLocation\output\ExOrg\GetMbx" | Where-Object {$_.name -match "~~GetMbx"}))
	{
		$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\GetMbx\" + $file)
	}
	# Send the data to the function to process and add to the Excel worksheet
	Process-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

# Format time/date columns
$EndRow = $DataFile.count + 1
# LitigationHoldDate
$Column_Range = $Worksheet.Range("Q1","Q$EndRow")
$Column_Range.cells.NumberFormat = "mm/dd/yy hh:mm:ss"
	#EndRegion Get-Mailbox sheet

#Region Get-MailboxFolderStatistics sheet
Write-Host -Object "---- Starting Get-MailboxFolderStatistics"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "MailboxFolderStatistics"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Recipient
	$row = 1
	$header = @()
	$header +=  "Mailbox"
	$header +=  "Name"
	$header +=  "FolderType"
	$header +=  "Identity"
	$header +=  "ItemsInFolder"
	$header +=  "FolderSize"
	$header +=  "FolderSize (MB)"
	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = [int]$header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\GetMbxFolderStatistics") -eq $true)
{
	foreach ($file in (Get-ChildItem -LiteralPath "$RunLocation\output\ExOrg\GetMbxFolderStatistics" | Where-Object {$_.name -match "~~GetMbxFolderStatistics"}))
	{
		$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\GetMbxFolderStatistics\" + $file) 
	}
	# Send the data to the function to process and add to the Excel worksheet
	Process-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}
	#EndRegion Get-MailboxFolderStatistics sheet

#Region Get-MailboxPermission sheet
Write-Host -Object "---- Starting Get-MailboxPermission"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "MailboxPermission"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Recipient
	$row = 1
	$header = @()
	$header +=  "Mailbox"
	$header +=  "Identity"
	$header +=  "User (ACL'ed on Mbx)"
	$header +=  "AccessRights"
	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\GetMbxPermission") -eq $true)
{
	foreach ($file in (Get-ChildItem -LiteralPath "$RunLocation\output\ExOrg\GetMbxPermission" | Where-Object {$_.name -match "~~GetMbxPerm"}))
	{
		$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\GetMbxPermission\" + $file)
	}
	# Send the data to the function to process and add to the Excel worksheet
	Process-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}
	#EndRegion Get-MailboxPermission sheet

#Region Get-MailboxStatistics sheet
Write-Host -Object "---- Starting Get-MailboxStatistics"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "MailboxStatistics"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Recipient
	$row = 1
	$header = @()
	$header +=  "Mailbox"
	$header +=  "DisplayName"
	$header +=  "ServerName"
	$header +=  "Database"
	$header +=  "ItemCount"
	$header +=  "TotalItemSize"
	$header +=  "TotalItemSize (MB)"
	$header +=  "TotalDeletedItemSize"
	$header +=  "TotalDeletedItemSize (MB)"
	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\GetMbxStatistics") -eq $true)
{
	foreach ($file in (Get-ChildItem -LiteralPath "$RunLocation\output\ExOrg\GetMbxStatistics" | Where-Object {$_.name -match "~~GetMbxStatistics"}))
	{
		$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\GetMbxStatistics\" + $file) 
	}
	# Send the data to the function to process and add to the Excel worksheet
	Process-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}
	#EndRegion Get-MailboxStatistics sheet

#Region Get-PublicFolder sheet
Write-Host -Object "---- Starting Get-PublicFolder"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "PublicFolder"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Recipient
	$row = 1
	$header = @()
	$header +=  "OriginatingServer"
	$header +=  "Name"
	$header +=  "ParentPath"
	$header +=  "UseDatabaseAgeDefaults"
	$header +=  "AgeLimit"
	$header +=  "UseDatabaseQuotaDefaults"
	$header +=  "StorageQuota"
	$header +=  "UseDatabaseReplicationSchedule"
	$header +=  "UseDatabaseRetentionDefaults"
	$header +=  "HasSubFolders"
	$header +=  "MailEnabled"
	$header +=  "MaxItemSize"
	$header +=  "Replicas"
	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\GetPublicFolder") -eq $true)
{
	foreach ($file in (Get-ChildItem -LiteralPath "$RunLocation\output\ExOrg\GetPublicFolder" | Where-Object {$_.name -match "~~GetPublicFolder"}))
	{
		$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\GetPublicFolder\" + $file) 
	}
	# Send the data to the function to process and add to the Excel worksheet
	Process-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}
	#EndRegion Get-PublicFolder sheet
		
#Region Get-PublicFolderStatistics sheet
Write-Host -Object "---- Starting Get-PublicFolderStatistics"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "PublicFolderStatistics"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Recipient
	$row = 1
	$header = @()
	$header +=  "OriginatingServer"
	$header +=  "AdminDisplayName"
	$header +=  "FolderPath"
	$header +=  "ItemCount"
	$header +=  "TotalItemSize"
	$header +=  "TotalItemSize (MB)"
	$header +=  "CreationTime"				# Column G
	$header +=  "LastModificationTime"		# Column H
	$header +=  "LastAccessTime"			# Column I
	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\GetPublicFolderStats") -eq $true)
{
	foreach ($file in (Get-ChildItem -LiteralPath "$RunLocation\output\ExOrg\GetPublicFolderStats" | Where-Object {$_.name -match "~~GetPublicFolderStats"}))
	{
		$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\GetPublicFolderStats\" + $file)
	}
	# Send the data to the function to process and add to the Excel worksheet
	Process-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

# Format time/date columns
$EndRow = $DataFile.count + 1
# CreationTime
$Column_Range = $Worksheet.Range("G1","G$EndRow")
$Column_Range.cells.NumberFormat = "mm/dd/yy hh:mm:ss"
# LastModificationTime
$Column_Range = $Worksheet.Range("H1","H$EndRow")
$Column_Range.cells.NumberFormat = "mm/dd/yy hh:mm:ss"
# LastAccessTime
$Column_Range = $Worksheet.Range("I1","I$EndRow")
$Column_Range.cells.NumberFormat = "mm/dd/yy hh:mm:ss"

	#EndRegion Get-PublicFolderStatistics sheet	

#Region Quota sheet
Write-Host -Object "---- Starting Quota"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "Quota"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Recipient
	$row = 1
	$header = @()
	$header +=  "ServerName"
	$header +=  "Alias"
	$header +=  "UseDatabaseQuotaDefaults"
	$header +=  "IssueWarningQuota"
	$header +=  "ProhibitSendQuota"
	$header +=  "ProhibitSendReceiveQuota"
	$header +=  "RecoverableItemsQuota"
	$header +=  "RecoverableItemsWarningQuota"
	$header +=  "LitigationHoldEnabled"
	$header +=  "RetentionHoldEnabled"
	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\Quota") -eq $true)
{	
	foreach ($file in (Get-ChildItem -LiteralPath "$RunLocation\output\ExOrg\Quota" | Where-Object {$_.name -match "~~Quota"}))
	{
		$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\Quota\" + $file)  
	}
	$RowCount = $DataFile.Count
	# Not using the Process-Datafile function because Quota needs special data handling for formatting
	$ArrayRow = 0
	$DataArray = New-Object 'object[,]' -ArgumentList $RowCount,$ColumnCount
	Foreach ($DataRow in $DataFile)
	{
		$DataField = $DataRow.Split("`t")
		for ($ArrayColumn=0;$ArrayColumn -le 2;$ArrayColumn++)
		{
            # Excel chokes if field starts with = so we'll prepend the ' to the string if it does
            If ($DataField[$ArrayColumn].substring(0,1) -eq "=") {$DataField[$ArrayColumn] = "'"+$DataField[$ArrayColumn]}		

			$DataArray[$ArrayRow,$ArrayColumn] = $DataField[$ArrayColumn]
		}
		if ($DataField[2] -eq "TRUE")
		{
			$DataArray[$ArrayRow,3] =  "- - -"
			$DataArray[$ArrayRow,4] =  "- - -"
			$DataArray[$ArrayRow,5] =  "- - -"
		}
		else
		{
			$DataArray[$ArrayRow,3] =  $DataField[3]		
			$DataArray[$ArrayRow,4] =  $DataField[4]		
			$DataArray[$ArrayRow,5] =  $DataField[5]		
		}		
		for ($ArrayColumn=6;$ArrayColumn -le $ColumnCount-1;$ArrayColumn++)
		{
			$DataArray[$ArrayRow,$ArrayColumn] = $DataField[$ArrayColumn]
		}
		#write-host $ArrayRow " of " $RowCount

		$ArrayRow++
	}

	$EndCellRow = ($RowCount+1)
	$Data_range = $Worksheet.Range("a2","$EndCellColumn$EndCellRow")
	$Data_range.Value2 = $DataArray	
}
	#EndRegion Quota sheet
	
# Transport
#Region Get-AcceptedDomain sheet
Write-Host -Object "---- Starting Get-AcceptedDomain"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "AcceptedDomain"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Transport
	$row = 1
	$header = @()
	$header +=  "SmtpDomain"
	$header +=  "DomainType"
	$header +=  "Default"
	$header +=  "Name"
	$header +=  "MatchSubDomains"
	$header +=  "IsValid"
	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\ExOrg_GetAcceptedDomain.txt") -eq $true)
{	
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\ExOrg_GetAcceptedDomain.txt") 
	# Send the data to the function to process and add to the Excel worksheet
	Process-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}
	#endRegion Get-AcceptedDomain sheet

#Region Get-AdSite sheet
Write-Host -Object "---- Starting Get-AdSite"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "AdSite"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Transport
	$row = 1
	$header = @()
	$header +=  "Name"
	$header +=  "Identity"
	$header +=  "HubSiteEnabled"
	$header +=  "InboundMailEnabled"
	$header +=  "IsValid"
	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\ExOrg_GetAdSite.txt") -eq $true)
{	
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\ExOrg_GetAdSite.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Process-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}
	#EndRegion Get-AdSite sheet

#Region Get-AdSiteLink sheet
Write-Host -Object "---- Starting Get-AdSiteLink"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "AdSiteLink"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Transport
	$row = 1
	$header = @()
	$header +=  "Name"
	$header +=  "Cost"
	$header +=  "ADCost"
	$header +=  "ExchangeCost"
	$header +=  "MaxMessageSize"
	$header +=  "Sites"
	$header +=  "IsValid"
	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\ExOrg_GetAdSiteLink.txt") -eq $true)
{	
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\ExOrg_GetAdSiteLink.txt")  
	# Send the data to the function to process and add to the Excel worksheet
	Process-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}
	#EndRegion Get-AdSiteLink sheet
	
#Region Get-ContentFilterConfig sheet
Write-Host -Object "---- Starting Get-ContentFilterConfig"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "ContentFilterConfig"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Transport
	$row = 1
	$header = @()
	$header +=  "Identity"
	$header +=  "Enabled"
	$header +=  "ExternalMailEnabled"
	$header +=  "InternalMailEnabled"
	$header +=  "OutlookEmailPostmarkValidationEnabled"
	$header +=  "BypassedRecipients"
	$header +=  "QuarantineMailbox"
	$header +=  "SCLRejectThreshold"
	$header +=  "SCLRejectEnabled"
	$header +=  "SCLDeleteThreshold"
	$header +=  "SCLDeleteEnabled"
	$header +=  "SCLQuarantineThreshold"
	$header +=  "SCLQuarantineEnabled"
	$header +=  "SCLJunkThreshold"
	$header +=  "BypassedSenders"
	$header +=  "BypassedSenderDomains"
	$header +=  "ExchangeVersion"
	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\ExOrg_GetContentFilterConfig.txt") -eq $true)
{	
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\ExOrg_GetContentFilterConfig.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Process-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}
	#EndRegion Get-ContentFilterConfig sheet

#Region Get-ReceiveConnector sheet
Write-Host -Object "---- Starting Get-ReceiveConnector"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "ReceiveConnector"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Transport
	$row = 1
	$header = @()
	$header +=  "Server"
	$header +=  "Name"
	$header +=  "Enabled"
	$header +=  "Banner"
	$header +=  "ConnectionTimeout"
	$header +=  "MaxHopCount"
	$header +=  "MaxMessageSize"
	$header +=  "MaxRecipientsPerMessage"
	$header +=  "AuthMechanism"
	$header +=  "PermissionGroups"
	$header +=  "RemoteIPRanges"
	$header +=  "RequireTLS"
	$header +=  "TarpitInterval"		# Column M
	$header +=  "Bindings"
	$header +=  "ChunkingEnabled"
	$header +=  "ConnectionInactivityTimeout"
	$header +=  "MessageRateLimit"
	$header +=  "MessageSourceLimit"
	$header +=  "MaxInboundConnection"
	$header +=  "MaxInboundConnectionPerSource"
	$header +=  "TransportRole"
	
	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\ExOrg_GetReceiveConnector.txt") -eq $true)
{	
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\ExOrg_GetReceiveConnector.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Process-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

# Format time/date columns
$EndRow = $DataFile.count + 1
# ConnectionTimeout
$Column_Range = $Worksheet.Range("E1","E$EndRow")
$Column_Range.cells.NumberFormat = "hh:mm:ss"
# TarpitInterval
$Column_Range = $Worksheet.Range("M1","M$EndRow")
$Column_Range.cells.NumberFormat = "hh:mm:ss"
	
	#EndRegion Get-ReceiveConnector sheet

#Region Get-RemoteDomain sheet
Write-Host -Object "---- Starting Get-RemoteDomain"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "RemoteDomain"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Transport
	$row = 1
	$header = @()
	$header +=  "Identity"
	$header +=  "DomainName"
	$header +=  "IsInternal"
	$header +=  "TrustedMailOutboundEnabled"
	$header +=  "TrustedMailInboundEnabled"
	$header +=  "AllowedOOFType"
	$header +=  "AutoForwardEnabled"
	$header +=  "AutoReplyEnabled"
	$header +=  "DeliveryReportEnabled"
	$header +=  "MeetingForwardNotificationEnabled"
	$header +=  "NDREnabled"
	$header +=  "DisplaySenderName"
	$header +=  "CharacterSet"
	$header +=  "NonMimeCharacterSet"
	$header +=  "ContentType"
	$header +=  "TNEFEnabled"
	$header +=  "LineWrapSize"
	$header +=  "UseSimpleDisplayName"

	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\ExOrg_GetRemoteDomain.txt") -eq $true)
{	
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\ExOrg_GetRemoteDomain.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Process-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}
	#EndRegion Get-RemoteDomain sheet

#Region Get-RoutingGroupConnector sheet
Write-Host -Object "---- Starting Get-RoutingGroupConnector"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "RoutingGroupConnector"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Transport
	$row = 1
	$header = @()
	$header +=  "Name"
	$header +=  "Cost"
	$header +=  "TargetRoutingGroup"
	$header +=  "TargetTransportServers"
	$header +=  "SourceRoutingGroup"
	$header +=  "SourceTransportServers"
	$header +=  "PublicFolderReferralsEnabled"
	$header +=  "MaxMessageSize"
	$header +=  "IsValid"
	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\ExOrg_GetRoutingGroupConnector.txt") -eq $true)
{	
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\ExOrg_GetRoutingGroupConnector.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Process-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}
	#EndRegion Get-RoutingGroupConnector sheet

#Region Get-SendConnector sheet
Write-Host -Object "---- Starting Get-SendConnector"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "SendConnector"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Transport
	$row = 1
	$header = @()
	$header +=  "Name"
	$header +=  "Fqdn"
	$header +=  "Enabled"
	$header +=  "AddressSpaces"
	$header +=  "ConnectedDomains"
	$header +=  "ConnectionInactivityTimeOut"
	$header +=  "DNSRoutingEnabled"
	$header +=  "DomainSecureEnabled"
	$header +=  "IgnoreStartTLS"
	$header +=  "IsScopedConnector"
	$header +=  "IsSmtpConnector"
	$header +=  "MaxMessageSize"
	$header +=  "Port"
	$header +=  "ProtocolLoggingLevel"
	$header +=  "RequireTLS"
	$header +=  "SmartHostAuthMechanism"
	$header +=  "SmartHosts"
	$header +=  "SourceIPAddress"
	$header +=  "SourceRoutingGroup"
	$header +=  "SourceTransportServers"
	$header +=  "TlsAuthLevel"
	$header +=  "TlsDomain"
	$header +=  "CloudServicesMailEnabled"
	$header +=  "Comment"
	$header +=  "SmtpMaxMessagesPerConnection"

	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\ExOrg_GetSendConnector.txt") -eq $true)
{	
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\ExOrg_GetSendConnector.txt")
	$RowCount = $DataFile.Count
	$ArrayRow = 0
	$BadArrayValue = @()
	$DataArray = New-Object 'object[,]' -ArgumentList $RowCount,$ColumnCount
	Foreach ($DataRow in $DataFile)
	{
		$DataField = $DataRow.Split("`t")
		for ($ArrayColumn = 0 ; $ArrayColumn -lt $ColumnCount ; $ArrayColumn++)
		{
			# Excel 2003 limit of 1823 characters
            if ($DataField[$ArrayColumn].length -lt 1823) 
                {$DataArray[$ArrayRow,$ArrayColumn] = $DataField[$ArrayColumn]}
			# Excel 2007 limit of 8203 characters
            elseif (($Excel_ExOrg.version -ge 12) -and ($DataField[$ArrayColumn].length -lt 8203)) 
                {$DataArray[$ArrayRow,$ArrayColumn] = $DataField[$ArrayColumn]} 
			# No known Excel 2010 limit
            elseif ($Excel_ExOrg.version -ge 14)
                {$DataArray[$ArrayRow,$ArrayColumn] = $DataField[$ArrayColumn]}
            else
            {
                Write-Host -Object "Number of characters in array member exceeds the version of limitations of this version of Excel" -ForegroundColor Yellow
				Write-Host -Object "-- Writing value to temp variable" -ForegroundColor Yellow
                $DataArray[$ArrayRow,$ArrayColumn] = $DataField[$ArrayColumn]
                $BadArrayValue += "$ArrayRow,$ArrayColumn"
            }
		}
		$ArrayRow++
	}

    # Replace big values in $DataArray
    $BadArrayValue_count = $BadArrayValue.count
    $BadArrayValue_Temp = @()
    for ($i = 0 ; $i -lt $BadArrayValue_count ; $i++)
    {
        $BadArray_Split = $badarrayvalue[$i].Split(",")
        $BadArrayValue_Temp += $DataArray[$BadArray_Split[0],$BadArray_Split[1]]
        $DataArray[$BadArray_Split[0],$BadArray_Split[1]] = "**TEMP**"
		Write-Host -Object "-- Replacing long value with **TEMP**" -ForegroundColor Yellow
    }

	$EndCellRow = ($RowCount+1)
	$Data_range = $Worksheet.Range("a2","$EndCellColumn$EndCellRow")
	$Data_range.Value2 = $DataArray

    # Paste big values back into the spreadsheet
    for ($i = 0 ; $i -lt $BadArrayValue_count ; $i++)
    {
        $BadArray_Split = $badarrayvalue[$i].Split(",")
        # Adjust for header and $i=0
        $CellRow = [int]$BadArray_Split[0] + 2
        # Adjust for $i=0
        $CellColumn = [int]$BadArray_Split[1] + 1
           
        $Range = $Worksheet.cells.item($CellRow,$CellColumn) 
        $Range.Value2 = $BadArrayValue_Temp[$i]
		Write-Host -Object "-- Pasting long value back in spreadsheet" -ForegroundColor Yellow
    }    
}

# Format time/date columns
$EndRow = $DataFile.count + 1
# ConnectionInactivityTimeout
$Column_Range = $Worksheet.Range("F1","F$EndRow")
$Column_Range.cells.NumberFormat = "hh:mm:ss"

	#EndRegion Get-SendConnector sheet

#Region Get-TransportConfig sheet
Write-Host -Object "---- Starting Get-TransportConfig"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "TransportConfig"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Transport
	$row = 1
	$header = @()
	$header += "AdminDisplayName"
    $header +="ClearCategories"
    $header += "ConvertDisclaimerWrapperToEml"
    $header += "ConvertReportToMessage"
   	$header += "DSNConversionMode"
    $header += "ExternalDelayDsnEnabled"
    $header += "ExternalDsnDefaultLanguage"
    $header += "ExternalDsnLanguageDetectionEnabled"
    $header += "ExternalDsnMaxMessageAttachSize"
    $header += "ExternalDsnReportingAuthority"
    $header += "ExternalDsnSendHtml"
    $header += "ExternalPostmasterAddress"
    $header += "GenerateCopyOfDSNFor"
    $header += "Guid"
    $header += "HeaderPromotionModeSetting"
    $header += "HygieneSuite"
    $header += "Identity"
    $header += "InternalDelayDsnEnabled"
    $header += "InternalDsnDefaultLanguage"
    $header += "InternalDsnLanguageDetectionEnabled"
    $header += "InternalDsnMaxMessageAttachSize"
    $header += "InternalDsnReportingAuthority"
    $header += "InternalDsnSendHtml"
    $header += "InternalSMTPServers"
    $header += "JournalingReportNdrTo"
    $header += "LegacyJournalingMigrationEnabled"
    $header += "MaxDumpsterSizePerDatabase"
    $header += "MaxDumpsterTime"
    $header += "MaxReceiveSize"
    $header += "MaxRecipientEnvelopeLimit"
    $header += "MaxSendSize"
    $header += "MigrationEnabled"
    $header += "OpenDomainRoutingEnabled"
    $header += "OrganizationFederatedMailbox"
    $header += "OrganizationId"
    $header += "OriginatingServer"
    $header += "OtherWellKnownObjects"
    $header += "PreserveReportBodypart"
    $header += "Rfc2231EncodingEnabled"
    $header += "ShadowHeartbeatRetryCount"
    $header += "ShadowHeartbeatTimeoutInterval"
    $header += "ShadowMessageAutoDiscardInterval"
    $header += "ShadowRedundancyEnabled"
    $header += "SupervisionTags"
    $header += "TLSReceiveDomainSecureList"
    $header += "TLSSendDomainSecureList"
    $header += "VerifySecureSubmitEnabled"
    $header += "VoicemailJournalingEnabled"
    $header += "WhenChangedUTC"
    $header += "WhenCreatedUTC"
    $header += "Xexch50Enabled"
	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\ExOrg_GetTransportConfig.txt") -eq $true)
{	
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\ExOrg_GetTransportConfig.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Process-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

# Format time/date columns
$EndRow = $DataFile.count + 1
# ShadowHeartbeatTimeoutInterval
$Column_Range = $Worksheet.Range("AO1","AO$EndRow")
$Column_Range.cells.NumberFormat = "hh:mm:ss"
# ShadowMessageAutoDiscardInterval
$Column_Range = $Worksheet.Range("AP1","AP$EndRow")
$Column_Range.cells.NumberFormat = "dd:hh:mm:ss"
# WhenCreatedUTC
$Column_Range = $Worksheet.Range("AW1","AW$EndRow")
$Column_Range.cells.NumberFormat = "mm/dd/yy hh:mm:ss"
# WhenChangedUTC
$Column_Range = $Worksheet.Range("AX1","AX$EndRow")
$Column_Range.cells.NumberFormat = "mm/dd/yy hh:mm:ss"
	#EndRegion Get-TransportConfig sheet

#Region Get-TransportRule sheet
Write-Host -Object "---- Starting Get-TransportRule"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "TransportRule"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Transport
	$row = 1
	$header = @()
	$header +=  "Identity"
	$header +=  "Priority"
	$header +=  "Comments"
	$header +=  "Description"
	$header +=  "RuleVersion"
	$header +=  "State"
	$header +=  "WhenChanged"
	
	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\ExOrg_GetTransportRule.xml") -eq $true)
{	
	$DataFile = Import-Clixml "$RunLocation\output\ExOrg\ExOrg_GetTransportRule.xml"
	$RowCount = $DataFile.Count
	$ArrayRow = 0
	$BadArrayValue = @()
	$DataArray = New-Object 'object[,]' -ArgumentList $RowCount,$ColumnCount
	
	Foreach ($DataRow in $DataFile)
	{	
		for ($ArrayColumn = 0 ; $ArrayColumn -lt $ColumnCount ; $ArrayColumn++)
        {
            $DataField = $([string]$DataRow.($header[($ArrayColumn)]))

			# Excel 2003 limit of 1823 characters
            if ($DataField.length -lt 1823) 
                {$DataArray[$ArrayRow,$ArrayColumn] = $DataField}
			# Excel 2007 limit of 8203 characters
            elseif (($Excel_ExOrg.version -ge 12) -and ($DataField.length -lt 8203)) 
                {$DataArray[$ArrayRow,$ArrayColumn] = $DataField} 
			# No known Excel 2010 limit
            elseif ($Excel_ExOrg.version -ge 14)
                {$DataArray[$ArrayRow,$ArrayColumn] = $DataField}
            else
            {
                Write-Host -Object "Number of characters in array member exceeds the version of limitations of this version of Excel" -ForegroundColor Yellow
				Write-Host -Object "-- Writing value to temp variable" -ForegroundColor Yellow
                $DataArray[$ArrayRow,$ArrayColumn] = $DataField
                $BadArrayValue += "$ArrayRow,$ArrayColumn"
            }
        }
		$ArrayRow++
	}

    # Replace big values in $DataArray
    $BadArrayValue_count = $BadArrayValue.count
    $BadArrayValue_Temp = @()
    for ($i = 0 ; $i -lt $BadArrayValue_count ; $i++)
    {
        $BadArray_Split = $badarrayvalue[$i].Split(",")
        $BadArrayValue_Temp += $DataArray[$BadArray_Split[0],$BadArray_Split[1]]
        $DataArray[$BadArray_Split[0],$BadArray_Split[1]] = "**TEMP**"
		Write-Host -Object "-- Replacing long value with **TEMP**" -ForegroundColor Yellow
    }

	$EndCellRow = ($RowCount+1)
	$Data_range = $Worksheet.Range("a2","$EndCellColumn$EndCellRow")
	$Data_range.Value2 = $DataArray

    # Paste big values back into the spreadsheet
    for ($i = 0 ; $i -lt $BadArrayValue_count ; $i++)
    {
        $BadArray_Split = $badarrayvalue[$i].Split(",")
        # Adjust for header and $i=0
        $CellRow = [int]$BadArray_Split[0] + 2
        # Adjust for $i=0
        $CellColumn = [int]$BadArray_Split[1] + 1
           
        $Range = $Worksheet.cells.item($CellRow,$CellColumn) 
        $Range.Value2 = $BadArrayValue_Temp[$i]
		Write-Host -Object "-- Pasting long value back in spreadsheet" -ForegroundColor Yellow
    }    
}

# Format time/date columns
$EndRow = $DataFile.count + 1
# WhenChangedUTC
$Column_Range = $Worksheet.Range("G1","G$EndRow")
$Column_Range.cells.NumberFormat = "mm/dd/yy hh:mm:ss"	
	
	#EndRegion Get-TransportRule sheet

#Region Get-TransportServer sheet
Write-Host -Object "---- Starting Get-TransportServer"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "TransportServer"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Transport
	$row = 1
	$header = @()
	$header +=  "Name"
	$header +=  "AntispamAgentsEnabled"
	$header +=  "AntispamUpdatesEnabled"
	$header +=  "DelayNotificationTimeout"		# Column D
	$header +=  "MessageExpirationTimeout"		# Column E
	$header +=  "MessageRetryInterval"			# Column F
	$header +=  "ExternalDNSAdapterEnabled"
	$header +=  "ExternalDNSServers"
	$header +=  "InternalDNSAdapterEnabled"
	$header +=  "InternalDNSServers"
	$header +=  "MaxOutboundConnections"
	$header +=  "MaxPerDomainOutboundConnections"
	$header +=  "TransientFailureRetryCount"
	$header +=  "ConnectivityLogEnabled"
	$header +=  "MessageTrackingLogEnabled"
	$header +=  "MessageTrackingLogSubjectLoggingEnabled"
	$header +=  "EdgeTransport.exe.config path"
	$header +=  "DatabaseMaxCacheSize value"
	$header +=  "ShadowRedundancyPromotionEnabled value"
	$header +=  "QueueDatabasePath value"
	$header +=  "QueueDatabaseLoggingPath value"
	$header +=  "ActiveUserStatisticsLogPath"
	$header +=  "ConnectivityLogPath"
	$header +=  "MessageTrackingLogPath"
	$header +=  "ReceiveProtocolLogPath"
	$header +=  "RoutingTableLogPath"
	$header +=  "SendProtocolLogPath"
	$header +=  "ServerStatisticsLogPath"
	
	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\ExOrg_GetTransportSvr.txt") -eq $true)
{	
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\ExOrg_GetTransportSvr.txt")  
	# Send the data to the function to process and add to the Excel worksheet
	Process-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}

# Format time/date columns
$EndRow = $DataFile.count + 1
# DelayNotificationTimeout
$Column_Range = $Worksheet.Range("D1","D$EndRow")
$Column_Range.cells.NumberFormat = "hh:mm:ss"
# MessageExpirationTimeout
$Column_Range = $Worksheet.Range("E1","E$EndRow")
$Column_Range.cells.NumberFormat = "dd.hh:mm:ss"
# MessageRetryInterval
$Column_Range = $Worksheet.Range("F1","F$EndRow")
$Column_Range.cells.NumberFormat = "hh:mm:ss"
	
	#EndRegion Get-TransportServer sheet

# Um
#Region Get-UmAutoAttendant sheet
Write-Host -Object "---- Starting Get-UmAutoAttendant"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "UmAutoAttendant"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Um
	$row = 1
	$header = @()
	$header +=  "Name"
	$header +=  "Identity"
	$header +=  "SpeechEnabled"
	$header +=  "AllowDialPlanSubscribers"
	$header +=  "AllowExtensions"
	$header +=  "AllowedInCountryOrRegionGroups"
	$header +=  "AllowedInternationalGroups"
	$header +=  "CallSomeoneEnabled"
	$header +=  "ContactScope"
	$header +=  "ContactAddressList"
	$header +=  "SendVoiceMsgEnabled"
	$header +=  "BusinessHourSchedule"
	$header +=  "PilotIdentifierList"
	$header +=  "UmDialPlan"
	$header +=  "DtmfFallbackAutoAttendant"
	$header +=  "HolidaySchedule"
	$header +=  "TimeZone"
	$header +=  "TimeZoneName"
	$header +=  "MatchedNameSelectionMethod"
	$header +=  "BusinessLocation"
	$header +=  "WeekStartDay"
	$header +=  "Status"
	$header +=  "Language"
	$header +=  "OperatorExtension"
	$header +=  "InfoAnnouncementFilename"
	$header +=  "InfoAnnouncementEnabled"
	$header +=  "NameLookupEnabled"
	$header +=  "StarOutToDialPlanEnabled"
	$header +=  "ForwardCallsToDefaultMailbox"
	$header +=  "DefaultMailbox"
	$header +=  "BusinessName"
	$header +=  "BusinessHoursWelcomeGreetingFilename"
	$header +=  "BusinessHoursWelcomeGreetingEnabled"
	$header +=  "BusinessHoursMainMenuCustomPromptFilename"
	$header +=  "BusinessHoursMainMenuCustomPromptEnabled"
	$header +=  "BusinessHoursTransferToOperatorEnabled"
	$header +=  "BusinessHoursKeyMapping"
	$header +=  "BusinessHoursKeyMappingEnabled"
	$header +=  "AfterHoursWelcomeGreetingFilename"
	$header +=  "AfterHoursWelcomeGreetingEnabled"
	$header +=  "AfterHoursMainMenuCustomPromptFilename"
	$header +=  "AfterHoursMainMenuCustomPromptEnabled"
	$header +=  "AfterHoursTransferToOperatorEnabled"
	$header +=  "AfterHoursKeyMapping"
	$header +=  "AfterHoursKeyMappingEnabled"
	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\ExOrg_GetUmAutoAttendant.txt") -eq $true)
{	
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\ExOrg_GetUmAutoAttendant.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Process-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}
	#EndRegion Get-UmAutoAttendant sheet

#Region Get-UmDialPlan sheet
Write-Host -Object "---- Starting Get-UmDialPlan"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "UmDialPlan"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Um
	$row = 1
	$header = @()
	$header +=  "Name"
	$header +=  "Identity"
	$header +=  "NumberOfDigitsInExtension"
	$header +=  "LogonFailuresBeforeDisconnect"
	$header +=  "AccessTelephoneNumbers"
	$header +=  "FaxEnabled"
	$header +=  "InputFailuresBeforeDisconnect"
	$header +=  "OutsideLineAccessCode"
	$header +=  "DialByNamePrimary"
	$header +=  "DialByNameSecondary"
	$header +=  "AudioCodec"
	$header +=  "AvailableLanguages"
	$header +=  "DefaultLanguage"
	$header +=  "VoIPSecurity"
	$header +=  "MaxCallDuration"
	$header +=  "MaxRecordingDuration"
	$header +=  "RecordingIdleTimeout"
	$header +=  "PilotIdentifierList"
	$header +=  "UMServers"
	$header +=  "UMMailboxPolicies"
	$header +=  "UMAutoAttendants"
	$header +=  "WelcomeGreetingEnabled"
	$header +=  "AutomaticSpeechRecognitionEnabled"
	$header +=  "PhoneContext"
	$header +=  "WelcomeGreetingFilename"
	$header +=  "InfoAnnouncementFilename"
	$header +=  "OperatorExtension"
	$header +=  "DefaultOutboundCallingLineId"
	$header +=  "Extension"
	$header +=  "MatchedNameSelectionMethod"
	$header +=  "InfoAnnouncementEnabled"
	$header +=  "InternationalAccessCode"
	$header +=  "NationalNumberPrefix"
	$header +=  "InCountryOrRegionNumberFormat"
	$header +=  "InternationalNumberFormat"
	$header +=  "CallSomeoneEnabled"
	$header +=  "ContactScope"
	$header +=  "ContactAddressList"
	$header +=  "SendVoiceMsgEnabled"
	$header +=  "UMAutoAttendant"
	$header +=  "AllowDialPlanSubscribers"
	$header +=  "AllowExtensions"
	$header +=  "AllowedInCountryOrRegionGroups"
	$header +=  "AllowedInternationalGroups"
	$header +=  "ConfiguredInCountryOrRegionGroups"
	$header +=  "LegacyPromptPublishingPoint"
	$header +=  "ConfiguredInternationalGroups"
	$header +=  "UMIPGateway"
	$header +=  "URIType"
	$header +=  "SubscriberType"
	$header +=  "GlobalCallRoutingScheme"
	$header +=  "TUIPromptEditingEnabled"
	$header +=  "CallAnsweringRulesEnabled"
	$header +=  "SipResourceIdentifierRequired"
	$header +=  "FDSPollingInterval"
	$header +=  "EquivalentDialPlanPhoneContexts"
	$header +=  "NumberingPlanFormats"
	$header +=  "AllowHeuristicADCallingLineIdResolution"
	$header +=  "CountryOrRegionCode"
	$header +=  "ExchangeVersion"
	$header +=  "IsValid"
	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\ExOrg_GetUmDialPlan.txt") -eq $true)
{	
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\ExOrg_GetUmDialPlan.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Process-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}
	#EndRegion Get-UmDialPlan sheet

#Region Get-UmIpGateway sheet
Write-Host -Object "---- Starting Get-UmIpGateway"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "UmIpGateway"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Um
	$row = 1
	$header = @()
	$header +=  "Name"
	$header +=  "Identity"
	$header +=  "Address"
	$header +=  "OutcallsAllowed"
	$header +=  "Status"
	$header +=  "Port"
	$header +=  "Simulator"
	$header +=  "DelayedSourcePartyInfoEnabled"
	$header +=  "MessageWaitingIndicatorAllowed"
	$header +=  "HuntGroups"
	$header +=  "GlobalCallRoutingScheme"
	$header +=  "ForwardingAddress"
	$header +=  "NumberOfDigitsInExtension"
	$header +=  "IsValid"
	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\ExOrg_GetUmIpGateway.txt") -eq $true)
{	
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\ExOrg_GetUmIpGateway.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Process-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}
	#EndRegion Get-UmIpGateway sheet

#Region Get-UmMailbox sheet
Write-Host -Object "---- Starting Get-UmMailbox"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "UmMailbox"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Um
	$row = 1
	$header = @()
	$header +=  "Name"
	$header +=  "Identity"
	$header +=  "EmailAddresses"
	$header +=  "UMAddresses"
	$header +=  "LegacyExchangeDN"
	$header +=  "LinkedMasterAccount"
	$header +=  "PrimarySmtpAddress"
	$header +=  "SamAccountName"
	$header +=  "ServerLegacyDN"
	$header +=  "ServerName"
	$header +=  "UMDtmfMap"
	$header +=  "UMEnabled"
	$header +=  "TUIAccessToCalendarEnabled"
	$header +=  "FaxEnabled"
	$header +=  "TUIAccessToEmailEnabled"
	$header +=  "SubscriberAccessEnabled"
	$header +=  "MissedCallNotificationEnabled"
	$header +=  "UMSMSNotificationOption"
	$header +=  "PinlessAccessToVoiceMailEnabled"
	$header +=  "AnonymousCallersCanLeaveMessages"
	$header +=  "AutomaticSpeechRecognitionEnabled"
	$header +=  "PlayOnPhoneEnabled"
	$header +=  "CallAnsweringRulesEnabled"
	$header +=  "AllowUMCallsFromNonUsers"
	$header +=  "OperatorNumber"
	$header +=  "PhoneProviderId"
	$header +=  "UMDialPlan"
	$header +=  "UMMailboxPolicy"
	$header +=  "Extensions"
	$header +=  "CallAnsweringAudioCodec"
	$header +=  "SIPResourceIdentifier"
	$header +=  "PhoneNumber"
	$header +=  "AirSyncNumbers"
	$header +=  "IsValid"
	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\GetUmMailbox") -eq $true)
{	
	foreach ($file in (Get-ChildItem -LiteralPath "$RunLocation\output\ExOrg\GetUmMailbox" | Where-Object {$_.name -match "~~GetUmMailbox"}))
	{
		$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\GetUmMailbox\" + $file)
	}
	# Send the data to the function to process and add to the Excel worksheet
	Process-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}
	#EndRegion Get-UmMailbox sheet

##Region Get-UmMailboxConfiguration sheet
#Write-Host -Object "---- Starting Get-UmMailboxConfiguration"
#	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
#	$Worksheet.name = "UmMailboxConfiguration"
#	$Worksheet.Tab.ColorIndex = $intColorIndex_Um
#	$row = 1
#	$header = @()
#	$header +=  "Identity"
#	$header +=  "Greeting"
#	$header +=  "HasCustomVoicemailGreeting"
#	$header +=  "HasCustomAwayGreeting"
#	$header +=  "IsValid"
#	$a = [int][char]'a' -1
#	if ($header.GetLength(0) -gt 26) 
#	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
#	else 
#	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
#	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
#	$Header_range.value2 = $header
#	$Header_range.cells.interior.colorindex = 45
#	$Header_range.cells.font.colorindex = 0
#	$Header_range.cells.font.bold = $true
#	$row++
#	$intSheetCount++	
#	$ColumnCount = $header.Count
#	$DataFile = @()
#	$EndCellRow = 1
#
#if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\GetUmMailboxConfiguration") -eq $true)
#{	
#	foreach ($file in (Get-ChildItem -LiteralPath "$RunLocation\output\ExOrg\GetUmMailboxConfiguration" | Where-Object {$_.name -match "~~GetUmMailboxConfiguration"}))
#	{
#		$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\GetUmMailboxConfiguration\" + $file)
#	}
#	$RowCount = $DataFile.Count
#	$ArrayRow = 0
#	$BadArrayValue = @()
#	$DataArray = New-Object 'object[,]' -ArgumentList $RowCount,$ColumnCount
#	Foreach ($DataRow in $DataFile)
#	{
#		$DataField = $DataRow.Split("`t")
#		for ($ArrayColumn = 0 ; $ArrayColumn -lt $ColumnCount ; $ArrayColumn++)
#		{
#			# Excel 2003 limit of 1823 characters
#            if ($DataField[$ArrayColumn].length -lt 1823) 
#                {$DataArray[$ArrayRow,$ArrayColumn] = $DataField[$ArrayColumn]}
#			# Excel 2007 limit of 8203 characters
#            elseif (($Excel_ExOrg.version -ge 12) -and ($DataField[$ArrayColumn].length -lt 8203)) 
#                {$DataArray[$ArrayRow,$ArrayColumn] = $DataField[$ArrayColumn]} 
#			# No known Excel 2010 limit
#            elseif ($Excel_ExOrg.version -ge 14)
#                {$DataArray[$ArrayRow,$ArrayColumn] = $DataField[$ArrayColumn]}
#            else
#            {
#                Write-Host -Object "Number of characters in array member exceeds the version of limitations of this version of Excel" -ForegroundColor Yellow
#				Write-Host -Object "-- Writing value to temp variable" -ForegroundColor Yellow
#                $DataArray[$ArrayRow,$ArrayColumn] = $DataField[$ArrayColumn]
#                $BadArrayValue += "$ArrayRow,$ArrayColumn"
#            }
#		}
#		$ArrayRow++
#	}
#
#    # Replace big values in $DataArray
#    $BadArrayValue_count = $BadArrayValue.count
#    $BadArrayValue_Temp = @()
#    for ($i = 0 ; $i -lt $BadArrayValue_count ; $i++)
#    {
#        $BadArray_Split = $badarrayvalue[$i].Split(",")
#        $BadArrayValue_Temp += $DataArray[$BadArray_Split[0],$BadArray_Split[1]]
#        $DataArray[$BadArray_Split[0],$BadArray_Split[1]] = "**TEMP**"
#		Write-Host -Object "-- Replacing long value with **TEMP**" -ForegroundColor Yellow
#    }
#
#	$EndCellRow = ($RowCount+1)
#	$Data_range = $Worksheet.Range("a2","$EndCellColumn$EndCellRow")
#	$Data_range.Value2 = $DataArray
#
#    # Paste big values back into the spreadsheet
#    for ($i = 0 ; $i -lt $BadArrayValue_count ; $i++)
#    {
#        $BadArray_Split = $badarrayvalue[$i].Split(",")
#        # Adjust for header and $i=0
#        $CellRow = [int]$BadArray_Split[0] + 2
#        # Adjust for $i=0
#        $CellColumn = [int]$BadArray_Split[1] + 1
#           
#        $Range = $Worksheet.cells.item($CellRow,$CellColumn) 
#        $Range.Value2 = $BadArrayValue_Temp[$i]
#		Write-Host -Object "-- Pasting long value back in spreadsheet" -ForegroundColor Yellow
#    }    
#}
#	#EndRegion Get-UmMailboxConfiguration sheet

##Region Get-UmMailboxPin sheet
#Write-Host -Object "---- Starting Get-UmMailboxPin"
#	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
#	$Worksheet.name = "UmMailboxPin"
#	$Worksheet.Tab.ColorIndex = $intColorIndex_Um
#	$row = 1
#	$header = @()
#	$header +=  "UserID"
#	$header +=  "PinExpired"
#	$header +=  "FirstTimeUser"
#	$header +=  "LockedOut"
#	$header +=  "ObjectState"
#	$header +=  "IsValid"
#	$a = [int][char]'a' -1
#	if ($header.GetLength(0) -gt 26) 
#	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
#	else 
#	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
#	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
#	$Header_range.value2 = $header
#	$Header_range.cells.interior.colorindex = 45
#	$Header_range.cells.font.colorindex = 0
#	$Header_range.cells.font.bold = $true
#	$row++
#	$intSheetCount++	
#	$ColumnCount = $header.Count
#	$DataFile = @()
#	$EndCellRow = 1
#
#if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\GetUmMailboxPin") -eq $true)
#{	
#	foreach ($file in (Get-ChildItem -LiteralPath "$RunLocation\output\ExOrg\GetUmMailboxPin" | Where-Object {$_.name -match "~~GetUmMailboxPin"}))
#	{
#		$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\GetUmMailboxPin\" + $file)
#	}
#	$RowCount = $DataFile.Count
#	$ArrayRow = 0
#	$BadArrayValue = @()
#	$DataArray = New-Object 'object[,]' -ArgumentList $RowCount,$ColumnCount
#	Foreach ($DataRow in $DataFile)
#	{
#		$DataField = $DataRow.Split("`t")
#		for ($ArrayColumn = 0 ; $ArrayColumn -lt $ColumnCount ; $ArrayColumn++)
#		{
#			# Excel 2003 limit of 1823 characters
#            if ($DataField[$ArrayColumn].length -lt 1823) 
#                {$DataArray[$ArrayRow,$ArrayColumn] = $DataField[$ArrayColumn]}
#			# Excel 2007 limit of 8203 characters
#            elseif (($Excel_ExOrg.version -ge 12) -and ($DataField[$ArrayColumn].length -lt 8203)) 
#                {$DataArray[$ArrayRow,$ArrayColumn] = $DataField[$ArrayColumn]} 
#			# No known Excel 2010 limit
#            elseif ($Excel_ExOrg.version -ge 14)
#                {$DataArray[$ArrayRow,$ArrayColumn] = $DataField[$ArrayColumn]}
#            else
#            {
#                Write-Host -Object "Number of characters in array member exceeds the version of limitations of this version of Excel" -ForegroundColor Yellow
#				Write-Host -Object "-- Writing value to temp variable" -ForegroundColor Yellow
#                $DataArray[$ArrayRow,$ArrayColumn] = $DataField[$ArrayColumn]
#                $BadArrayValue += "$ArrayRow,$ArrayColumn"
#            }
#		}
#		$ArrayRow++
#	}
#
#    # Replace big values in $DataArray
#    $BadArrayValue_count = $BadArrayValue.count
#    $BadArrayValue_Temp = @()
#    for ($i = 0 ; $i -lt $BadArrayValue_count ; $i++)
#    {
#        $BadArray_Split = $badarrayvalue[$i].Split(",")
#        $BadArrayValue_Temp += $DataArray[$BadArray_Split[0],$BadArray_Split[1]]
#        $DataArray[$BadArray_Split[0],$BadArray_Split[1]] = "**TEMP**"
#		Write-Host -Object "-- Replacing long value with **TEMP**" -ForegroundColor Yellow
#    }
#
#	$EndCellRow = ($RowCount+1)
#	$Data_range = $Worksheet.Range("a2","$EndCellColumn$EndCellRow")
#	$Data_range.Value2 = $DataArray
#
#    # Paste big values back into the spreadsheet
#    for ($i = 0 ; $i -lt $BadArrayValue_count ; $i++)
#    {
#        $BadArray_Split = $badarrayvalue[$i].Split(",")
#        # Adjust for header and $i=0
#        $CellRow = [int]$BadArray_Split[0] + 2
#        # Adjust for $i=0
#        $CellColumn = [int]$BadArray_Split[1] + 1
#           
#        $Range = $Worksheet.cells.item($CellRow,$CellColumn) 
#        $Range.Value2 = $BadArrayValue_Temp[$i]
#		Write-Host -Object "-- Pasting long value back in spreadsheet" -ForegroundColor Yellow
#    }    
#}
#	#EndRegion Get-UmMailboxPin sheet

#Region Get-UmMailboxPolicy sheet
Write-Host -Object "---- Starting Get-UmMailboxPolicy"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "UmMailboxPolicy"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Um
	$row = 1
	$header = @()
	$header +=  "Name"
	$header +=  "Identity"
	$header +=  "MaxGreetingDuration"
	$header +=  "MaxLogonAttempts"
	$header +=  "AllowCommonPatterns"
	$header +=  "PINLifetime"
	$header +=  "PINHistoryCount"
	$header +=  "AllowSMSNotification"
	$header +=  "ProtectUnauthenticatedVoiceMail"
	$header +=  "ProtectAuthenticatedVoiceMail"
	$header +=  "ProtectedVoiceMailText"
	$header +=  "RequireProtectedPlayOnPhone"
	$header +=  "MinPINLength"
	$header +=  "FaxMessageText"
	$header +=  "UMEnabledText"
	$header +=  "ResetPINText"
	$header +=  "SourceForestPolicyNames"
	$header +=  "VoiceMailText"
	$header +=  "UMDialPlan"
	$header +=  "FaxServerURI"
	$header +=  "AllowedInCountryOrRegionGroups"
	$header +=  "AllowedInternationalGroups"
	$header +=  "AllowDialPlanSubscribers"
	$header +=  "AllowExtensions"
	$header +=  "LogonFailuresBeforePINReset"
	$header +=  "AllowMissedCallNotifications"
	$header +=  "AllowFax"
	$header +=  "AllowTUIAccessToCalendar"
	$header +=  "AllowTUIAccessToEmail"
	$header +=  "AllowSubscriberAccess"
	$header +=  "AllowTUIAccessToDirectory"
	$header +=  "AllowTUIAccessToPersonalContacts"
	$header +=  "AllowAutomaticSpeechRecognition"
	$header +=  "AllowPlayOnPhone"
	$header +=  "AllowVoiceMailPreview"
	$header +=  "AllowCallAnsweringRules"
	$header +=  "AllowMessageWaitingIndicator"
	$header +=  "AllowPinlessVoiceMailAccess"
	$header +=  "AllowVoiceResponseToOtherMessageTypes"
	$header +=  "AllowVoiceMailAnalysis"
	$header +=  "AllowVoiceNotification"
	$header +=  "InformCallerOfVoiceMailAnalysis"
	$header +=  "VoiceMailPreviewPartnerAddress"
	$header +=  "VoiceMailPreviewPartnerAssignedID"
	$header +=  "VoiceMailPreviewPartnerMaxMessageDuration"
	$header +=  "VoiceMailPreviewPartnerMaxDeliveryDelay"
	$header +=  "IsDefault"
	$header +=  "IsValid"
	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\ExOrg_GetUmMailboxPolicy.txt") -eq $true)
{	
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\ExOrg_GetUmMailboxPolicy.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Process-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}
	#EndRegion Get-UmMailboxPolicy sheet

#Region Get-UmServer sheet
Write-Host -Object "---- Starting Get-UmServer"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "UmServer"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Um
	$row = 1
	$header = @()
	$header +=  "Name"
	$header +=  "Identity"
	$header +=  "Status"
	$header +=  "MaxCallsAllowed"
	$header +=  "SipTcpListeningPort"
	$header +=  "SipTlsListeningPort"
	$header +=  "Languages"
	$header +=  "DialPlans"
	$header +=  "GrammarGenerationSchedule"
	$header +=  "UmStartupMode"
	$header +=  "IrmLogEnabled"
	$header +=  "IrmLogPath"
	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "$RunLocation\output\ExOrg\ExOrg_GetUmSvr.txt") -eq $true)
{	
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\ExOrg_GetUmSvr.txt")
	# Send the data to the function to process and add to the Excel worksheet
	Process-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}
	#EndRegion Get-UmServer sheet

# Misc
#Region Misc_AdminGroups sheet
Write-Host -Object "---- Starting Misc_AdminGroups"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "Misc_AdminGroups"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Misc
	$row = 1
	$header = @()
	$header +=  "Group Name"
	$header +=  "Member Count"
	$header +=  "Member"
	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "output\ExOrg\ExOrg_Misc_AdminGroups.txt") -eq $true)
{	
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\ExOrg_Misc_AdminGroups.txt") 
	# Send the data to the function to process and add to the Excel worksheet
	Process-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}
	#endRegion Misc_AdminGroups sheet

#Region Misc_Fsmo sheet
Write-Host -Object "---- Starting Misc_Fsmo"
	$Worksheet = $Excel_ExOrg_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "Misc_Fsmo"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Misc
	$row = 1
	$header = @()
	$header +=  "Domain Name"
	$header +=  "PdcRoleOwner (domain)"
	$header +=  "RidRoleOwner (domain)"
	$header +=  "InfrastructureRoleOwner (domain)"
	$header +=  "NamingRoleMaster (forest)"
	$header +=  "SchemaMaster (forest)"

	$a = [int][char]'a' -1
	if ($header.GetLength(0) -gt 26) 
	{$EndCellColumn = [char]([int][math]::Floor($header.GetLength(0)/26) + $a) + [char](($header.GetLength(0)%26) + $a)} 
	else 
	{$EndCellColumn = [char]($header.GetLength(0) + $a)}
	$Header_range = $Worksheet.Range("A1","$EndCellColumn$row")
	$Header_range.value2 = $header
	$Header_range.cells.interior.colorindex = 45
	$Header_range.cells.font.colorindex = 0
	$Header_range.cells.font.bold = $true
	$row++
	$intSheetCount++	
	$ColumnCount = $header.Count
	$DataFile = @()
	$EndCellRow = 1

if ((Test-Path -LiteralPath "output\ExOrg\ExOrg_Misc_Fsmo.txt") -eq $true)
{	
	$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\ExOrg\ExOrg_Misc_Fsmo.txt") 
	# Send the data to the function to process and add to the Excel worksheet
	Process-Datafile $ColumnCount $DataFile $Worksheet $Excel_Version
}
	#endRegion Misc_Fsmo sheet
	
# Autofit columns
Write-Host -Object "---- Starting Autofit"	
$Excel_ExOrgWorksheetCount = $Excel_ExOrg_workbook.worksheets.count
$AutofitSheetCount = 1
while ($AutofitSheetCount -le $Excel_ExOrgWorksheetCount)
{
	$ActiveWorksheet = $Excel_ExOrg_workbook.worksheets.item($AutofitSheetCount)
	$objRange = $ActiveWorksheet.usedrange
	[Void]	$objRange.entirecolumn.autofit()
	$AutofitSheetCount++
}
$Excel_ExOrg_workbook.saveas($ExDC_ExOrg_XLS)
Write-Host -Object "---- Spreadsheet saved"
$Excel_ExOrg.workbooks.close()
Write-Host -Object "---- Workbook closed"
$Excel_ExOrg.quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel_ExOrg)
Remove-Variable -Name Excel_ExOrg
# If the ReleaseComObject doesn't do it..
#spps -n excel

	$EventLog = New-Object -TypeName System.Diagnostics.EventLog -ArgumentList Application
	$EventLog.MachineName = "."
	$EventLog.Source = "ExDC"
	#$EventLog.WriteEntry("Ending Core_Assemble_ExOrg_Excel","Information", 43)	

