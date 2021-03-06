#############################################################################
#                      Core_Assemble_DC_Excel.ps1		 					#
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
$ErrorText = "Core_Assemble_DC_Excel " + "`n" + $server + "`n"
$ErrorText += $_

$ErrorLog = New-Object -TypeName System.Diagnostics.EventLog -ArgumentList Application
$ErrorLog.MachineName = "."
$ErrorLog.Source = "ExDC"
#$ErrorLog.WriteEntry($ErrorText,"Error", 100)
}

set-location -LiteralPath $RunLocation

	$EventLog = New-Object -TypeName System.Diagnostics.EventLog -ArgumentList Application
	$EventLog.MachineName = "."
	$EventLog.Source = "ExDC"
	#$EventLog.WriteEntry("Starting Core_Assemble_DC_Excel","Information", 42)	


Write-Host -Object "---- Starting to create com object for Excel"
$Excel_DC = New-Object -ComObject excel.application
Write-Host -Object "---- Hiding Excel"
$Excel_DC.visible = $false
Write-Host -Object "---- Setting ShowStartupDialog to false" 
$Excel_DC.ShowStartupDialog = $false 
Write-Host -Object "---- Setting DefaultFilePath"
$Excel_DC.DefaultFilePath = $RunLocation + "\output"
Write-Host -Object "---- Setting SheetsInNewWorkbook" 
$Excel_DC.SheetsInNewWorkbook = 15
Write-Host -Object "---- Checking Excel version"
$Excel_Version = $Excel_DC.version
if ($Excel_version -ge 12)
{
    $Excel_DC.DefaultSaveFormat = 51
    $excel_Extension = ".xlsx"
}
else
{
    $Excel_DC.DefaultSaveFormat = 56
    $excel_Extension = ".xls"
}
Write-Host -Object "---- Excel version $excel_version and DefaultSaveFormat $Excel_extension"

# Create new Excel workbook
Write-Host -Object "---- Adding workbook"
$Excel_DC_workbook = $Excel_DC.workbooks.add()
Write-Host -Object "---- Setting output file"
$ExDC_DC_XLS = $RunLocation + "\output\ExDC_DC" + $excel_Extension

Write-Host -Object "---- Setting workbook properties"
$Excel_DC_workbook.author = "Exchange Data Collector v2 (ExDC v2)"
$Excel_DC_workbook.title = "ExDC v2 - Domain Controllers"
$Excel_DC_workbook.comments = "ExDC v2.0.2 Build a1"

$intSheetCount = 1
$intColorIndex_DC = 45

#Region Win32_Bios sheet
Write-Host -Object "---- Starting Win32_Bios"
	$Worksheet = $Excel_DC_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "Win32_Bios"
	$Worksheet.Tab.ColorIndex = $intColorIndex_DC
	$row = 1
	$header = @()
	$header += "Computer"
	$header += "Manufacturer"
	$header += "Name"
	$header += "SerialNumber"
	$header += "Version"
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

if ((Test-Path -LiteralPath "$RunLocation\output\DC_W32_Bios") -eq $true)
{	
	foreach ($file in (Get-ChildItem -LiteralPath "$RunLocation\output\dc_w32_Bios" | Where-Object {$_.name -match "~~DC_W32_bios"}))
	{
		$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\dc_w32_bios\" + $file) 
	}
	$RowCount = $DataFile.Count
	$ArrayRow = 0
	$DataArray = New-Object 'object[,]' -ArgumentList $RowCount,$ColumnCount
	Foreach ($DataRow in $DataFile)
	{
		$DataField = $DataRow.Split("`t")
		for ($ArrayColumn=0;$ArrayColumn -le $ColumnCount-1;$ArrayColumn++)
		{
			$DataArray[$ArrayRow,$ArrayColumn] = $DataField[$ArrayColumn]
		}
		$ArrayRow++
	}

	$EndCellRow = ($RowCount+1)
	$Data_range = $Worksheet.Range("a2","$EndCellColumn$EndCellRow")
	$Data_range.Value2 = $DataArray	
}
	#endRegion	
	
#Region Win32_ComputerSystem sheet
Write-Host -Object "---- Starting Win32_ComputerSystem"
	$Worksheet = $Excel_DC_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "Win32_ComputerSystem"
	$Worksheet.Tab.ColorIndex = $intColorIndex_DC
	$row = 1
	$header = @()
	$header += "Computer"
	$header += "Manufacturer"
	$header += "Model"
	$header += "NumberOfLogicalProcessors"
	$header += "NumberOfProcessors"
	$header += "TotalPhysicalMemory"
	$header += "AutomaticManagedPagefile"
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

if ((Test-Path -LiteralPath "$RunLocation\output\DC_W32_CS") -eq $true)
{	
	foreach ($file in (Get-ChildItem -LiteralPath "$RunLocation\output\dc_w32_cs" | Where-Object {$_.name -match "~~DC_W32_CS"}))
	{
		$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\dc_w32_cs\" + $file) 
	}
	$RowCount = $DataFile.Count
	$ArrayRow = 0
	$DataArray = New-Object 'object[,]' -ArgumentList $RowCount,$ColumnCount
	Foreach ($DataRow in $DataFile)
	{
		$DataField = $DataRow.Split("`t")
		for ($ArrayColumn=0;$ArrayColumn -le $ColumnCount-1;$ArrayColumn++)
		{
			$DataArray[$ArrayRow,$ArrayColumn] = $DataField[$ArrayColumn]
		}
		$ArrayRow++
	}

	$EndCellRow = ($RowCount+1)
	$Data_range = $Worksheet.Range("a2","$EndCellColumn$EndCellRow")
	$Data_range.Value2 = $DataArray	
}
	#endRegion	

#Region Win32_LogicalDisk sheet
Write-Host -Object "---- Starting Win32_LogicalDisk"
	$Worksheet = $Excel_DC_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "Win32_LogicalDisk"
	$Worksheet.Tab.ColorIndex = $intColorIndex_DC
	$row = 1
	$header = @()
	$header += "Computer"
	$header += "DriveType"
	$header += "Caption"
	$header += "Compressed"
	$header += "Description"
	$header += "FileSystem"
	$header += "Name"
	$header += "VolumeName"	
	$header += "Size"
	$header += "FreeSpace"
	$header += "Size (GB)"
	$header += "FreeSpace (GB)"
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

if ((Test-Path -LiteralPath "$RunLocation\output\DC_W32_LD") -eq $true)
{
	foreach ($file in (Get-ChildItem -LiteralPath "$RunLocation\output\dc_w32_ld" | Where-Object {$_.name -match "~~DC_W32_LD"}))
	{
		$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\dc_w32_ld\" + $file) 
	}
	$RowCount = $DataFile.Count
	$ArrayRow = 0
	$DataArray = New-Object 'object[,]' -ArgumentList $RowCount,$ColumnCount
	Foreach ($DataRow in $DataFile)
	{
		$DataField = $DataRow.Split("`t")
		for ($ArrayColumn=0;$ArrayColumn -le $ColumnCount-1;$ArrayColumn++)
		{
			$DataArray[$ArrayRow,$ArrayColumn] = $DataField[$ArrayColumn]
		}
		$ArrayRow++
	}

	$EndCellRow = ($RowCount+1)
	$Data_range = $Worksheet.Range("a2","$EndCellColumn$EndCellRow")
	$Data_range.Value2 = $DataArray	
}
	#EndRegion

#Region Win32_NetworkAdapter sheet
Write-Host -Object "---- Starting Win32_NetworkAdapter"
	$Worksheet = $Excel_DC_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "Win32_NetworkAdapter"
	$Worksheet.Tab.ColorIndex = $intColorIndex_DC
	$row = 1
	$header = @()
	$header += "Computer"
	$header += "DeviceID"
	$header += "Caption"
	$header += "Description"
	$header += "Manufacturer"
	$header += "Name"
	$header += "NetConnectionStatus"
	$header += "ProductName"
	$header += "Speed"
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

if ((Test-Path -LiteralPath "$RunLocation\output\DC_W32_NA") -eq $true)
{
	foreach ($file in (Get-ChildItem -LiteralPath "$RunLocation\output\dc_w32_na" | Where-Object {$_.name -match "~~DC_W32_NA"}))
	{
		$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\dc_w32_na\" + $file) 
	}
	$RowCount = $DataFile.Count
	$ArrayRow = 0
	$DataArray = New-Object 'object[,]' -ArgumentList $RowCount,$ColumnCount
	Foreach ($DataRow in $DataFile)
	{
		$DataField = $DataRow.Split("`t")
		for ($ArrayColumn=0;$ArrayColumn -le $ColumnCount-1;$ArrayColumn++)
		{
			$DataArray[$ArrayRow,$ArrayColumn] = $DataField[$ArrayColumn]
		}
		$ArrayRow++
	}

	$EndCellRow = ($RowCount+1)
	$Data_range = $Worksheet.Range("a2","$EndCellColumn$EndCellRow")
	$Data_range.Value2 = $DataArray	
}
	#EndRegion

#Region Win32_NetworkAdapterConfiguration sheet
Write-Host -Object "---- Starting Win32_NetworkAdapterConfiguration"	
	$Worksheet = $Excel_DC_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "Win32_NetworkAdapterConfig"
	$Worksheet.Tab.ColorIndex = $intColorIndex_DC
	$row = 1
	$header = @()
	$header += "Computer"
	$header += "Index"
	$header += "Caption"
	$header += "Description"
	$header += "DefaultIPGateway"
	$header += "DNSHostName"
	$header += "DNSServerSearchOrder"
	$header += "IPAddress"
	$header += "IPSubnet"
	$header += "DHCPEnabled"
	$header += "TCPWindowSize"
	$header += "WINSPrimaryServer"
	$header += "WINSSecondaryServer"
	$header += "DomainDNSRegistrationEnabled"
	$header += "FullDNSRegistrationEnabled"
	$header += "DNSDomainSuffixSearchOrder"
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

if ((Test-Path -LiteralPath "$RunLocation\output\DC_W32_NAC") -eq $true)
{
	foreach ($file in (Get-ChildItem -LiteralPath "$RunLocation\output\dc_w32_nac" | Where-Object {$_.name -match "~~DC_W32_NAC"}))
	{
		$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\dc_w32_nac\" + $file) 
	}
	$RowCount = $DataFile.Count
	$ArrayRow = 0
	$DataArray = New-Object 'object[,]' -ArgumentList $RowCount,$ColumnCount
	Foreach ($DataRow in $DataFile)
	{
		$DataField = $DataRow.Split("`t")
		for ($ArrayColumn=0;$ArrayColumn -le $ColumnCount-1;$ArrayColumn++)
		{
			$DataArray[$ArrayRow,$ArrayColumn] = $DataField[$ArrayColumn]
		}
		$ArrayRow++
	}

	$EndCellRow = ($RowCount+1)
	$Data_range = $Worksheet.Range("a2","$EndCellColumn$EndCellRow")
	$Data_range.Value2 = $DataArray	
}
	#EndRegion

#Region Win32_OperatingSystem sheet
Write-Host -Object "---- Starting Win32_OperatingSystem"
	$Worksheet = $Excel_DC_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "Win32_OperatingSystem"
	$Worksheet.Tab.ColorIndex = $intColorIndex_DC
	$row = 1
	$header = @()
	$header += "Computer"
	$header += "CSName"
	$header += "Version"
	$header += "ServicePackMajorVersion"
	$header += "SystemDrive"
	$header += "OSArchitecture"
	$header += "MaxProcessMemorySize"
	$header += "Caption"
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

if ((Test-Path -LiteralPath "$RunLocation\output\DC_W32_OS") -eq $true)
{
	foreach ($file in (Get-ChildItem -LiteralPath "$RunLocation\output\dc_w32_os" | Where-Object {$_.name -match "~~DC_W32_OS"}))
	{
		$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\dc_w32_os\" + $file) 
	}
	$RowCount = $DataFile.Count
	$ArrayRow = 0
	$DataArray = New-Object 'object[,]' -ArgumentList $RowCount,$ColumnCount
	Foreach ($DataRow in $DataFile)
	{
		$DataField = $DataRow.Split("`t")
		for ($ArrayColumn=0;$ArrayColumn -le $ColumnCount-1;$ArrayColumn++)
		{
			$DataArray[$ArrayRow,$ArrayColumn] = $DataField[$ArrayColumn]
		}
		$ArrayRow++
	}

	$EndCellRow = ($RowCount+1)
	$Data_range = $Worksheet.Range("a2","$EndCellColumn$EndCellRow")
	$Data_range.Value2 = $DataArray	
}
	#EndRegion

#Region Win32_PageFileUsage sheet
Write-Host -Object "---- Starting Win32_PageFileUsage"
	$Worksheet = $Excel_DC_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "Win32_PageFileUsage"
	$Worksheet.Tab.ColorIndex = $intColorIndex_DC
	$row = 1
	$header = @()
	$header += "Computer"
	$header += "Name"
	$header += "AllocatedBaseSize"
	$header += "Caption"
	$header += "CurrentUsage"
	$header += "Description"
	$header += "PeakUsage"
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

if ((Test-Path -LiteralPath "$RunLocation\output\DC_W32_PFU") -eq $true)
{
	foreach ($file in (Get-ChildItem -LiteralPath "$RunLocation\output\dc_w32_pfu" | Where-Object {$_.name -match "~~DC_W32_PFU"}))
	{
		$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\dc_w32_pfu\" + $file) 
	}
	$RowCount = $DataFile.Count
	$ArrayRow = 0
	$DataArray = New-Object 'object[,]' -ArgumentList $RowCount,$ColumnCount
	Foreach ($DataRow in $DataFile)
	{
		$DataField = $DataRow.Split("`t")
		for ($ArrayColumn=0;$ArrayColumn -le $ColumnCount-1;$ArrayColumn++)
		{
			$DataArray[$ArrayRow,$ArrayColumn] = $DataField[$ArrayColumn]
		}
		$ArrayRow++
	}

	$EndCellRow = ($RowCount+1)
	$Data_range = $Worksheet.Range("a2","$EndCellColumn$EndCellRow")
	$Data_range.Value2 = $DataArray	
}
	#EndRegion

#Region Win32_PhysicalMemory sheet
Write-Host -Object "---- Starting Win32_PhysicalMemory"
	$Worksheet = $Excel_DC_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "Win32_PhysicalMemory"
	$Worksheet.Tab.ColorIndex = $intColorIndex_DC
	$row = 1
	$header = @()
	$header += "Computer"
	$header += "Tag"
	$header += "BankLabel"
	$header += "Capacity"
	$header += "DeviceLocator"
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

if ((Test-Path -LiteralPath "$RunLocation\output\DC_W32_PM") -eq $true)
{
	foreach ($file in (Get-ChildItem -LiteralPath "$RunLocation\output\dc_w32_pm" | Where-Object {$_.name -match "~~DC_W32_PM"}))
	{
		$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\dc_w32_pm\" + $file) 
	}
	$RowCount = $DataFile.Count
	$ArrayRow = 0
	$DataArray = New-Object 'object[,]' -ArgumentList $RowCount,$ColumnCount
	Foreach ($DataRow in $DataFile)
	{
		$DataField = $DataRow.Split("`t")
		for ($ArrayColumn=0;$ArrayColumn -le $ColumnCount-1;$ArrayColumn++)
		{
			$DataArray[$ArrayRow,$ArrayColumn] = $DataField[$ArrayColumn]
		}
		$ArrayRow++
	}

	$EndCellRow = ($RowCount+1)
	$Data_range = $Worksheet.Range("a2","$EndCellColumn$EndCellRow")
	$Data_range.Value2 = $DataArray	
}
	#EndRegion

#Region Win32_Processor sheet
Write-Host -Object "---- Starting Win32_Processor"
	$Worksheet = $Excel_DC_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "Win32_Processor"
	$Worksheet.Tab.ColorIndex = $intColorIndex_DC
	$row = 1
	$header = @()
	$header +=  "Computer"
	$header +=  "DeviceID"
	$header +=  "CurrentClockSpeed"
	$header +=  "Description"
	$header +=  "Manufacturer"
	$header +=  "AddressWidth"
	$header +=  "Architecture"
	$header +=  "NumberOfCores"
	$header +=  "NumberOfLogicalProcessors"
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

if ((Test-Path -LiteralPath "$RunLocation\output\DC_W32_Proc") -eq $true)
{
	foreach ($file in (Get-ChildItem -LiteralPath "$RunLocation\output\dc_w32_proc" | Where-Object {$_.name -match "~~DC_W32_Proc"}))
	{
		$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\dc_w32_proc\" + $file) 
	}
	$RowCount = $DataFile.Count
	$ArrayRow = 0
	$DataArray = New-Object 'object[,]' -ArgumentList $RowCount,$ColumnCount
	Foreach ($DataRow in $DataFile)
	{
		$DataField = $DataRow.Split("`t")
		for ($ArrayColumn=0;$ArrayColumn -le $ColumnCount-1;$ArrayColumn++)
		{
			$DataArray[$ArrayRow,$ArrayColumn] = $DataField[$ArrayColumn]
		}
		$ArrayRow++
	}

	$EndCellRow = ($RowCount+1)
	$Data_range = $Worksheet.Range("a2","$EndCellColumn$EndCellRow")
	$Data_range.Value2 = $DataArray	
}
	#EndRegion

#Region MicrosoftDNS_Zone sheet
Write-Host -Object "---- Starting MicrosoftDNS_Zone"	
	$Worksheet = $Excel_DC_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "MicrosoftDNS_Zone"
	$Worksheet.Tab.ColorIndex = $intColorIndex_DC
	$row = 1
	$header = @()
	$header += "Computer"
	$header += "Name"
	$header += "DSIntegrated"
	$header += "ZoneType"
	$header += "Reverse"
	$header += "AllowUpdate"
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

if ((Test-Path -LiteralPath "$RunLocation\output\DC_DNS") -eq $true)
{
	foreach ($file in (Get-ChildItem -LiteralPath "$RunLocation\output\dc_dns" | Where-Object {$_.name -match "~~DC_DNS"}))
	{
		$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\dc_dns\" + $file) 
	}
	$RowCount = $DataFile.Count
	$ArrayRow = 0
	$DataArray = New-Object 'object[,]' -ArgumentList $RowCount,$ColumnCount
	Foreach ($DataRow in $DataFile)
	{
		$DataField = $DataRow.Split("`t")
		for ($ArrayColumn=0;$ArrayColumn -le $ColumnCount-1;$ArrayColumn++)
		{
			$DataArray[$ArrayRow,$ArrayColumn] = $DataField[$ArrayColumn]
		}
		$ArrayRow++
	}

	$EndCellRow = ($RowCount+1)
	$Data_range = $Worksheet.Range("a2","$EndCellColumn$EndCellRow")
	$Data_range.Value2 = $DataArray	
}
	#EndRegion

#Region MSAD_ReplNeighbor sheet
Write-Host -Object "---- Starting MSAD_ReplNeighbor"	
	$Worksheet = $Excel_DC_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "MSAD_ReplNeighbor"
	$Worksheet.Tab.ColorIndex = $intColorIndex_DC
	$row = 1
	$header = @()
	$header += "Computer"
	$header += "SourceDsaCN"
	$header += "SourceDsaSite"
	$header += "Domain"
	$header += "CompressChanges"
	$header += "DisableScheduledSync"
	$header += "DoScheduledSyncs"
	$header += "IgnoreChangeNotifications"
	$header += "IsDeletedSourceDsa"
	$header += "LastSyncResult"
	$header += "ModifiedNumConsecutiveSyncFailures"
	$header += "NeverSynced"
	$header += "NoChangeNotifications"
	$header += "NumConsecutiveSyncFailures"
	$header += "SyncOnStartup"
	$header += "TimeOfLastSyncAttempt"
	$header += "TimeofLastSyncSuccess"
	$header += "TwoWaySync"
	$header += "Writeable"
	
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

if ((Test-Path -LiteralPath "$RunLocation\output\DC_MSAD_ReplNeighbor") -eq $true)
{
	foreach ($file in (Get-ChildItem -LiteralPath "$RunLocation\output\DC_MSAD_ReplNeighbor" | Where-Object {$_.name -match "~~DC_MSAD_ReplNeighbor"}))
	{
		$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\DC_MSAD_ReplNeighbor\" + $file) 
	}
	$RowCount = $DataFile.Count
	$ArrayRow = 0
	$DataArray = New-Object 'object[,]' -ArgumentList $RowCount,$ColumnCount
	Foreach ($DataRow in $DataFile)
	{
		$DataField = $DataRow.Split("`t")
		for ($ArrayColumn=0;$ArrayColumn -le $ColumnCount-1;$ArrayColumn++)
		{
			$DataArray[$ArrayRow,$ArrayColumn] = $DataField[$ArrayColumn]
		}
		$ArrayRow++
	}

	$EndCellRow = ($RowCount+1)
	$Data_range = $Worksheet.Range("a2","$EndCellColumn$EndCellRow")
	$Data_range.Value2 = $DataArray	
}
	#EndRegion

#Region MSAD_DomainController sheet
Write-Host -Object "---- Starting MSAD_DomainController"	
	$Worksheet = $Excel_DC_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "MSAD_DomainController"
	$Worksheet.Tab.ColorIndex = $intColorIndex_DC
	$row = 1
	$header = @()
	$header += "Computer"
	$header += "IsAdvertisingToLocator"
	$header += "IsGC"
	$header += "IsNextRIDPoolAvailable"
	$header += "IsRegisteredInDNS"
	$header += "IsSysVolReady"
	$header += "PercentOfRIDsLeft"
	$header += "SiteName"
	
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

if ((Test-Path -LiteralPath "$RunLocation\output\DC_MSAD_DomainController") -eq $true)
{
	foreach ($file in (Get-ChildItem -LiteralPath "$RunLocation\output\DC_MSAD_DomainController" | Where-Object {$_.name -match "~~DC_MSAD_DomainController"}))
	{
		$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\DC_MSAD_DomainController\" + $file) 
	}
	$RowCount = $DataFile.Count
	$ArrayRow = 0
	$DataArray = New-Object 'object[,]' -ArgumentList $RowCount,$ColumnCount
	Foreach ($DataRow in $DataFile)
	{
		$DataField = $DataRow.Split("`t")
		for ($ArrayColumn=0;$ArrayColumn -le $ColumnCount-1;$ArrayColumn++)
		{
			$DataArray[$ArrayRow,$ArrayColumn] = $DataField[$ArrayColumn]
		}
		$ArrayRow++
	}

	$EndCellRow = ($RowCount+1)
	$Data_range = $Worksheet.Range("a2","$EndCellColumn$EndCellRow")
	$Data_range.Value2 = $DataArray	
}
	#EndRegion

#Region Registry - AD sheet
Write-Host -Object "---- Starting Registry - AD"
	$Worksheet = $Excel_DC_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "Registry - AD"
	$Worksheet.Tab.ColorIndex = $intColorIndex_DC
	$row = 1
	$header = @()
	$header += "Computer"
	$header += "ForestName"
	$header += "DomainName"
	$header += "SiteName"
	$header += "IsGC"
	$header += "Roles"
	$header += "ForestMode"
	$header += "DomainMode"
	$header += "NTDS\TCP/IP Port"
	$header += "NTFRS\RPC TCP/IP Port Assignment"
	$header += "Netlogon\DCTcpipPort"
	$header += "Netlogon\MaxConcurrentApi"
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

if ((Test-Path -LiteralPath "$RunLocation\output\DC_Reg_AD") -eq $true)
{
	foreach ($file in (Get-ChildItem -LiteralPath "$RunLocation\output\dc_reg_ad" | Where-Object {$_.name -match "~~DC_Reg_AD"}))
	{
		$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\dc_reg_AD\" + $file) 
	}
	$RowCount = $DataFile.Count
	$ArrayRow = 0
	$DataArray = New-Object 'object[,]' -ArgumentList $RowCount,$ColumnCount
	Foreach ($DataRow in $DataFile)
	{
		$DataField = $DataRow.Split("`t")
		for ($ArrayColumn=0;$ArrayColumn -le $ColumnCount-1;$ArrayColumn++)
		{
			$DataArray[$ArrayRow,$ArrayColumn] = $DataField[$ArrayColumn]
		}
		$ArrayRow++
	}

	$EndCellRow = ($RowCount+1)
	$Data_range = $Worksheet.Range("a2","$EndCellColumn$EndCellRow")
	$Data_range.Value2 = $DataArray	
}
	#EndRegion Registry - AD sheet

#Region Registry - OS sheet
Write-Host -Object "---- Starting Registry - OS"
	$Worksheet = $Excel_DC_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "Registry - OS"
	$Worksheet.Tab.ColorIndex = $intColorIndex_DC
	$row = 1
	$header = @()
	$header += "Computer"
	$header += "W32 Time - Type"
	$header += "W32 Time - NTP Server"
	$header += "W32 Time - MaxNegPhaseCorrection"
	$header += "W32 Time - MaxPosPhaseCorrection"
	$header += "Tcpip - RSS"
	$header += "Tcpip - TCPA"
	$header += "Tcpip - TCPChimney"
	$header += "IPv6 - DisabledComponents"
	$header += "Environment - TEMP"
	$header += "Environment - TMP"
	$header += "RestrictAnonymous"
	$header += "RestrictAnonymousSam"
	$header += "NtlmMinClientSec"
	$header += "NtlmMinServerSec"
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

if ((Test-Path -LiteralPath "$RunLocation\output\DC_Reg_OS") -eq $true)
{
	foreach ($file in (Get-ChildItem -LiteralPath "$RunLocation\output\dc_reg_os" | Where-Object {$_.name -match "~~DC_Reg_OS"}))
	{
		$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\dc_reg_OS\" + $file) 
	}
	$RowCount = $DataFile.Count
	$ArrayRow = 0
	$DataArray = New-Object 'object[,]' -ArgumentList $RowCount,$ColumnCount
	Foreach ($DataRow in $DataFile)
	{
		$DataField = $DataRow.Split("`t")
		for ($ArrayColumn=0;$ArrayColumn -le $ColumnCount-1;$ArrayColumn++)
		{
			$DataArray[$ArrayRow,$ArrayColumn] = $DataField[$ArrayColumn]
		}
		$ArrayRow++
	}

	$EndCellRow = ($RowCount+1)
	$Data_range = $Worksheet.Range("a2","$EndCellColumn$EndCellRow")
	$Data_range.Value2 = $DataArray	
}
	#EndRegion Registry - OS sheet

#Region Registry - Software sheet
Write-Host -Object "---- Starting Registry - Software"
	$Worksheet = $Excel_DC_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "Registry - Software"
	$Worksheet.Tab.ColorIndex = $intColorIndex_DC
	$row = 1
	$header = @()
	$header += "Computer"
	$header += "Publisher"
	$header += "DisplayName"
	$header += "DisplayVersion"
	$header += "InstallDate"
	$header += "InstallLocation"
	$header += "InstallSource"
	$header += "Version"
	$header += "VersionMajor"
	$header += "VersionMinor"
	
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

if ((Test-Path -LiteralPath "$RunLocation\output\DC_Reg_Software") -eq $true)
{
	foreach ($file in (Get-ChildItem -LiteralPath "$RunLocation\output\dc_reg_software" | Where-Object {$_.name -match "~~DC_Reg_Software"}))
	{
		$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\dc_reg_Software\" + $file) 
	}
	$RowCount = $DataFile.Count
	$ArrayRow = 0
	$DataArray = New-Object 'object[,]' -ArgumentList $RowCount,$ColumnCount
	Foreach ($DataRow in $DataFile)
	{
		$DataField = $DataRow.Split("`t")
		for ($ArrayColumn=0;$ArrayColumn -le $ColumnCount-1;$ArrayColumn++)
		{
			$DataArray[$ArrayRow,$ArrayColumn] = $DataField[$ArrayColumn]
		}
		$ArrayRow++
	}

	$EndCellRow = ($RowCount+1)
	$Data_range = $Worksheet.Range("a2","$EndCellColumn$EndCellRow")
	$Data_range.Value2 = $DataArray	
}
	#EndRegion Registry - Software sheet

# Autofit columns
Write-Host -Object "---- Starting Autofit"	
$excel_DCWorksheetCount = $excel_dc_workbook.worksheets.count
$AutofitSheetCount = 1
while ($AutofitSheetCount -le $excel_DCWorksheetCount)
{
	$ActiveWorksheet = $excel_dc_workbook.worksheets.item($AutofitSheetCount)
	$objRange = $ActiveWorksheet.usedrange
	[Void]	$objRange.entirecolumn.autofit()
	$AutofitSheetCount++
}
$excel_dc_workbook.saveas($ExDC_DC_XLS)
Write-Host -Object "---- Spreadsheet saved"
$excel_dc.workbooks.close()
Write-Host -Object "---- Workbook closed"
$Excel_dc.quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel_dc)
Remove-Variable -Name Excel_dc
# If the ReleaseComObject doesn't do it..
#spps -n excel

	$EventLog = New-Object -TypeName System.Diagnostics.EventLog -ArgumentList Application
	$EventLog.MachineName = "."
	$EventLog.Source = "ExDC"
	#$EventLog.WriteEntry("Ending Core_Assemble_DC_Excel","Information", 43)	

