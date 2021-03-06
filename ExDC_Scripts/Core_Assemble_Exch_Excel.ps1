#############################################################################
#                     Core_Assemble_Exch_Excel.ps1		 					#
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
$ErrorText = "Core_Assemble_Exch_Excel " + "`n" + $server + "`n"
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
	#$EventLog.WriteEntry("Starting Core_Assemble_Exch_Excel","Information", 42)	

Write-Host -Object "---- Starting to create com object for Excel"
$Excel_Exch = New-Object -ComObject excel.application
Write-Host -Object "---- Hiding Excel"
$Excel_Exch.visible = $false
Write-Host -Object "---- Setting ShowStartupDialog to false" 
$Excel_Exch.ShowStartupDialog = $false 
Write-Host -Object "---- Setting DefaultFilePath"
$Excel_Exch.DefaultFilePath = $RunLocation + "\output"
Write-Host -Object "---- Setting SheetsInNewWorkbook" 
$Excel_Exch.SheetsInNewWorkbook = 16
Write-Host -Object "---- Checking Excel version"
$Excel_Version = $Excel_Exch.version
if ($Excel_version -ge 12)
{
	$Excel_Exch.DefaultSaveFormat = 51
	$excel_Extension = ".xlsx"
}
else
{
	$Excel_Exch.DefaultSaveFormat = 56
	$excel_Extension = ".xls"
}
Write-Host -Object "---- Excel version $Excel_version and DefaultSaveFormat $Excel_extension"

# Create new Excel workbook
Write-Host -Object "---- Adding workbook"
$Excel_Exch_workbook = $Excel_Exch.workbooks.add()
Write-Host -Object "---- Setting output file"
$ExDC_Exch_XLS = $RunLocation + "\output\ExDC_Exch" + $excel_Extension

Write-Host -Object "---- Setting workbook properties"
$Excel_Exch_workbook.author = "Exchange Data Collector v2 (ExDC v2)"
$Excel_Exch_workbook.title = "ExDC v2 - Exchange Servers"
$Excel_Exch_workbook.comments = "ExDC v2.0.2 Build A1"

$intSheetCount = 1
$intColorIndex_Exch = 45

#Region Win32_Bios sheet
Write-Host -Object "---- Starting Win32_Bios"
	$Worksheet = $Excel_Exch_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "Win32_Bios"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Exch
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

if ((Test-Path -LiteralPath "$RunLocation\output\Exch_W32_bios") -eq $true)
{	
	foreach ($file in (Get-ChildItem -LiteralPath "$RunLocation\output\Exch_w32_bios" | Where-Object {$_.name -match "~~Exch_W32_bios"}))
	{
		$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Exch_w32_bios\" + $file) 
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

#Region Win32_ComputerSystem sheet
Write-Host -Object "---- Starting Win32_ComputerSystem"
	$Worksheet = $Excel_Exch_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "Win32_ComputerSystem"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Exch
	$row = 1
	$header = @()
	$header +=  "Computer"
	$header +=  "Manufacturer"
	$header +=  "Model"
	$header +=  "NumberOfLogicalProcessors"
	$header +=  "NumberOfProcessors"
	$header +=  "TotalPhysicalMemory"
	$header +=  "AutomaticManagedPagefile"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Exch_W32_CS") -eq $true)
{	
	foreach ($file in (Get-ChildItem -LiteralPath "$RunLocation\output\Exch_w32_cs" | Where-Object {$_.name -match "~~Exch_W32_CS"}))
	{
		$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Exch_w32_cs\" + $file) 
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

#Region Win32_LogicalDisk sheet
Write-Host -Object "---- Starting Win32_LogicalDisk"
	$Worksheet = $Excel_Exch_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "Win32_LogicalDisk"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Exch
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

if ((Test-Path -LiteralPath "$RunLocation\output\Exch_W32_LD") -eq $true)
{
	foreach ($file in (Get-ChildItem -LiteralPath "$RunLocation\output\Exch_w32_ld" | Where-Object {$_.name -match "~~Exch_W32_LD"}))
	{
		$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Exch_w32_ld\" + $file) 
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
	$Worksheet = $Excel_Exch_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "Win32_NetworkAdapter"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Exch
	$row = 1
	$header = @()
	$header +=  "Computer"
	$header +=  "DeviceID"
	$header +=  "Caption"
	$header +=  "Description"
	$header +=  "Manufacturer"
	$header +=  "Name"
	$header +=  "NetConnectionStatus"
	$header +=  "ProductName"
	$header +=  "Speed"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Exch_W32_NA") -eq $true)
{
	foreach ($file in (Get-ChildItem -LiteralPath "$RunLocation\output\Exch_w32_na" | Where-Object {$_.name -match "~~Exch_W32_NA"}))
	{
		$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Exch_w32_na\" + $file) 
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
	$Worksheet = $Excel_Exch_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "Win32_NetworkAdapterConfig"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Exch
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

if ((Test-Path -LiteralPath "$RunLocation\output\Exch_W32_NAC") -eq $true)
{
	foreach ($file in (Get-ChildItem -LiteralPath "$RunLocation\output\Exch_w32_nac" | Where-Object {$_.name -match "~~Exch_W32_NAC"}))
	{
		$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Exch_w32_nac\" + $file) 
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
	$Worksheet = $Excel_Exch_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "Win32_OperatingSystem"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Exch
	$row = 1
	$header = @()
	$header +=  "Computer"
	$header +=  "CSName"
	$header +=  "Version"
	$header +=  "ServicePackMajorVersion"
	$header +=  "SystemDrive"
	$header +=  "OSArchitecture"
	$header +=  "MaxProcessMemorySize"
	$header +=  "Caption"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Exch_W32_OS") -eq $true)
{
	foreach ($file in (Get-ChildItem -LiteralPath "$RunLocation\output\Exch_w32_os" | Where-Object {$_.name -match "~~Exch_W32_OS"}))
	{
		$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Exch_w32_os\" + $file) 
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
	$Worksheet = $Excel_Exch_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "Win32_PageFileUsage"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Exch
	$row = 1
	$header = @()
	$header +=  "Computer"
	$header +=  "Name"
	$header +=  "AllocatedBaseSize"
	$header +=  "Caption"
	$header +=  "CurrentUsage"
	$header +=  "Description"
	$header +=  "PeakUsage"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Exch_W32_PFU") -eq $true)
{
	foreach ($file in (Get-ChildItem -LiteralPath "$RunLocation\output\Exch_w32_pfu" | Where-Object {$_.name -match "~~Exch_W32_PFU"}))
	{
		$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Exch_w32_pfu\" + $file) 
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
	$Worksheet = $Excel_Exch_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "Win32_PhysicalMemory"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Exch
	$row = 1
	$header = @()
	$header +=  "Computer"
	$header +=  "Tag"
	$header +=  "BankLabel"
	$header +=  "Capacity"
	$header +=  "DeviceLocator"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Exch_W32_PM") -eq $true)
{
	foreach ($file in (Get-ChildItem -LiteralPath "$RunLocation\output\Exch_w32_pm" | Where-Object {$_.name -match "~~Exch_W32_PM"}))
	{
		$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Exch_w32_pm\" + $file) 
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
	$Worksheet = $Excel_Exch_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "Win32_Processor"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Exch
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

if ((Test-Path -LiteralPath "$RunLocation\output\Exch_W32_Proc") -eq $true)
{
	foreach ($file in (Get-ChildItem -LiteralPath "$RunLocation\output\Exch_w32_proc" | Where-Object {$_.name -match "~~Exch_W32_Proc"}))
	{
		$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Exch_w32_proc\" + $file) 
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

#Region Registry - Exchange sheet
Write-Host -Object "---- Starting Registry - Exchange"
	$Worksheet = $Excel_Exch_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "Registry - Exchange"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Exch
	$row = 1
	$header = @()
	$header +=  "Server"
	$header +=  "MinUserDC"
	$header +=  "MSExchangeIS\VirusScan\Enabled"
	$header +=  "MSExchangeIS\VirusScan\Vendor"
	$header +=  "MSExchangeIS\VirusScan\Version"
	$header +=  "MSExchangeSA\Parameters\TCP/IP Port"
	$header +=  "MSExchangeSA\Parameters\TCP/IP NSPI Port"
	$header +=  "MSExchangeSA\Parameters\NSPI Target Server"
	$header +=  "MSExchangeSA\Parameters\RFR Target Server"
	$header +=  "MSExchangeIS\ParametersSystem\TCP/IP Port"
	$header +=  "MSExchangeIS\ParametersSystem\Maximum Polling Frequency"
	$header +=  "MSExchangeIS\ParametersSystem\Online Maintenance Checksum"
	$header +=  "MSExchangeRPC\ParametersSystem\TCP/IP Port"
	$header +=  "MSExchangeRPC\ParametersSystem\EnablePushNotifications"
	$header +=  "MSExchangeAB\Parameters\RpcTcpPort"
	$header +=  "ReSvc\Parameters\SuppressStateChanges"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Exch_Reg_Ex") -eq $true)
{
	foreach ($file in (Get-ChildItem -LiteralPath "$RunLocation\output\Exch_Reg_Ex" | Where-Object {$_.name -match "~~Exch_Reg_Ex"}))
	{
		$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Exch_Reg_Ex\" + $file) 
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
	#EndRegion Registry - Exchange sheet

#Region Registry - OS sheet
Write-Host -Object "---- Starting Registry - OS"
	$Worksheet = $Excel_Exch_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "Registry - OS"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Exch
	$row = 1
	$header = @()
	$header +=  "Computer"
	$header +=  "W32 Time - Type"
	$header +=  "W32 Time - NTP Server"
	$header +=  "W32 Time - MaxNegPhaseCorrection"
	$header +=  "W32 Time - MaxPosPhaseCorrection"
	$header +=  "Tcpip - RSS"
	$header +=  "Tcip - TCPA"
	$header +=  "Tcpip - TCPChimney"
	$header +=  "IPv6 - DisabledComponents"
	$header +=  "Environment - TEMP"
	$header +=  "Environment - TMP"
	$header +=  "Netlogon - Dynamic Site Name"
	$header +=  "Netlogon - Site Name"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Exch_Reg_OS") -eq $true)
{
	foreach ($file in (Get-ChildItem -LiteralPath "$RunLocation\output\Exch_reg_OS" | Where-Object {$_.name -match "~~Exch_Reg_OS"}))
	{
		$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Exch_reg_OS\" + $file) 
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
	$Worksheet = $Excel_Exch_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "Registry - Software"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Exch
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

if ((Test-Path -LiteralPath "$RunLocation\output\Exch_Reg_Software") -eq $true)
{
	foreach ($file in (Get-ChildItem -LiteralPath "$RunLocation\output\Exch_reg_Software" | Where-Object {$_.name -match "~~Exch_Reg_Software"}))
	{
		$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Exch_reg_Software\" + $file) 
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

#Region ClusterNetwork sheet
Write-Host -Object "---- Starting MSCluster_Network"
	$Worksheet = $Excel_Exch_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "MSCluster_Network"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Exch
	$row = 1
	$header = @()
	$header +=  "Cluster Name"
	$header +=  "Name"
	$header +=  "Role"
	$header +=  "State"
	$header +=  "Status"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Cluster_MSCluster_Network") -eq $true)
{
	foreach ($file in (Get-ChildItem -LiteralPath "$RunLocation\output\Cluster_MSCluster_Network" | Where-Object {$_.name -match "~~Cluster_MSCluster_Network"}))
	{
		$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Cluster_MSCluster_Network\" + $file) 
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
	#EndRegion ClusterNetwork sheet	

#Region ClusterNode sheet
Write-Host -Object "---- Starting MSCluster_Node"
	$Worksheet = $Excel_Exch_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "MSCluster_Node"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Exch
	$row = 1
	$header = @()
	$header +=  "Cluster Name"
	$header +=  "Name"
	$header +=  "OS Version"
	$header +=  "EnableEventLogReplication"
	$header +=  "Cluster Network Priorities"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Cluster_MSCluster_Node") -eq $true)
{
	foreach ($file in (Get-ChildItem -LiteralPath "$RunLocation\output\Cluster_MSCluster_Node" | Where-Object {$_.name -match "~~Cluster_MSCluster_Node"}))
	{
		$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Cluster_MSCluster_Node\" + $file) 
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
	#EndRegion ClusterNode sheet		
	
#Region ClusterResource sheet
Write-Host -Object "---- Starting MSCluster_Resource"
	$Worksheet = $Excel_Exch_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "MSCluster_Resource"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Exch
	$row = 1
	$header = @()
	$header +=  "Cluster Name"
	$header +=  "Name"
	$header +=  "RestartAction"
	$header +=  "RestartPeriod"
	$header +=  "RestartThreshold"
	$header +=  "RetryPeriodOnFailure"
	$header +=  "State"
	$header +=  "Type"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Cluster_MSCluster_Resource") -eq $true)
{
	foreach ($file in (Get-ChildItem -LiteralPath "$RunLocation\output\Cluster_MSCluster_Resource" | Where-Object {$_.name -match "~~Cluster_MSCluster_Resource"}))
	{
		$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Cluster_MSCluster_Resource\" + $file) 
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
	#EndRegion ClusterResource sheet

#Region ClusterResourceGroup sheet
Write-Host -Object "---- Starting MSCluster_ResourceGroup"
	$Worksheet = $Excel_Exch_workbook.worksheets.item($intSheetCount)
	$Worksheet.name = "MSCluster_ResourceGroup"
	$Worksheet.Tab.ColorIndex = $intColorIndex_Exch
	$row = 1
	$header = @()
	$header +=  "Cluster Name"
	$header +=  "Name"
	$header +=  "FailoverPeriod"
	$header +=  "FailoverThreshold"
	$header +=  "State"
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

if ((Test-Path -LiteralPath "$RunLocation\output\Cluster_MSCluster_ResourceGroup") -eq $true)
{
	foreach ($file in (Get-ChildItem -LiteralPath "$RunLocation\output\Cluster_MSCluster_ResourceGroup" | Where-Object {$_.name -match "~~Cluster_MSCluster_ResourceGroup"}))
	{
		$DataFile += [System.IO.File]::ReadAllLines("$RunLocation\output\Cluster_MSCluster_ResourceGroup\" + $file) 
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
	#EndRegion ClusterResourceGroup sheet	
	
# Autofit columns
Write-Host -Object "---- Starting Autofit"	
$excel_ExchWorksheetCount = $excel_Exch_workbook.worksheets.count
$AutofitSheetCount = 1
while ($AutofitSheetCount -le $excel_ExchWorksheetCount)
{
	$ActiveWorksheet = $excel_Exch_workbook.worksheets.item($AutofitSheetCount)
	$objRange = $ActiveWorksheet.usedrange
	[Void]	$objRange.entirecolumn.autofit()
	$AutofitSheetCount++
}
$excel_Exch_workbook.saveas($ExDC_Exch_XLS)
Write-Host -Object "---- Spreadsheet saved"
$excel_Exch.workbooks.close()
Write-Host -Object "---- Workbook closed"
$Excel_Exch.quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel_Exch)
Remove-Variable -Name Excel_Exch
# If the ReleaseComObject doesn't do it..
#spps -n excel

	$EventLog = New-Object -TypeName System.Diagnostics.EventLog -ArgumentList Application
	$EventLog.MachineName = "."
	$EventLog.Source = "ExDC"
	#$EventLog.WriteEntry("Ending Core_Assemble_Exch_Excel","Information", 43)	
