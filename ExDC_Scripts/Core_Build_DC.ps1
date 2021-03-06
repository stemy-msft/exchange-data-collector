#############################################################################
#                          Core_Build_DC		 							#
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
$DCs_outputfile = $location.path + "\dc.txt"

if ((Test-Path -LiteralPath $DCs_outputfile) -eq $true)
	{
	Write-Host -Object "DC.txt is already present in this folder." -ForegroundColor Red
	Write-Host -Object "Loading values from text file that is present." -ForegroundColor Red
    $status_Step1.Text = "Step 1 Status: Failed - DC.txt is already present. Loading values from existing file."
	exit
	}
	

###########################
# Initialize Arrays
###########################
$arrayDCs = @()

###########################
# Populate Domain Controllers in Forest
###########################
$status_Step1.Text = "Step 1 Status: Running - Collecting list of domain controllers"
$Forest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
$arrayDCs += $forest.Domains | ForEach-Object {$_.DomainControllers} | ForEach-Object {$_.Name}  

$error.clear()

# Output each array to text files with dupes removed
$status_Step1.Text = "Step 1 Status: Running - Sorting arrays and writing output files."
	
foreach ($DC in $arrayDCs | Sort-Object -u) 
	{
		$DC | Out-File -FilePath $DCs_outputfile -append
	}
	    
$status_Step1.Text = "Step 1 Status: Idle"
