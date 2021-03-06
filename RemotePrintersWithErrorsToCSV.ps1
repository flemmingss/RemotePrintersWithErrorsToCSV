<#
.SYNOPSIS
  Name: RemotePrintersWithErrorsToCSV.ps1
  The purpose of this script is to obtain a list of printers with
  errors and when the error occurred from one or multiple servers.

.DESCRIPTION
  This script is made to create and mantain a CSV file updated
  with error information from one or multiple Print Servers.
  The main purpose is to have a file with infomation about
  which printers that has gone offline, or have an error and
  when the error occurred.

  This script is intended to run routinely with for example
  Windows Task Scheduler to make ErrorOccurredDate reliable.

  "Windows Remote Management" need to be enabled on the servers
  and the user who run the script needs the necessary permissions.
  The "PrintManagement" module must be installed on the servers and
  the script execution machine.

.NOTES
    Original release Date: 08.02.2018 Updated: 26.02.2018

  Author: Flemming Sørvollen Skaret (https://github.com/flemmingss/)

.LINK
  https://github.com/flemmingss/

.EXAMPLE
  Path To CSV file
  $CsvPath = "C:\PrintersWithErrors"
  $CsvPath = "\\server\share\PrintersWithErrors"

  Servers
  $Servers = "server-print-01"
  $Servers = "server-print-01","server-print-02","server-print-03"

  Properties of Printer Objects
  $ObjectProperties = "Name","ErrorOccurredDate"
  $ObjectProperties = "Name","PSComputerName","Location","PrinterStatus","ErrorOccurredDate"

  Important!: Property "Name" and "ErrorOccurredDate" are mandatory for this script to work

#requires -version 3
#Requires -Modules PrintManagement @{
  ModuleName="PrintManagement"
  ModuleVersion="1.1"
}
#>

Write-Output "Running script with the following configuration:"

### Configuration Start ###
$CsvPath = "C:\PrintersWithErrors.csv" #CSV file path
$Servers = "server-print-01","server-print-02" #Server(s)
$ObjectProperties = "Name","PSComputerName","Location","PrinterStatus","ErrorOccurredDate" #Printer Properties
### Configuration End ###

Write-Output "`t CSV file path: $CsvPath"
Write-Output "`t Server(s): $Servers"
Write-Output "`t Printer Properties: $ObjectProperties"

$RemoteScript = {Get-Printer | Where-Object {($_.PrinterStatus -like "*Error*") -or ($_.PrinterStatus -like "*Offline*")}}
$NewResults = Invoke-Command -ComputerName $Servers -ScriptBlock $RemoteScript
$NewResults | ForEach-Object -Begin {$Date = Get-Date} `
                             -process { $_ | Add-Member -Name ErrorOccurredDate -MemberType NoteProperty -Value $Date}
$NewResults = ($NewResults | Select-Object $ObjectProperties | ConvertTo-Csv|  ConvertFrom-CSV)

Import-Module PrintManagement

try
	{
	$NewResults = Invoke-Command -ComputerName $Servers -ScriptBlock $RemoteScript -ErrorAction Stop
	$NewResults | ForEach-Object -Begin {$Date = Get-Date} `
                 	  	     	-process { $_ | Add-Member -Name ErrorOccurredDate -MemberType NoteProperty -Value $Date}
	$NewResults = ($NewResults | Select-Object $ObjectProperties | ConvertTo-Csv|  ConvertFrom-CSV)

	if (Test-Path -Path $CsvPath)
		{
		Write-Output "CSV file found, importing data"
		$OldResults = Import-Csv -Path "$CsvPath"
		$NewErrorsName = (Compare-Object -ReferenceObject $NewResults -DifferenceObject $OldResults -Property Name | Where-Object {$_.SideIndicator -eq "<="}).Name
		$EqualErrorsName = (Compare-Object -ReferenceObject $NewResults -DifferenceObject $OldResults -Property Name -IncludeEqual | Where-Object {$_.SideIndicator -eq "=="}).Name
		$NewErrorsToExport = $NewResults | Where-Object { $NewErrorsName -contains $_.Name }
		$EqualErrorsToExport = $OldResults | Where-Object { $EqualErrorsName -contains $_.Name }
		$UpdatedErrors = $EqualErrorsToExport + $NewErrorsToExport
		}
	else
		{
		Write-Output "CSV file not found. This script will automatically create the file"
		$UpdatedErrors = $NewResults
		}	
		
	Write-Output "Exporting data to CSV file"
	$UpdatedErrors | Select-Object $ObjectProperties | Export-Csv "$CsvPath" -NoTypeInformation		
	}
catch
	{
    Write-Output "Error in execution detected. Aborting export to prevent errors in CSV file"
	}
finally
	{
	Clear-Variable CsvPath,Servers,ObjectProperties,RemoteScript,NewResults,Date,OldResults,NewErrorsName,EqualErrorsName,NewErrorsToExport,EqualErrorsToExport,UpdatedErrors -ErrorAction SilentlyContinue
	Write-Output "Script execution complete"
	}

# End for Script
