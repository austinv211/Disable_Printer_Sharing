<#
Name: printerdisable.ps1
Description: disables print queue if found in spreadsheet
Author: Austin Vargason
Date Modified: 6/15/18
#>

#function to disable the printer
function Disable-Printer {
    [CmdletBinding()]

    #takes printer name as a parameter
    param (
        [Parameter(Mandatory=$true)]
        [String]$PrinterName
    )

    #begin by trying to get the printer object
    begin {
        try {
            $printer = Get-Printer -Name $PrinterName -ErrorAction Stop
        }
        catch {
            $printer = $null
        }   
    }
    #set the printer shared to false and write to the console
    process {
        if ($printer -ne $null) {
            if ($printer.Shared) {
                $printer.Shared = $false
                Set-PrinterProperty -InputObject $printer
                Write-Host "Printer Sharing set to false for printer: $PrinterName" -ForegroundColor Green
            }
            else {
                Write-Host "Printer Sharing already set to false for printer: $PrinterName" -ForegroundColor Green
            }
        }
        else {
            Write-Host "Printer $PrinterName not located on Server" -ForegroundColor Cyan
        }
    }
}

#import the csv to use for disabling
$csv = Import-Csv -Path ..\printers.csv

#get the names of the print queues
$printerList = $csv | Select-Object -ExpandProperty "Old_Print_Queue"

#integer variable for counting
$i = 0

foreach ($p in $printerList) {
    
    if ($p -ne "") {
        #disable the printer
        Disable-Printer -PrinterName $p
    }

    #increase the counter
    $i++

    #write the progress
    Write-Progress -Activity "Disabling Printers" -Status "Disabled Printer: $p" -PercentComplete (($i/ $printerList.Count) * 100)
}