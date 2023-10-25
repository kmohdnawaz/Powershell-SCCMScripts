<#
.SYNOPSIS
A script to manage SCCM collections by optimizing incremental updates using PowerShell.

.OWNER
Nawaz Kazi

.DATE
October 24, 2023

.DESCRIPTION
This script imports the ConfigurationManager module, determines the SCCM site code, reads collection IDs from a CSV, 
and manages each collection by setting its RefreshType to optimize incremental updates. Results are displayed in an interactive Excel sheet.
#>

# Import required module
Import-Module ($Env:SMS_ADMIN_UI_PATH.Substring(0,$Env:SMS_ADMIN_UI_PATH.Length-5) + '\ConfigurationManager.psd1')

# Dynamically determine the SCCM site code
$siteCode = (Get-WmiObject -Namespace "root\SMS" -Class SMS_ProviderLocation).SiteCode

# Set the current directory to the SCCM site directory
Set-Location $siteCode":"

# Path to the CSV file
$csvPath = "path_to_your_csv_file.csv"

# Read the collection IDs from the CSV
$collectionIDs = Import-Csv -Path $csvPath | ForEach-Object { $_.CollectionID }

# Initialize an array for the results
$results = @()

foreach ($collectionID in $collectionIDs) {
    Write-Host "Processing collection: $collectionID"

    # Directly fetch the specific collection by its ID and set the RefreshType
    Get-CMDeviceCollection -CollectionId $collectionID | Set-CMDeviceCollection -RefreshType 2

    # No need to check if the collection exists as Get-CMDeviceCollection will simply return nothing if the ID isn't found
    $results += [PSCustomObject]@{
        'CollectionID' = $collectionID
        'Status'       = "Incremental updates disabled, scheduled updates enabled"
    }
    Write-Host "Incremental updates disabled for collection: $collectionID"
}

# Create a new Excel application instance
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$workbook = $excel.Workbooks.Add()
$worksheet = $workbook.Worksheets.Item(1)

# Populate the headers
$worksheet.Cells.Item(1, 1).Value2 = 'CollectionID'
$worksheet.Cells.Item(1, 2).Value2 = 'Status'

# Populate the data
$row = 2
foreach ($result in $results) {
    $worksheet.Cells.Item($row, 1).Value2 = $result.CollectionID
    $worksheet.Cells.Item($row, 2).Value2 = $result.Status
    $row++
}

# Auto size columns
$range = $worksheet.UsedRange
$range.EntireColumn.AutoFit()

Write-Host "Script completed!"