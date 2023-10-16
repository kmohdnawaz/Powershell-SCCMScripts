# SCCM Software Update Group Addition Script
# This script reads a list of SCCM software updates from a CSV file, checks if they exist and have been downloaded,
# and attempts to add them to the specified Software Update Group. The results are then written to an Excel spreadsheet.

# Load the SCCM module for PowerShell
# This allows PowerShell to interact with SCCM
Import-Module ($Env:SMS_ADMIN_UI_PATH.Substring(0,$Env:SMS_ADMIN_UI_PATH.Length-5) + '\ConfigurationManager.psd1')

# Determine the SCCM site code
# This is used to identify the correct SCCM site for further actions
$siteCode = (Get-WmiObject -Namespace "root\SMS" -Class SMS_ProviderLocation).SiteCode

# Change the current directory to the SCCM site directory
Set-Location $siteCode":"

# Initialize and configure Excel
# This section sets up Excel for recording the script's results
$excelApp = New-Object -comobject Excel.Application
$excelApp.visible = $True
$workbook = $excelApp.Workbooks.Add()
$worksheet = $workbook.Worksheets.Item(1)

# Set up Excel headers and format them
$worksheet.Cells.Item(1, 1) = "Update Title"
$worksheet.Cells.Item(1, 2) = "Status"
$usedRange = $worksheet.UsedRange
$usedRange.Interior.ColorIndex = 19
$usedRange.Font.ColorIndex = 11
$usedRange.Font.Bold = $True

# Start processing from the second row in Excel, as the first row contains headers
$row = 2

# Define the name of the Software Update Group (SUG) we're adding updates to
$sugName = "Baseline - Lenovo Driver Updates"

# Import update titles from the provided CSV file
$updates = Import-Csv -Path E:\Scripts\updatetitle.csv | ForEach-Object { $_.UpdateTitle }

# Loop through each update title
foreach ($updateTitle in $updates) {
    # Retrieve the specified Software Update Group
    $sug = Get-CMSoftwareUpdateGroup -Name $sugName

    # Check if the Software Update Group exists
    if ($sug) {
        # Search for the software update by its title
        $update = Get-CMSoftwareUpdate -Fast | Where-Object {$_.LocalizedDisplayName -eq $updateTitle} | Select-Object -First 1

        # Check if the software update exists
        if ($update) {
            # Check if the update content has been downloaded
            if ($update.IsContentProvisioned -eq $true) {
                # Add the software update to the Software Update Group
                Add-CMSoftwareUpdateToGroup -SoftwareUpdateName $update.LocalizedDisplayName -SoftwareUpdateGroupName $sug.LocalizedDisplayName
                
                # Record the result in Excel
                $worksheet.Cells.Item($row, 1) = $updateTitle
                $worksheet.Cells.Item($row, 2) = "Added to SUG"
            } else {
                # If the update hasn't been downloaded, note it in Excel
                $worksheet.Cells.Item($row, 1) = $updateTitle
                $worksheet.Cells.Item($row, 2) = "Not Downloaded"
            }
        } else {
            # If the software update was not found, note it in Excel
            $worksheet.Cells.Item($row, 1) = $updateTitle
            $worksheet.Cells.Item($row, 2) = "Update Not Found"
        }
    } else {
        # If the Software Update Group was not found, note it in Excel
        $worksheet.Cells.Item($row, 1) = $updateTitle
        $worksheet.Cells.Item($row, 2) = "SUG Not Found"
    }
    
    # Move to the next row in Excel for the next software update
    $row++
}

# Adjust the width of the columns in Excel based on the content
$usedRange.EntireColumn.AutoFit()