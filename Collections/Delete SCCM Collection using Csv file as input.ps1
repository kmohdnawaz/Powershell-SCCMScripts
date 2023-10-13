# SCCM Collection Deletion Script
# This script reads a list of SCCM collections from a CSV file, checks if they exist, determines their type (User or Device),
# and attempts to delete them. The results are then written to an Excel spreadsheet.

# Load the SCCM module to interact with SCCM via PowerShell
Import-Module ($Env:SMS_ADMIN_UI_PATH.Substring(0,$Env:SMS_ADMIN_UI_PATH.Length-5) + '\ConfigurationManager.psd1')

# Determine the SCCM site code dynamically
$siteCode = (Get-WmiObject -Namespace "root\SMS" -Class SMS_ProviderLocation).SiteCode

# Navigate to the SCCM site directory using the dynamically determined site code
cd "$siteCode:\"

# Initialize an instance of the Excel application and create a new workbook for output
$excelApp = New-Object -comobject Excel.Application
$excelApp.visible = $True
$workbook = $excelApp.Workbooks.Add()
$worksheet = $workbook.Worksheets.Item(1)

# Define and format the headers for the Excel worksheet
$worksheet.Cells.Item(1, 1) = "Collection Name"
$worksheet.Cells.Item(1, 2) = "Collection Type"
$worksheet.Cells.Item(1, 3) = "Status"

# Format the header row with colors and bold text
$usedRange = $worksheet.UsedRange
$usedRange.Interior.ColorIndex = 19
$usedRange.Font.ColorIndex = 11
$usedRange.Font.Bold = $True

# Initialize the starting row for data population
$row = 2

# Read collection names from the specified CSV file
$collections = Import-Csv -Path "C:\path\to\your\input.csv" | ForEach-Object { $_.CollectionName }

# Loop through and process each collection name
foreach ($collectionName in $collections) {
    # Ensure collection names are not empty or null
    if (![string]::IsNullOrEmpty($collectionName)) {
        try {
            # Query SCCM to see if the specified collection exists
            $collection = Get-CMCollection -Name $collectionName -ErrorAction Stop

            # If the collection exists, populate its details in the Excel sheet
            if ($collection) {
                $worksheet.Cells.Item($row, 1) = $collectionName
                $worksheet.Cells.Item($row, 2) = if ($collection.CollectionType -eq '1') { "User Collection" } else { "Device Collection" }
                
                # Attempt to delete the collection based on its type and record the status in Excel
                switch ($collection.CollectionType) {
                    '1' { 
                        Remove-CMUserCollection -Name $collectionName -Force -ErrorAction Stop
                        $worksheet.Cells.Item($row, 3) = "Deleted"
                    }
                    '2' { 
                        Remove-CMDeviceCollection -Name $collectionName -Force -ErrorAction Stop
                        $worksheet.Cells.Item($row, 3) = "Deleted"
                    }
                    default { 
                        $worksheet.Cells.Item($row, 3) = "Unknown collection type"
                    }
                }
            } else {
                # Record in Excel when a collection does not exist in SCCM
                $worksheet.Cells.Item($row, 1) = $collectionName
                $worksheet.Cells.Item($row, 2) = "N/A"
                $worksheet.Cells.Item($row, 3) = "Collection not found"
            }
        } catch {
            # Handle any exceptions and record the error details in Excel
            $errorMessage = "$_"
             if ($errorMessage -match "ErrorCode = 3242722566;") {
                $errorMessage = "This collection cannot be deleted as it's acting as a referenced collection"
            }
            $worksheet.Cells.Item($row, 1) = $collectionName
            $worksheet.Cells.Item($row, 2) = "Error"
            $worksheet.Cells.Item($row, 3) = $errorMessage
        }
        # Move to the next row in preparation for the next collection
        $row++
    }
}

# Adjust the Excel column widths to fit the content
$usedRange.EntireColumn.AutoFit()