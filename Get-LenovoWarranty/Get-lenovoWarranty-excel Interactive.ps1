param(
    [string]$ExcelFilePath, # Path to the Excel file
    [switch]$ThisDevice     # Switch to indicate whether to use the current device's serial number
)

# Initialize the Excel application and create a new workbook
$excelApp = New-Object -comobject Excel.Application
$excelApp.Visible = $true # Makes the Excel application visible
$workbook = $excelApp.Workbooks.Add() # Adds a new workbook
$worksheet = $workbook.Worksheets.Item(1) # Selects the first worksheet

# Define and format the headers for the Excel worksheet
$worksheet.Cells.Item(1, 1) = "Serial Number"
$worksheet.Cells.Item(1, 2) = "Model"
$worksheet.Cells.Item(1, 3) = "Status"
$worksheet.Cells.Item(1, 4) = "Is Active"
$worksheet.Cells.Item(1, 5) = "Start Date"
$worksheet.Cells.Item(1, 6) = "End Date"

# Formatting for the header row
$headerRange = $worksheet.Range("A1:F1")
$headerRange.Interior.ColorIndex = 19 # Background color for header
$headerRange.Font.ColorIndex = 11 # Font color for header
$headerRange.Font.Bold = $true # Makes the font bold

# Start populating data from the second row
$row = 2

# Read serial numbers from Excel or use the current device's serial number
if (!$ThisDevice) {
    try {
        # Create an Excel COM object to read the input file
        $Excel = New-Object -ComObject Excel.Application
        $Excel.Visible = $false
        $Excel.DisplayAlerts = $false

        # Open the Excel workbook
        $InputWorkbook = $Excel.Workbooks.Open($ExcelFilePath)
        $InputWorksheet = $InputWorkbook.Sheets.Item(1) # Assuming the data is in the first sheet
        $UsedRange = $InputWorksheet.UsedRange
        $Rows = $UsedRange.Rows.Count

        # Read serial numbers from the first column
        $SNList = @()
        for ($i = 1; $i -le $Rows; $i++) {
            $SNList += $InputWorksheet.Cells.Item($i, 1).Value2
        }

        # Close the input workbook without saving changes
        $InputWorkbook.Close($false)
        $Excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null

    } catch {
        Write-Warning "Error reading from Excel file: $($_.Exception.Message)"
        exit
    }
} else {
    # Get serial number of this device
    $SNList = @((Get-CimInstance Win32_BIOS).SerialNumber)
}

# Process each serial number
foreach ($serialNumber in $SNList) {
    try {
        # Fetch device information
        $Device_Info = Invoke-RestMethod "https://pcsupport.lenovo.com/us/en/api/v4/mse/getproducts?productId=$serialNumber"
        $Device_ID = $Device_Info.id
        $Warranty_url = "https://pcsupport.lenovo.com/us/en/products/$Device_ID/warranty"

    } catch {
        Write-Warning "Cannot get information for the serial number: $serialNumber"
        continue
    }

    try {
        # Fetch warranty information
        $Web_Response = Invoke-WebRequest -Uri $Warranty_url -Method GET
    } catch {
        Write-Warning "Cannot get warranty info for the serial number: $serialNumber"
        continue
    }

    if ($Web_Response.StatusCode -eq 200) {
        $HTML_Content = $Web_Response.Content

        # Extract necessary details using regex patterns
        $Pattern_Status = '"warrantystatus":"(.*?)"'
        $Pattern_Status2 = '"StatusV2":"(.*?)"'
        $Pattern_StartDate = '"Start":"(.*?)"'
        $Pattern_EndDate = '"End":"(.*?)"'
        $Pattern_DeviceModel = '"Name":"(.*?)"'
        
        $Status_Result = [regex]::Matches($HTML_Content, $Pattern_Status)[0].Groups[1].Value.Trim()
        $Statusv2_Result = [regex]::Matches($HTML_Content, $Pattern_Status2)[0].Groups[1].Value.Trim()
        $StartDate_Result = [regex]::Matches($HTML_Content, $Pattern_StartDate)[0].Groups[1].Value.Trim()
        $EndDate_Result = [regex]::Matches($HTML_Content, $Pattern_EndDate)[0].Groups[1].Value.Trim()
        $Model_Result = [regex]::Matches($HTML_Content, $Pattern_DeviceModel)[0].Groups[1].Value.Trim()
    } else {
        Write-Output "Failed to retrieve warranty information. Status Code: $($Web_Response.StatusCode)"
        continue
    }

    # Populate the Excel worksheet with the data
    $worksheet.Cells.Item($row, 1) = $serialNumber
    $worksheet.Cells.Item($row, 2) = $Model_Result
    $worksheet.Cells.Item($row, 3) = $Status_Result
    $worksheet.Cells.Item($row, 4) = $Statusv2_Result
    $worksheet.Cells.Item($row, 5) = $StartDate_Result
    $worksheet.Cells.Item($row, 6) = $EndDate_Result

    # Move to the next row
    $row++
}

# Auto adjust the width of the columns to fit the content
$worksheet.UsedRange.EntireColumn.AutoFit()

# Save the workbook to a desired location
$savePath = "C:\temp\WarrantyResults.xlsx"
$workbook.SaveAs($savePath)

# Inform the user about the successful completion
Write-Output "Warranty information has been successfully written to the Excel sheet and saved at $savePath."

# Release COM objects to free resources
$workbook.Close($true)
$excelApp.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelApp) | Out-Null
