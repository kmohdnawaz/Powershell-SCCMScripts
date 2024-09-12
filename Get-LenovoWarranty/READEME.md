# Warranty Information Fetcher

This PowerShell script fetches warranty information for Lenovo devices using their serial numbers and saves the results in an interactive Excel file.

## Attribution

This script is based on an original script by [Syst & Deploy](https://www.systanddeploy.com/2024/08/using-powershell-to-know-if-lenovo.html). It has been modified to:
- Process multiple serial numbers from an Excel sheet.
- Display the warranty information interactively in a new Excel workbook.

## Requirements

- PowerShell
- Excel (with COM support enabled)
- Internet connection

## Usage

```powershell
.\YourScriptName.ps1 -ExcelFilePath "C:\Path\To\SerialNumbers.xlsx"

Parameters
ExcelFilePath: The path to the Excel file containing serial numbers.
ThisDevice: A switch to use the current device's serial number.
Output
The script saves the warranty information in an Excel file located at C:\temp\WarrantyResults.xlsx.