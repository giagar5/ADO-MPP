# Create Microsoft Project Import Template and Mapping Instructions
# This script creates a template and mapping file to help with Microsoft Project import

# Configuration
$templatePath = "C:\temp\MSProject_Import_Template.xlsx"
$mappingInstructionsPath = "C:\temp\MSProject_Field_Mapping_Instructions.txt"

Write-Host "Creating Microsoft Project Import Template and Instructions..." -ForegroundColor Green

# Create field mapping instructions file
$mappingInstructions = @"
MICROSOFT PROJECT IMPORT FIELD MAPPING INSTRUCTIONS
==================================================

When importing the Excel file 'AzureDevOpsExport_ProjectImport.xlsx' into Microsoft Project, use the following field mappings:

TASK MAPPING TAB:
================
Excel Column Name          →  Microsoft Project Field Name
-----------------             ---------------------------
Unique ID                 →  ID (or Unique ID)
Name                      →  Name
Duration                  →  Duration
Start                     →  Start
Finish                    →  Finish
Predecessors              →  Predecessors
Resource Names            →  Resource Names
Outline Level             →  Outline Level
ADO ID                    →  Number1 (Azure DevOps Work Item ID)
Text1                     →  Text1 (Work Item Type: Epic, Feature, User Story, etc.)
Text2                     →  Text2 (Work Item State: New, Active, Done, Closed, etc.)
Text3                     →  Text3 (Area Path: Organizational hierarchy)
Text4                     →  Text4 (Work item tags: comma-separated)
Text5                     →  Text5 (Direct link to Azure DevOps work item)

RESOURCE MAPPING TAB:
====================
- Leave empty or skip this tab since resources are included in task mapping

ASSIGNMENT MAPPING TAB:
======================
- Leave empty or skip this tab since assignments are included in task mapping

STEP-BY-STEP IMPORT PROCESS:
===========================

1. Open Microsoft Project
2. Go to File → Open
3. Select "AzureDevOpsExport_ProjectImport.xlsx"
4. Choose "Project Excel Template" or "Excel Workbook" in the file type dropdown
5. In the Import Wizard:
   a. Select "Project Excel Template" if available
   b. Click "Next"
   c. Select the "Project Import" worksheet
   d. Click "Next"
   e. In the "Map" step, click on each mapping tab and set up the field mappings as shown above
   f. Make sure to map at least: ID, Name, Duration, Start, Finish, Predecessors, Resource Names, Outline Level
   g. Click "Finish"

TROUBLESHOOTING:
===============
- If you get "Map does not map any fields" error, make sure you've specified mappings in the Task Mapping tab
- Ensure the worksheet name is exactly "Project Import" (with space)
- The most critical fields are: ID, Name, Outline Level (for hierarchy)
- Predecessors field enables task dependencies
- If import fails, try importing without dates first, then add dates manually

ALTERNATIVE METHOD:
==================
If the Excel import continues to have issues, you can:
1. Save the Excel file as CSV format
2. Import the CSV file instead (File → Open → change file type to CSV)
3. Follow the same field mapping instructions

"@

# Write mapping instructions to file
$mappingInstructions | Out-File -FilePath $mappingInstructionsPath -Encoding UTF8
Write-Host "Created field mapping instructions: $mappingInstructionsPath" -ForegroundColor Yellow

# Check if the original Excel file exists and display its structure
$originalFile = "C:\temp\AzureDevOpsExport_ProjectImport.xlsx"
if (Test-Path $originalFile) {
    Write-Host "`nOriginal Excel file exists: $originalFile" -ForegroundColor Green
    
    try {
        Import-Module ImportExcel -ErrorAction Stop
        $data = Import-Excel $originalFile -WorksheetName "Project Import" | Select-Object -First 1
        
        Write-Host "`nColumn headers in the Excel file:" -ForegroundColor Cyan
        $data.PSObject.Properties.Name | ForEach-Object { Write-Host "  - $_" -ForegroundColor White }
        
        Write-Host "`nSample data:" -ForegroundColor Cyan
        $data | Format-List
        
    } catch {
        Write-Host "Could not read Excel file details: $($_.Exception.Message)" -ForegroundColor Red
    }
} else {
    Write-Host "Original Excel file not found: $originalFile" -ForegroundColor Red
}

Write-Host "`n=== NEXT STEPS ===" -ForegroundColor Green
Write-Host "1. Read the mapping instructions: $mappingInstructionsPath" -ForegroundColor Yellow
Write-Host "2. Open Microsoft Project" -ForegroundColor Yellow
Write-Host "3. Import the Excel file: $originalFile" -ForegroundColor Yellow
Write-Host "4. Use the field mappings provided in the instructions" -ForegroundColor Yellow
Write-Host "5. If you still have issues, run this script again for alternative solutions" -ForegroundColor Yellow
