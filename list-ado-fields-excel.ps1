# Azure DevOps Field Discovery Script - Excel Output
# This script discovers all fields (including custom fields) from Azure DevOps work items
# and exports them to a single Excel file with multiple worksheets

# Load configuration
. .\config.ps1

Write-Host "=== Azure DevOps Field Discovery - Excel Export ===" -ForegroundColor Green
Write-Host "Organization: $($ProductionConfig.AdoOrganizationUrl)" -ForegroundColor Cyan
Write-Host "Project: $($ProductionConfig.AdoProjectName)" -ForegroundColor Cyan

# Check if ImportExcel module is available
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "ImportExcel module not found. Installing..." -ForegroundColor Yellow
    try {
        Install-Module -Name ImportExcel -Scope CurrentUser -Force
        Write-Host "ImportExcel module installed successfully" -ForegroundColor Green
    } catch {
        Write-Host "Failed to install ImportExcel module: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Please install manually: Install-Module -Name ImportExcel" -ForegroundColor Yellow
        return
    }
}

# Import the module
Import-Module ImportExcel

# Setup headers for Azure DevOps API
$headers = @{
    'Authorization' = "Basic $([Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$($ProductionConfig.PersonalAccessToken)")))"
    'Content-Type' = 'application/json'
}

try {
    Write-Host "Discovering fields from Azure DevOps work items..." -ForegroundColor Yellow    # Use the existing WIQL query from config (we know this works) and limit results
    $wiqlQuery = $ProductionConfig.WiqlQuery
    
    $wiqlUrl = "$($ProductionConfig.AdoOrganizationUrl)/$($ProductionConfig.AdoProjectName)/_apis/wit/wiql?api-version=7.1"
    $wiqlRequest = @{ query = $wiqlQuery }
    $wiqlBody = $wiqlRequest | ConvertTo-Json -Depth 3
    
    $wiqlResponse = Invoke-RestMethod -Uri $wiqlUrl -Method Post -Headers $headers -Body $wiqlBody
    
    if (-not $wiqlResponse.workItems -or $wiqlResponse.workItems.Count -eq 0) {
        Write-Host "No work items found. Check your project name and query." -ForegroundColor Red
        return
    }
    
    Write-Host "Found $($wiqlResponse.workItems.Count) work items. Analyzing first 50 for field discovery..." -ForegroundColor Green
    
    # Limit to first 50 work items to avoid processing too many
    $workItemsToProcess = $wiqlResponse.workItems | Select-Object -First 50
    
    # Data structures for field discovery
    $allFields = @{}
    $fieldsByType = @{}
      # Analyze each work item to discover fields
    $processed = 0
    foreach ($workItem in $workItemsToProcess) {
        $processed++
        Write-Progress -Activity "Analyzing work items for fields" -Status "Processing item $($workItem.id)" -PercentComplete (($processed / $workItemsToProcess.Count) * 100)
        
        try {
            # Get detailed work item information
            $detailUrl = "$($ProductionConfig.AdoOrganizationUrl)/$($ProductionConfig.AdoProjectName)/_apis/wit/workitems/$($workItem.id)?`$expand=All&api-version=7.1"
            $workItemDetail = Invoke-RestMethod -Uri $detailUrl -Method Get -Headers $headers
            
            $workItemType = $workItemDetail.fields.'System.WorkItemType'
            
            # Initialize work item type if not exists
            if (-not $fieldsByType.ContainsKey($workItemType)) {
                $fieldsByType[$workItemType] = @{}
            }
            
            # Process each field in the work item
            foreach ($fieldProperty in $workItemDetail.fields.PSObject.Properties) {
                $fieldName = $fieldProperty.Name
                $fieldValue = $fieldProperty.Value
                
                # Initialize field if not exists
                if (-not $allFields.ContainsKey($fieldName)) {
                    $allFields[$fieldName] = @{
                        ReferenceName = $fieldName
                        Type = if ($null -eq $fieldValue) { "Unknown" } 
                               elseif ($fieldValue -is [DateTime]) { "DateTime" }
                               elseif ($fieldValue -is [int] -or $fieldValue -is [double]) { "Number" }
                               elseif ($fieldValue -is [bool]) { "Boolean" }
                               else { "String" }
                        IsCustom = $fieldName -like "Custom.*"
                        IsSystem = $fieldName -like "System.*"
                        IsMicrosoft = $fieldName -like "Microsoft.*"
                        IsDateField = $fieldName -like "*Date*" -or $fieldName -like "*Time*" -or $fieldName -like "*Due*"
                        WorkItemTypes = @()
                        SampleValues = @()
                    }
                }
                
                # Add work item type to field's usage
                if ($allFields[$fieldName].WorkItemTypes -notcontains $workItemType) {
                    $allFields[$fieldName].WorkItemTypes += $workItemType
                }
                
                # Add sample value
                if ($null -ne $fieldValue -and $allFields[$fieldName].SampleValues.Count -lt 3) {
                    $sampleValue = $fieldValue.ToString()
                    if ($sampleValue.Length -gt 100) {
                        $sampleValue = $sampleValue.Substring(0, 97) + "..."
                    }
                    if ($allFields[$fieldName].SampleValues -notcontains $sampleValue) {
                        $allFields[$fieldName].SampleValues += $sampleValue
                    }
                }
                
                # Track field usage by type
                $fieldsByType[$workItemType][$fieldName] = $fieldValue
            }
            
        } catch {
            Write-Host "Error processing work item $($workItem.id): $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    
    Write-Progress -Activity "Analyzing work items for fields" -Completed
    
    # Prepare data for Excel export
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $excelPath = "C:\temp\ADO_FieldDiscovery_$timestamp.xlsx"
    
    Write-Host "`nPreparing Excel export..." -ForegroundColor Yellow
    
    # 1. All Fields worksheet
    $allFieldsData = $allFields.Values | Sort-Object ReferenceName | ForEach-Object {
        [PSCustomObject]@{
            'Field Name' = $_.ReferenceName
            'Field Type' = $_.Type
            'Category' = if ($_.IsCustom) { "Custom" } 
                        elseif ($_.IsSystem) { "System" }
                        elseif ($_.IsMicrosoft) { "Microsoft VSTS" }
                        else { "Other" }
            'Is Date Field' = if ($_.IsDateField) { "Yes" } else { "No" }
            'Used In Work Item Types' = ($_.WorkItemTypes | Sort-Object) -join ", "
            'Sample Values' = ($_.SampleValues) -join " | "
            'Used In Export Config' = if ($ProductionConfig.FieldsToFetch -contains $_.ReferenceName) { "Yes" } else { "No" }
        }
    }
    
    # 2. Custom Fields worksheet
    $customFieldsData = $allFields.Values | Where-Object { $_.IsCustom } | Sort-Object ReferenceName | ForEach-Object {
        [PSCustomObject]@{
            'Field Name' = $_.ReferenceName
            'Field Type' = $_.Type
            'Used In Work Item Types' = ($_.WorkItemTypes | Sort-Object) -join ", "
            'Sample Values' = ($_.SampleValues) -join " | "
            'Used In Export Config' = if ($ProductionConfig.FieldsToFetch -contains $_.ReferenceName) { "Yes" } else { "No" }
            'Notes' = ""
        }
    }
    
    # 3. Date Fields worksheet
    $dateFieldsData = $allFields.Values | Where-Object { $_.IsDateField } | Sort-Object ReferenceName | ForEach-Object {
        [PSCustomObject]@{
            'Field Name' = $_.ReferenceName
            'Field Type' = $_.Type
            'Category' = if ($_.IsCustom) { "Custom" } 
                        elseif ($_.IsSystem) { "System" }
                        elseif ($_.IsMicrosoft) { "Microsoft VSTS" }
                        else { "Other" }
            'Used In Work Item Types' = ($_.WorkItemTypes | Sort-Object) -join ", "
            'Sample Values' = ($_.SampleValues) -join " | "
            'Used In Export Config' = if ($ProductionConfig.FieldsToFetch -contains $_.ReferenceName) { "Yes" } else { "No" }
            'Used For Finish Date' = if ($_.ReferenceName -in @("Custom.RevisedDueDate", "Custom.OriginalDueDate", "Microsoft.VSTS.Scheduling.TargetDate")) { "Yes" } else { "No" }
        }
    }
    
    # 4. Summary worksheet
    $summaryData = @(
        [PSCustomObject]@{ 'Metric' = "Discovery Date"; 'Value' = (Get-Date -Format "yyyy-MM-dd HH:mm:ss") }
        [PSCustomObject]@{ 'Metric' = "Organization"; 'Value' = $ProductionConfig.AdoOrganizationUrl }
        [PSCustomObject]@{ 'Metric' = "Project"; 'Value' = $ProductionConfig.AdoProjectName }
        [PSCustomObject]@{ 'Metric' = "Work Items Analyzed"; 'Value' = $workItemsToProcess.Count }
        [PSCustomObject]@{ 'Metric' = "Total Fields Found"; 'Value' = $allFields.Count }
        [PSCustomObject]@{ 'Metric' = "Custom Fields"; 'Value' = ($allFields.Values | Where-Object IsCustom).Count }
        [PSCustomObject]@{ 'Metric' = "System Fields"; 'Value' = ($allFields.Values | Where-Object IsSystem).Count }
        [PSCustomObject]@{ 'Metric' = "Microsoft VSTS Fields"; 'Value' = ($allFields.Values | Where-Object IsMicrosoft).Count }
        [PSCustomObject]@{ 'Metric' = "Date Fields"; 'Value' = ($allFields.Values | Where-Object IsDateField).Count }
        [PSCustomObject]@{ 'Metric' = "Fields in Export Config"; 'Value' = ($ProductionConfig.FieldsToFetch | Measure-Object).Count }
    )
    
    # Export to Excel
    Write-Host "Creating Excel file: $excelPath" -ForegroundColor Yellow
    
    # Export each worksheet
    $allFieldsData | Export-Excel -Path $excelPath -WorksheetName "All Fields" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    
    if ($customFieldsData.Count -gt 0) {
        $customFieldsData | Export-Excel -Path $excelPath -WorksheetName "Custom Fields" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    }
    
    if ($dateFieldsData.Count -gt 0) {
        $dateFieldsData | Export-Excel -Path $excelPath -WorksheetName "Date Fields" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    }
    
    $summaryData | Export-Excel -Path $excelPath -WorksheetName "Summary" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow
    
    # Console output summary
    Write-Host "`n=== FIELD DISCOVERY COMPLETE ===" -ForegroundColor Green
    Write-Host "Excel file created: $excelPath" -ForegroundColor Cyan
    Write-Host "File size: $([math]::Round((Get-Item $excelPath).Length / 1KB, 2)) KB" -ForegroundColor Cyan
    
    Write-Host "`n=== SUMMARY ===" -ForegroundColor Yellow
    Write-Host "Total fields found: $($allFields.Count)" -ForegroundColor White
    Write-Host "Custom fields: $(($allFields.Values | Where-Object IsCustom).Count)" -ForegroundColor Magenta
    Write-Host "System fields: $(($allFields.Values | Where-Object IsSystem).Count)" -ForegroundColor White
    Write-Host "Microsoft VSTS fields: $(($allFields.Values | Where-Object IsMicrosoft).Count)" -ForegroundColor White
    Write-Host "Date fields: $(($allFields.Values | Where-Object IsDateField).Count)" -ForegroundColor Yellow
    
    if (($allFields.Values | Where-Object IsCustom).Count -gt 0) {
        Write-Host "`nCustom fields found:" -ForegroundColor Magenta
        $allFields.Values | Where-Object IsCustom | Sort-Object ReferenceName | ForEach-Object {
            Write-Host "  - $($_.ReferenceName)" -ForegroundColor White
        }
    }
    
    Write-Host "`nDate fields found:" -ForegroundColor Yellow
    $allFields.Values | Where-Object IsDateField | Sort-Object ReferenceName | ForEach-Object {
        $category = if ($_.IsCustom) { "[CUSTOM]" } elseif ($_.IsSystem) { "[SYSTEM]" } else { "[VSTS]" }
        Write-Host "  - $($_.ReferenceName) $category" -ForegroundColor White
    }
    
    Write-Host "`nExcel file contains the following worksheets:" -ForegroundColor Cyan
    Write-Host "  - All Fields: Complete list of all discovered fields" -ForegroundColor White
    if ($customFieldsData.Count -gt 0) {
        Write-Host "  - Custom Fields: Custom fields specific to your project" -ForegroundColor White
    }
    if ($dateFieldsData.Count -gt 0) {
        Write-Host "  - Date Fields: Date/time related fields for scheduling" -ForegroundColor White
    }
    Write-Host "  - Summary: Overall statistics and metadata" -ForegroundColor White
    
    Write-Host "`nField discovery completed successfully!" -ForegroundColor Green
    
} catch {
    Write-Host "Error occurred during field discovery: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Full error details: $_" -ForegroundColor Red
}
