# Enhanced Critical Timeline Export for PowerPoint Office Timeline Expert
# This script exports Azure DevOps Milestones and Dependencies based on priority criteria
# Works with existing tags or can be configured to export high-priority items

param(
    [string]$ConfigPath = ".\config\config.ps1",
    [string]$OutputPath = "",
    [switch]$DebugMode = $false,
    [string[]]$PriorityTags = @(),
    [switch]$ExportAll = $false,
    [string[]]$IncludeStates = @("New", "Active", "In Progress", "Done", "Closed")
)

# Load configuration
if (Test-Path $ConfigPath) {
    . $ConfigPath
} else {
    Write-Error "Configuration file not found: $ConfigPath"
    exit 1
}

# Set default output path if not provided
if ([string]::IsNullOrEmpty($OutputPath)) {
    $timestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
    $OutputPath = "C:\temp\CriticalTimeline_$timestamp.xlsx"
}

# Set default priority tags if none provided
if ($PriorityTags.Count -eq 0 -and -not $ExportAll) {
    $PriorityTags = @("Critical", "MG1", "DQT-Phase1", "Modernization", "DataProduct")
    Write-Host "Using default priority tags: $($PriorityTags -join ', ')" -ForegroundColor Yellow
}

Write-Host "=== Enhanced Critical Timeline Export for Office Timeline Expert ===" -ForegroundColor Green
Write-Host "Output file: $OutputPath" -ForegroundColor Cyan
if ($ExportAll) {
    Write-Host "Mode: Export ALL Milestones and Dependencies" -ForegroundColor Magenta
} else {
    Write-Host "Priority tags: $($PriorityTags -join ', ')" -ForegroundColor Cyan
}

# Import required modules for Excel export
try {
    Import-Module ImportExcel -Force -ErrorAction Stop
    Write-Host "ImportExcel module loaded successfully" -ForegroundColor Green
} catch {
    Write-Host "ImportExcel module not found. Installing..." -ForegroundColor Yellow
    try {
        Install-Module ImportExcel -Force -Scope CurrentUser
        Import-Module ImportExcel -Force
        Write-Host "ImportExcel module installed and loaded successfully" -ForegroundColor Green
    } catch {
        Write-Error "Failed to install ImportExcel module: $($_.Exception.Message)"
        Write-Host "Please install manually: Install-Module ImportExcel" -ForegroundColor Red
        exit 1
    }
}

# Setup headers for Azure DevOps API with proper authentication
$headers = @{
    'Authorization' = "Basic $([Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$($ProductionConfig.PersonalAccessToken)")))"
    'Content-Type' = 'application/json'
    'User-Agent' = 'PowerShell-EnhancedTimelineExport/1.0'
}

# Function to safely get field value with null checking
function Get-SafeFieldValue {
    param(
        [object]$WorkItem,
        [string]$FieldName,
        [string]$DefaultValue = ""
    )
    
    try {
        if ($WorkItem.fields -and $WorkItem.fields.PSObject.Properties[$FieldName]) {
            $value = $WorkItem.fields.$FieldName
            if ($null -eq $value -or $value -eq "") {
                return $DefaultValue
            }
            return $value.ToString()
        }
        return $DefaultValue
    } catch {
        if ($DebugMode) {
            Write-Host "Warning: Could not get field '$FieldName': $($_.Exception.Message)" -ForegroundColor Yellow
        }
        return $DefaultValue
    }
}

# Function to check if item matches priority criteria
function Test-PriorityItem {
    param(
        [object]$WorkItem
    )
    
    if ($ExportAll) {
        return $true
    }
    
    # Check tags
    $tags = Get-SafeFieldValue -WorkItem $WorkItem -FieldName "System.Tags"
    if ($tags) {
        foreach ($priorityTag in $PriorityTags) {
            if ($tags.ToLower().Contains($priorityTag.ToLower())) {
                return $true
            }
        }
    }
    
    # Check title for priority indicators
    $title = Get-SafeFieldValue -WorkItem $WorkItem -FieldName "System.Title"
    $priorityKeywords = @("critical", "urgent", "milestone", "key", "important", "phase1", "phase 1", "mg1")
    foreach ($keyword in $priorityKeywords) {
        if ($title.ToLower().Contains($keyword.ToLower())) {
            return $true
        }
    }
    
    # Check if it's a milestone (milestones are generally important)
    $workItemType = Get-SafeFieldValue -WorkItem $WorkItem -FieldName "System.WorkItemType"
    if ($workItemType -eq "Milestone") {
        return $true
    }
    
    return $false
}

# Function to get priority level
function Get-PriorityLevel {
    param(
        [object]$WorkItem
    )
    
    $tags = Get-SafeFieldValue -WorkItem $WorkItem -FieldName "System.Tags"
    $title = Get-SafeFieldValue -WorkItem $WorkItem -FieldName "System.Title"
    $workItemType = Get-SafeFieldValue -WorkItem $WorkItem -FieldName "System.WorkItemType"
    
    # Highest priority
    if ($tags.ToLower().Contains("critical") -or $title.ToLower().Contains("critical")) {
        return "Critical"
    }
    
    # High priority
    if ($tags.ToLower().Contains("mg1") -or $tags.ToLower().Contains("phase1") -or $tags.ToLower().Contains("dqt-phase1")) {
        return "High"
    }
    
    # Medium priority for milestones or important dependencies
    if ($workItemType -eq "Milestone" -or $tags.ToLower().Contains("modernization") -or $tags.ToLower().Contains("dataproduct")) {
        return "Medium"
    }
    
    return "Normal"
}

# Function to format date for Office Timeline
function Format-TimelineDate {
    param(
        [string]$DateString
    )
    
    if ([string]::IsNullOrEmpty($DateString)) {
        return ""
    }
    
    try {
        $date = [DateTime]::Parse($DateString)
        return $date.ToString("MM/dd/yyyy")
    } catch {
        if ($DebugMode) {
            Write-Host "Warning: Could not parse date '$DateString'" -ForegroundColor Yellow
        }
        return ""
    }
}

# Function to retry API calls with exponential backoff
function Invoke-ApiWithRetry {
    param(
        [string]$Uri,
        [hashtable]$Headers,
        [string]$Method = "GET",
        [object]$Body = $null,
        [int]$MaxRetries = 3,
        [int]$BaseDelaySeconds = 2
    )
    
    for ($attempt = 1; $attempt -le $MaxRetries; $attempt++) {
        try {
            $params = @{
                Uri = $Uri
                Headers = $Headers
                Method = $Method
                TimeoutSec = $ProductionConfig.ApiTimeout
            }
            
            if ($Body) {
                $params.Body = $Body
            }
            
            return Invoke-RestMethod @params
        } catch {
            $statusCode = $_.Exception.Response.StatusCode
            $statusDescription = $_.Exception.Response.StatusDescription
            
            if ($attempt -eq $MaxRetries) {
                throw "API call failed after $MaxRetries attempts. Status: $statusCode - $statusDescription. Error: $($_.Exception.Message)"
            }
            
            $delay = $BaseDelaySeconds * [Math]::Pow(2, $attempt - 1)
            Write-Host "API call attempt $attempt failed (Status: $statusCode). Retrying in $delay seconds..." -ForegroundColor Yellow
            Start-Sleep -Seconds $delay
        }
    }
}

try {    # Build WIQL query for Milestones and Dependencies
    $itemsQuery = @"
SELECT [System.Id] 
FROM WorkItems 
WHERE [System.TeamProject] = '$($ProductionConfig.AdoProjectName)' 
AND [System.WorkItemType] IN ('Milestone', 'Dependency')
AND [System.State] <> 'Removed'
AND ([System.AreaPath] UNDER 'Azure-Cloud-Transformation-Program\Workstream D- Data Estate Modernization (E07)' OR [System.Id] = 2449)
ORDER BY [System.Id]
"@

    Write-Host "`nQuerying for Milestones and Dependencies..." -ForegroundColor Yellow
    if ($DebugMode) {
        Write-Host "WIQL Query:" -ForegroundColor Gray
        Write-Host $itemsQuery -ForegroundColor DarkGray
    }
    
    # Execute WIQL query
    $wiqlUrl = "$($ProductionConfig.AdoOrganizationUrl)/$($ProductionConfig.AdoProjectName)/_apis/wit/wiql?api-version=7.1"
    $wiqlRequest = @{ query = $itemsQuery }
    $wiqlBody = $wiqlRequest | ConvertTo-Json -Depth 3
    
    $queryResponse = Invoke-ApiWithRetry -Uri $wiqlUrl -Headers $headers -Method "POST" -Body $wiqlBody
    
    if (-not $queryResponse.workItems -or $queryResponse.workItems.Count -eq 0) {
        Write-Host "No Milestones or Dependencies found in the specified scope." -ForegroundColor Red
        exit 0
    }
    
    Write-Host "Found $($queryResponse.workItems.Count) Milestone(s) and Dependency(ies)" -ForegroundColor Green
    
    # Get detailed work item information in batches
    $timelineItems = @()
    $batchSize = $ProductionConfig.BatchSize
    $workItemIds = $queryResponse.workItems | ForEach-Object { $_.id }
    
    Write-Host "Retrieving detailed work item information..." -ForegroundColor Yellow
    
    for ($i = 0; $i -lt $workItemIds.Count; $i += $batchSize) {
        $batch = $workItemIds[$i..([Math]::Min($i + $batchSize - 1, $workItemIds.Count - 1))]
        $batchIds = $batch -join ","
        
        Write-Progress -Activity "Processing work items" -Status "Batch $([Math]::Floor($i/$batchSize) + 1)" -PercentComplete (($i / $workItemIds.Count) * 100)
        
        $detailUrl = "$($ProductionConfig.AdoOrganizationUrl)/$($ProductionConfig.AdoProjectName)/_apis/wit/workitems?ids=$batchIds&`$expand=All&api-version=7.1"
        
        try {
            $batchResponse = Invoke-ApiWithRetry -Uri $detailUrl -Headers $headers
            
            foreach ($workItem in $batchResponse.value) {
                # Skip items in Removed state (additional safety check)
                $state = Get-SafeFieldValue -WorkItem $workItem -FieldName "System.State"
                if ($state -eq "Removed") {
                    if ($DebugMode) {
                        Write-Host "Skipping removed item: $($workItem.id)" -ForegroundColor DarkGray
                    }
                    continue
                }
                
                # Check if item matches priority criteria
                $isPriority = Test-PriorityItem -WorkItem $workItem
                
                if ($isPriority) {
                    $workItemType = Get-SafeFieldValue -WorkItem $workItem -FieldName "System.WorkItemType"
                    $title = Get-SafeFieldValue -WorkItem $workItem -FieldName "System.Title"
                    $assignedTo = Get-SafeFieldValue -WorkItem $workItem -FieldName "System.AssignedTo"
                    $description = Get-SafeFieldValue -WorkItem $workItem -FieldName "System.Description"
                    $priority = Get-PriorityLevel -WorkItem $workItem
                    
                    # Extract display name from assigned to field
                    $assignedToName = ""
                    if ($assignedTo -match '<([^>]+)>') {
                        $assignedToName = $matches[1]
                    } elseif ($assignedTo) {
                        $assignedToName = $assignedTo
                    }
                    
                    # Get dates - try multiple date fields
                    $targetDate = Format-TimelineDate (Get-SafeFieldValue -WorkItem $workItem -FieldName "Microsoft.VSTS.Scheduling.TargetDate")
                    $startDateRaw = Get-SafeFieldValue -WorkItem $workItem -FieldName "Microsoft.VSTS.Scheduling.StartDate"
                    $finishDateRaw = Get-SafeFieldValue -WorkItem $workItem -FieldName "Microsoft.VSTS.Scheduling.FinishDate"
                    
                    $startDate = ""
                    $finishDate = ""
                    
                    if ($startDateRaw) {
                        $startDate = Format-TimelineDate $startDateRaw
                    } elseif ($targetDate) {
                        $startDate = $targetDate
                    } else {
                        # Default to today for items without dates
                        $startDate = (Get-Date).ToString("MM/dd/yyyy")
                    }
                    
                    if ($finishDateRaw) {
                        $finishDate = Format-TimelineDate $finishDateRaw
                    } elseif ($targetDate) {
                        $finishDate = $targetDate
                    } else {
                        $finishDate = $startDate
                    }
                    
                    # Create timeline item object formatted for Office Timeline Expert
                    $timelineItem = [PSCustomObject]@{
                        'Task Name' = "$($workItemType): $title"
                        'Start Date' = $startDate
                        'End Date' = $finishDate
                        'Duration (Days)' = if ($startDate -and $finishDate -and $startDate -ne $finishDate) { 
                            try {
                                $start = [DateTime]::Parse($startDate)
                                $end = [DateTime]::Parse($finishDate)
                                [Math]::Max(1, ($end - $start).Days)
                            } catch { 1 }
                        } else { 1 }
                        'Milestone' = if ($workItemType -eq "Milestone") { "Yes" } else { "No" }
                        'Resource Names' = $assignedToName
                        'Percent Complete' = if ($state -eq "Done" -or $state -eq "Closed") { 100 } 
                                           elseif ($state -eq "Active" -or $state -eq "In Progress") { 50 } 
                                           else { 0 }
                        'Priority' = $priority
                        'Notes' = if ($description.Length -gt 255) { $description.Substring(0, 252) + "..." } else { $description }
                        'Work Item ID' = $workItem.id
                        'Work Item Type' = $workItemType
                        'State' = $state
                        'Tags' = Get-SafeFieldValue -WorkItem $workItem -FieldName "System.Tags"
                        'Area Path' = Get-SafeFieldValue -WorkItem $workItem -FieldName "System.AreaPath"
                    }
                    
                    $timelineItems += $timelineItem
                    
                    if ($DebugMode) {
                        Write-Host "Added $priority priority $workItemType`: $($workItem.id) - $title" -ForegroundColor Green
                    }
                } else {
                    if ($DebugMode) {
                        $title = Get-SafeFieldValue -WorkItem $workItem -FieldName "System.Title"
                        Write-Host "Skipped non-priority item: $($workItem.id) - $title" -ForegroundColor DarkGray
                    }
                }
            }
        } catch {
            Write-Host "Error processing batch starting at index $i`: $($_.Exception.Message)" -ForegroundColor Red
            if ($DebugMode) {
                Write-Host "Full error: $_" -ForegroundColor Red
            }
        }
    }
    
    Write-Progress -Activity "Processing work items" -Completed
    
    if ($timelineItems.Count -eq 0) {
        Write-Host "No priority items found with the specified criteria." -ForegroundColor Red
        if (-not $ExportAll) {
            Write-Host "Consider using -ExportAll switch to export all items, or adjust -PriorityTags parameter." -ForegroundColor Yellow
        }
        exit 0
    }
    
    Write-Host "`nFound $($timelineItems.Count) priority timeline item(s)" -ForegroundColor Green
    
    # Display summary by priority
    $prioritySummary = $timelineItems | Group-Object Priority | Sort-Object Name
    foreach ($group in $prioritySummary) {
        Write-Host "  $($group.Name): $($group.Count) items" -ForegroundColor White
    }
    
    # Sort by priority and start date for better timeline visualization
    $priorityOrder = @{ "Critical" = 1; "High" = 2; "Medium" = 3; "Normal" = 4 }
    $timelineItems = $timelineItems | Sort-Object { $priorityOrder[$_.Priority] }, { 
        try { [DateTime]::Parse($_.'Start Date') } catch { [DateTime]::MaxValue }
    }
    
    # Create Excel file with Office Timeline Expert compatible format
    Write-Host "`nExporting to Excel file: $OutputPath" -ForegroundColor Yellow
    
    # Ensure output directory exists
    $outputDir = Split-Path $OutputPath -Parent
    if (-not (Test-Path $outputDir)) {
        New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
    }
    
    # Remove existing file if it exists
    if (Test-Path $OutputPath) {
        Remove-Item $OutputPath -Force
    }
    
    # Export to Excel with formatting optimized for Office Timeline Expert
    $timelineItems | Export-Excel -Path $OutputPath -WorksheetName "Priority Timeline" -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow
    
    # Add a summary sheet with metadata
    $summaryData = @(
        [PSCustomObject]@{ Property = "Export Date"; Value = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss") }
        [PSCustomObject]@{ Property = "Total Priority Items"; Value = $timelineItems.Count }
        [PSCustomObject]@{ Property = "Milestones"; Value = ($timelineItems | Where-Object { $_.'Work Item Type' -eq "Milestone" }).Count }
        [PSCustomObject]@{ Property = "Dependencies"; Value = ($timelineItems | Where-Object { $_.'Work Item Type' -eq "Dependency" }).Count }
        [PSCustomObject]@{ Property = "Critical Priority"; Value = ($timelineItems | Where-Object { $_.Priority -eq "Critical" }).Count }
        [PSCustomObject]@{ Property = "High Priority"; Value = ($timelineItems | Where-Object { $_.Priority -eq "High" }).Count }
        [PSCustomObject]@{ Property = "Medium Priority"; Value = ($timelineItems | Where-Object { $_.Priority -eq "Medium" }).Count }
        [PSCustomObject]@{ Property = "Normal Priority"; Value = ($timelineItems | Where-Object { $_.Priority -eq "Normal" }).Count }
        [PSCustomObject]@{ Property = "Export Mode"; Value = if ($ExportAll) { "All Items" } else { "Priority Tags: $($PriorityTags -join ', ')" } }
        [PSCustomObject]@{ Property = "Azure DevOps Project"; Value = $ProductionConfig.AdoProjectName }
        [PSCustomObject]@{ Property = "Organization"; Value = $ProductionConfig.AdoOrganizationUrl }
        [PSCustomObject]@{ Property = "Export Script"; Value = "export-critical-timeline.ps1" }
        [PSCustomObject]@{ Property = "Office Timeline Compatible"; Value = "Yes" }
    )
    
    $summaryData | Export-Excel -Path $OutputPath -WorksheetName "Export Summary" -AutoSize -BoldTopRow -Append
    
    Write-Host "`n=== Export Complete ===" -ForegroundColor Green
    Write-Host "Excel file created: $OutputPath" -ForegroundColor Cyan
    Write-Host "Total priority items exported: $($timelineItems.Count)" -ForegroundColor White
    Write-Host "  - Milestones: $(($timelineItems | Where-Object { $_.'Work Item Type' -eq 'Milestone' }).Count)" -ForegroundColor White
    Write-Host "  - Dependencies: $(($timelineItems | Where-Object { $_.'Work Item Type' -eq 'Dependency' }).Count)" -ForegroundColor White
    
    Write-Host "`nPriority Breakdown:" -ForegroundColor Yellow
    foreach ($group in $prioritySummary) {
        Write-Host "  - $($group.Name): $($group.Count) items" -ForegroundColor White
    }
    
    Write-Host "`n=== Office Timeline Import Instructions ===" -ForegroundColor Yellow
    Write-Host "1. Open PowerPoint and go to Office Timeline Expert" -ForegroundColor White
    Write-Host "2. Click 'New' > 'Import Data' > 'Excel'" -ForegroundColor White
    Write-Host "3. Select the generated Excel file: $OutputPath" -ForegroundColor White
    Write-Host "4. Choose the 'Priority Timeline' worksheet" -ForegroundColor White
    Write-Host "5. Map the columns appropriately (should auto-detect)" -ForegroundColor White
    Write-Host "6. Generate your priority timeline visualization" -ForegroundColor White
    
    # Display sample of exported data if debug mode is enabled
    if ($DebugMode -and $timelineItems.Count -gt 0) {
        Write-Host "`n=== Sample Export Data ===" -ForegroundColor Cyan
        $timelineItems | Select-Object -First 5 | Format-Table -AutoSize
    }
    
} catch {
    Write-Host "`nError occurred during export: $($_.Exception.Message)" -ForegroundColor Red
    if ($DebugMode) {
        Write-Host "Full error details:" -ForegroundColor Red
        Write-Host $_ -ForegroundColor Red
        Write-Host "Stack trace:" -ForegroundColor Red
        Write-Host $_.ScriptStackTrace -ForegroundColor Red
    }
    exit 1
} finally {
    Write-Host "`nExport process completed." -ForegroundColor Green
}
