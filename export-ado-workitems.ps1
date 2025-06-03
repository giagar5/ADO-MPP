#Requires -Version 5.1

<#
.SYNOPSIS
    Azure DevOps to Microsoft Project Excel Export Script - Production Version

.DESCRIPTION
    This script exports work items from Azure DevOps to Excel format compatible with Microsoft Project import.
    Supports hierarchical work item structures (Epic > Feature > User Story > Task) and task dependencies.

.PARAMETER ConfigPath
    Path to the configuration file. If not specified, uses the default config.ps1 in the same directory.

.PARAMETER OutputPath
    Override the output Excel file path specified in configuration.

.PARAMETER AreaPath
    Override the area path filter to export work items from a specific area.

.PARAMETER WorkItemTypes
    Override work item types to export (comma-separated). Default: Epic,Feature,User Story,Task,Bug

.EXAMPLE
    .\export-ado-workitems.ps1
    Exports work items using default configuration

.EXAMPLE
    .\export-ado-workitems.ps1 -OutputPath "C:\exports\my-project.xlsx"
    Exports to a specific output file

.EXAMPLE
    .\export-ado-workitems.ps1 -AreaPath "MyProject\Team Alpha" -WorkItemTypes "Epic,Feature,User Story"
    Exports specific work item types from a specific area
#>

param(
    [string]$ConfigPath,
    [string]$OutputPath,
    [string]$AreaPath,
    [string]$WorkItemTypes
)

# Get script directory for relative paths
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Load configuration
if (-not $ConfigPath) {
    $ConfigPath = Join-Path $ScriptDir "config.ps1"
}

if (-not (Test-Path $ConfigPath)) {
    Write-Error "Configuration file not found: $ConfigPath"
    Write-Host "Please ensure config.ps1 exists in the same directory as this script."
    exit 1
}

try {
    . $ConfigPath
    $config = $ProductionConfig
} catch {
    Write-Error "Failed to load configuration: $($_.Exception.Message)"
    exit 1
}

# Apply parameter overrides
if ($OutputPath) {
    $config.OutputExcelPath = $OutputPath
}

if ($AreaPath) {
    $config.WiqlQuery = "SELECT [System.Id] FROM WorkItems WHERE [System.TeamProject] = '$($config.AdoProjectName)' AND [System.WorkItemType] IN ('Epic', 'Feature', 'User Story', 'Task', 'Bug') AND [System.AreaPath] UNDER '$($config.AdoProjectName)\$AreaPath'"
}

if ($WorkItemTypes) {
    $types = $WorkItemTypes -split ',' | ForEach-Object { "'$($_.Trim())'" }
    $typeFilter = $types -join ', '
    $config.WiqlQuery = "SELECT [System.Id] FROM WorkItems WHERE [System.TeamProject] = '$($config.AdoProjectName)' AND [System.WorkItemType] IN ($typeFilter)"
}

# =============================================================================
# REGIONAL SETTINGS FUNCTIONS
# =============================================================================

function Get-RegionalSettings {
    param([hashtable]$Config)
    
    Write-Log "Determining regional settings..." "DEBUG"
    
    $settings = @{
        DecimalSeparator = "."
        ListSeparator = ";"
        ThousandsSeparator = ""
    }
    
    switch ($Config.RegionalFormat) {
        "Auto" {
            Write-Log "Auto-detecting regional settings from system..." "DEBUG"
            try {
                $culture = [System.Globalization.CultureInfo]::CurrentCulture
                $systemDecimal = $culture.NumberFormat.NumberDecimalSeparator
                $systemList = $culture.TextInfo.ListSeparator
                
                # For Microsoft Project compatibility, we always use period as decimal
                # but respect system list separator preference
                $settings.DecimalSeparator = "."
                $settings.ListSeparator = if ($systemList -eq ",") { ";" } else { $systemList }
                $settings.ThousandsSeparator = ""
                
                Write-Log "System culture: $($culture.Name)" "DEBUG"
                Write-Log "System decimal separator: $systemDecimal, using: $($settings.DecimalSeparator)" "DEBUG"
                Write-Log "System list separator: $systemList, using: $($settings.ListSeparator)" "DEBUG"
            } catch {
                Write-Log "Failed to detect system settings, using defaults" "WARNING"
            }
        }
        "US" {
            Write-Log "Using US regional format" "DEBUG"
            $settings.DecimalSeparator = "."
            $settings.ListSeparator = ","
            $settings.ThousandsSeparator = ""
        }
        "European" {
            Write-Log "Using European regional format" "DEBUG"
            $settings.DecimalSeparator = "."
            $settings.ListSeparator = ";"
            $settings.ThousandsSeparator = ""
        }
        "Custom" {
            Write-Log "Using custom regional format" "DEBUG"
            $settings.DecimalSeparator = $Config.CustomDecimalSeparator
            $settings.ListSeparator = $Config.CustomListSeparator
            $settings.ThousandsSeparator = $Config.CustomThousandsSeparator
        }
        default {
            Write-Log "Unknown regional format '$($Config.RegionalFormat)', using defaults" "WARNING"
        }
    }
    
    Write-Log "Regional settings - Decimal: '$($settings.DecimalSeparator)', List: '$($settings.ListSeparator)', Thousands: '$($settings.ThousandsSeparator)'" "INFO"
    return $settings
}

# =============================================================================
# CORE FUNCTIONS
# =============================================================================

function Write-Log {
    param(
        [string]$Message, 
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $color = switch ($Level) {
        "ERROR" { "Red" }
        "WARNING" { "Yellow" }  
        "SUCCESS" { "Green" }
        "DEBUG" { if ($config.EnableDebugLogging) { "Gray" } else { return } }
        default { "White" }
    }
    Write-Host "[$timestamp] [$Level] $Message" -ForegroundColor $color
}

function Test-AdoConnection {
    param([hashtable]$Headers, [string]$OrgUrl, [string]$ProjectName)
    
    Write-Log "Testing Azure DevOps connection..."
    
    try {
        $projectApiUrl = "$OrgUrl/_apis/projects/$ProjectName"
        $response = Invoke-RestMethod -Uri $projectApiUrl -Method Get -Headers $Headers -TimeoutSec $config.ApiTimeout
        Write-Log "Successfully connected to project: $($response.name)" "SUCCESS"
        return $true
    } catch {
        Write-Log "Failed to connect to Azure DevOps: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

function Get-WorkItemIds {
    param([hashtable]$Headers, [string]$OrgUrl, [string]$ProjectName, [string]$Query)
    
    Write-Log "Executing WIQL query to get work item IDs..."
    Write-Log "Query: $Query" "DEBUG"
    
    try {
        $wiqlApiUrl = "$OrgUrl/$ProjectName/_apis/wit/wiql?api-version=7.1"
        $queryBody = @{ query = $Query } | ConvertTo-Json
        $response = Invoke-RestMethod -Uri $wiqlApiUrl -Method Post -Headers $Headers -Body $queryBody -TimeoutSec $config.ApiTimeout
        
        $workItemIds = $response.workItems | ForEach-Object { $_.id }
        Write-Log "Found $($workItemIds.Count) work items" "SUCCESS"
        return $workItemIds
    } catch {
        Write-Log "Failed to get work item IDs: $($_.Exception.Message)" "ERROR"
        return @()
    }
}

function Get-WorkItemDetails {
    param([hashtable]$Headers, [string]$OrgUrl, [string]$ProjectName, [array]$WorkItemIds, [array]$Fields, [int]$BatchSize)
    
    Write-Log "Fetching work item details with relationships..."
    
    $allWorkItems = @()
    $totalItems = $WorkItemIds.Count
    
    # First, get basic work item details in batches (faster for large datasets)
    Write-Log "Step 1: Fetching basic work item details in batches of $BatchSize..."
    $totalBatches = [Math]::Ceiling($WorkItemIds.Count / $BatchSize)
    
    for ($i = 0; $i -lt $WorkItemIds.Count; $i += $BatchSize) {
        $batchNum = [Math]::Floor($i / $BatchSize) + 1
        $endIndex = [Math]::Min($i + $BatchSize - 1, $WorkItemIds.Count - 1)
        $batchIds = $WorkItemIds[$i..$endIndex]
        
        Write-Log "Processing batch $batchNum of $totalBatches (IDs: $($batchIds[0]) - $($batchIds[-1]))"
        
        try {
            $batchApiUrl = "$OrgUrl/$ProjectName/_apis/wit/workitemsbatch?api-version=7.1"
            $batchRequest = @{
                ids = $batchIds
                fields = $Fields
            }
            $batchBody = $batchRequest | ConvertTo-Json -Depth 3
            
            $response = Invoke-RestMethod -Uri $batchApiUrl -Method Post -Headers $Headers -Body $batchBody -TimeoutSec $config.ApiTimeout
            
            if ($response.value) {
                $allWorkItems += $response.value
                Write-Log "Batch $batchNum completed: $($response.value.Count) items retrieved"
            }
        } catch {
            Write-Log "Error in batch $batchNum`: $($_.Exception.Message)" "ERROR"
        }
    }
    
    Write-Log "Basic details retrieved for $($allWorkItems.Count) work items" "SUCCESS"
      # Step 2: Get relationships using individual API calls (batch API doesn't support Relations properly)
    if ($config.ProcessRelationships) {
        Write-Log "Step 2: Fetching relationships for work items (this may take a moment)..."
        
        $processedCount = 0
        $itemsWithRelations = 0
        $relationshipBatchSize = if ($config.RelationshipBatchSize) { $config.RelationshipBatchSize } else { 25 }
        
        for ($i = 0; $i -lt $allWorkItems.Count; $i++) {
            $workItem = $allWorkItems[$i]
            $processedCount++
            
            # Progress reporting every batch
            if ($processedCount % $relationshipBatchSize -eq 0 -or $processedCount -eq $allWorkItems.Count) {
                Write-Log "Relationships progress: $processedCount / $($allWorkItems.Count) work items processed"
                # Small delay every batch to avoid overwhelming the API
                Start-Sleep -Milliseconds 500
            }
            
            try {
                # Get individual work item with relationships using the single work item API
                $workItemApiUrl = "$OrgUrl/$ProjectName/_apis/wit/workitems/$($workItem.id)?`$expand=Relations&api-version=7.1"
                $workItemWithRelations = Invoke-RestMethod -Uri $workItemApiUrl -Method Get -Headers $Headers -TimeoutSec $config.ApiTimeout
                
                if ($workItemWithRelations.relations -and $workItemWithRelations.relations.Count -gt 0) {
                    $allWorkItems[$i] = $workItemWithRelations
                    $itemsWithRelations++
                    Write-Log "Work item $($workItem.id) has $($workItemWithRelations.relations.Count) relations" "DEBUG"
                }
            } catch {
                Write-Log "Failed to get relationships for work item $($workItem.id): $($_.Exception.Message)" "DEBUG"
                # Continue processing other items even if one fails
            }
        }
        
        Write-Log "Relationships retrieved: $itemsWithRelations work items have relationship data" "SUCCESS"
    } else {
        Write-Log "Relationship processing disabled in configuration" "INFO"
    }
    
    Write-Log "Total work items with full details: $($allWorkItems.Count)" "SUCCESS"
    return $allWorkItems
}

function Get-WorkItemRelationships {
    param([array]$WorkItems)
    
    Write-Log "Analyzing work item relationships..."
    
    $relationships = @{}
    $hierarchyCount = 0
    $dependencyCount = 0
    $totalItemsWithRelations = 0
    
    foreach ($workItem in $WorkItems) {
        if ($workItem.relations) {
            $totalItemsWithRelations++
            Write-Log "Work item $($workItem.id) has $($workItem.relations.Count) relations" "DEBUG"
            
            foreach ($relation in $workItem.relations) {
                $relType = $relation.rel
                Write-Log "  Relation type: $relType, URL: $($relation.url)" "DEBUG"
                
                # Extract related work item ID from URL
                if ($relation.url -match '/(\d+)$') {
                    $relatedWorkItemId = [int]$matches[1]
                    
                    # Check for dependency relationships (predecessor/successor)
                    if ($relType -eq "System.LinkTypes.Dependency-Forward" -or $relType -eq "Microsoft.VSTS.Common.TestedBy-Forward") {
                        if (-not $relationships.ContainsKey($workItem.id)) {
                            $relationships[$workItem.id] = @()
                        }
                        $relationships[$workItem.id] += $relatedWorkItemId
                        $dependencyCount++
                        Write-Log "  Dependency: $($workItem.id) depends on $relatedWorkItemId" "DEBUG"
                    }
                    elseif ($relType -eq "System.LinkTypes.Hierarchy-Forward") {
                        $hierarchyCount++
                        Write-Log "  Hierarchy Forward: $($workItem.id) is parent of $relatedWorkItemId" "DEBUG"
                    }
                    elseif ($relType -eq "System.LinkTypes.Hierarchy-Reverse") {
                        $hierarchyCount++
                        Write-Log "  Hierarchy Reverse: $($workItem.id) is child of $relatedWorkItemId" "DEBUG"
                    }
                } else {
                    Write-Log "  Could not extract work item ID from URL: $($relation.url)" "DEBUG"
                }
            }
        }
    }
    
    Write-Log "Total items with relations: $totalItemsWithRelations" "DEBUG"
    Write-Log "Found $hierarchyCount hierarchy relationships and $dependencyCount dependency relationships" "SUCCESS"
    return $relationships
}

function Get-OutlineLevel {
    param([string]$WorkItemType)
    
    switch ($WorkItemType) {
        'Epic' { return 1 }
        'Feature' { return 2 }
        'User Story' { return 3 }
        'Task' { return 4 }
        'Bug' { return 4 }
        default { return 5 }
    }
}

function Convert-EffortToDuration {
    param($EffortHours)
    
    # Handle null or invalid values - more robust null checking
    if ($null -eq $EffortHours -or $EffortHours -eq "" -or $EffortHours -le 0) {
        return 1  # Default to 1 day
    }
    
    # Try to convert to double if it's a string
    $effortValue = 0
    if ([double]::TryParse($EffortHours, [ref]$effortValue)) {
        if ($effortValue -le 0) {
            return 1
        }
        $days = [Math]::Ceiling($effortValue / $config.HoursPerDay)
        return [Math]::Max(1, $days)
    } else {
        return 1  # Default if conversion fails
    }
}

function Format-NumberForRegion {
    param($Number, $RegionalSettings = $null)
    
    # Handle null, empty, or invalid values - more robust checking
    if ($null -eq $Number -or $Number -eq "" -or $Number -eq 0) {
        return "0"
    }
    
    # Try to convert to double if it's not already a number
    $numericValue = 0
    if ([double]::TryParse($Number, [ref]$numericValue)) {
        # Use configured decimal separator (default to period for international compatibility)
        $decimalSep = if ($RegionalSettings -and $RegionalSettings.DecimalSeparator) { 
            $RegionalSettings.DecimalSeparator 
        } else { 
            "." 
        }
        
        # Use configured thousands separator (default to none)
        $thousandsSep = if ($RegionalSettings -and $RegionalSettings.ThousandsSeparator) { 
            $RegionalSettings.ThousandsSeparator 
        } else { 
            "" 
        }
        
        # Create custom number format
        if ($thousandsSep -eq "") {
            # No thousands separator
            $formatString = "0" + $decimalSep + "##"
        } else {
            # With thousands separator
            $formatString = "#" + $thousandsSep + "##0" + $decimalSep + "##"
        }
        
        # Use invariant culture for consistent formatting, then replace separators
        $formatted = $numericValue.ToString("0.##", [System.Globalization.CultureInfo]::InvariantCulture)
        
        # Replace decimal separator if different from period
        if ($decimalSep -ne ".") {
            $formatted = $formatted.Replace(".", $decimalSep)
        }
        
        return $formatted
    } else {
        return "0"  # Default if conversion fails
    }
}

function Export-CsvWithSemicolon {
    param(
        [Parameter(ValueFromPipeline)]
        [PSObject[]]$InputObject,
        [string]$Path,
        [string]$Encoding = "UTF8",
        [string]$Delimiter = ";"
    )
    
    begin {
        $allObjects = @()
    }
    
    process {
        $allObjects += $InputObject
    }
    
    end {
        if ($allObjects.Count -eq 0) { return }
        
        # Get headers from first object
        $headers = $allObjects[0].PSObject.Properties.Name
        
        # Create CSV content with specified delimiter
        $csvContent = @()
        
        # Add header line
        $csvContent += $headers -join $Delimiter
        
        # Add data lines
        foreach ($obj in $allObjects) {
            $values = @()
            foreach ($header in $headers) {
                $value = $obj.$header
                if ($null -eq $value) {
                    $value = ""
                }
                
                # Handle values containing delimiter or quotes by wrapping in quotes
                $valueStr = $value.ToString()
                if ($valueStr -match "[$Delimiter`"]") {
                    $valueStr = '"' + $valueStr.Replace('"', '""') + '"'
                }
                
                $values += $valueStr
            }
            $csvContent += $values -join $Delimiter
        }
        
        # Write to file with UTF8 encoding
        $csvContent | Out-File -FilePath $Path -Encoding $Encoding
        Write-Log "Created CSV file with '$Delimiter' delimiter: $Path" "SUCCESS"
    }
}

function Format-DateForProject {
    param([string]$DateString)
    
    if ([string]::IsNullOrEmpty($DateString)) {
        return ""
    }
    
    try {
        $date = [DateTime]::Parse($DateString)
        return $date.ToString("M/d/yyyy")
    } catch {
        Write-Log "Could not parse date: $DateString" "WARNING"
        return ""
    }
}

function Get-HierarchicallyOrderedWorkItems {
    param([array]$WorkItems)

    Write-Log "Ordering work items hierarchically to maintain parent-child relationships..."

    # Create lookup maps
    $workItemsById = @{}
    $parentChildMap = @{}
    $childParentMap = @{}

    # Build lookup maps first
    foreach ($item in $WorkItems) {
        $workItemsById[$item.id] = $item
    }
    
    Write-Log "Building parent-child relationships from work item links..."
    $relationshipCount = 0
    
    # Build parent-child relationships from hierarchy links
    foreach ($item in $WorkItems) {
        if ($item.relations) {
            Write-Log "  Checking $($item.relations.Count) relations for work item $($item.id)" "DEBUG"
            foreach ($relation in $item.relations) {
                Write-Log "    Relation type: $($relation.rel)" "DEBUG"
                if ($relation.rel -eq "System.LinkTypes.Hierarchy-Forward") {
                    if ($relation.url -match '/(\d+)$') {
                        $childId = [int]$matches[1]
                        # Only include relationships where the child is also in our work item set
                        if ($workItemsById.ContainsKey($childId)) {
                            if (-not $parentChildMap.ContainsKey($item.id)) {
                                $parentChildMap[$item.id] = @()
                            }
                            $parentChildMap[$item.id] += $childId
                            $childParentMap[$childId] = $item.id
                            $relationshipCount++
                            Write-Log "  Found hierarchy: $($item.id) → $childId" "DEBUG"
                        } else {
                            Write-Log "  Skipping child $childId of $($item.id) - not in filtered set" "DEBUG"
                        }
                    }
                }
                elseif ($relation.rel -eq "System.LinkTypes.Hierarchy-Reverse") {
                    if ($relation.url -match '/(\d+)$') {
                        $parentId = [int]$matches[1]
                        # Only include relationships where the parent is also in our work item set
                        if ($workItemsById.ContainsKey($parentId)) {
                            if (-not $parentChildMap.ContainsKey($parentId)) {
                                $parentChildMap[$parentId] = @()
                            }
                            if ($parentChildMap[$parentId] -notcontains $item.id) {
                                $parentChildMap[$parentId] += $item.id
                            }
                            $childParentMap[$item.id] = $parentId
                            $relationshipCount++
                            Write-Log "  Found reverse hierarchy: $parentId → $($item.id)" "DEBUG"
                        } else {
                            Write-Log "  Skipping parent $parentId of $($item.id) - not in filtered set" "DEBUG"
                        }
                    }
                }
            }
        }
    }

    Write-Log "Found $relationshipCount total hierarchy relationships: $($parentChildMap.Keys.Count) items with children, $($childParentMap.Keys.Count) items with parents"

    # If no hierarchy relationships found, use type-based grouping
    if ($relationshipCount -eq 0) {
        Write-Log "No explicit hierarchy relationships found. Using type-based hierarchical ordering..." "WARNING"
        
        # Group by work item type and sort hierarchically
        $epics = $WorkItems | Where-Object { $_.fields.'System.WorkItemType' -eq 'Epic' } | Sort-Object { $_.fields.'System.Title' }
        $features = $WorkItems | Where-Object { $_.fields.'System.WorkItemType' -eq 'Feature' } | Sort-Object { $_.fields.'System.Title' }
        $userStories = $WorkItems | Where-Object { $_.fields.'System.WorkItemType' -eq 'User Story' } | Sort-Object { $_.fields.'System.Title' }
        $others = $WorkItems | Where-Object { $_.fields.'System.WorkItemType' -notin @('Epic', 'Feature', 'User Story') } | Sort-Object { $_.fields.'System.Title' }
        
        Write-Log "Type-based grouping: $($epics.Count) Epics, $($features.Count) Features, $($userStories.Count) User Stories, $($others.Count) Others"
        
        $orderedWorkItems = @()
        $orderedWorkItems += $epics
        $orderedWorkItems += $features
        $orderedWorkItems += $userStories
        $orderedWorkItems += $others
        
        Write-Log "Type-based hierarchical ordering completed: $($orderedWorkItems.Count) work items ordered by type hierarchy"
        return $orderedWorkItems
    }

    function Get-OrderedItemsWithChildren {
        param($Items)
        $orderedItems = @()
        $sortedItems = $Items | Sort-Object { $_.fields.'System.Title' }
        foreach ($item in $sortedItems) {
            $orderedItems += $item
            if ($parentChildMap.ContainsKey($item.id)) {
                $childIds = $parentChildMap[$item.id]
                $children = @()
                foreach ($childId in $childIds) {
                    if ($workItemsById.ContainsKey($childId)) {
                        $children += $workItemsById[$childId]
                    }
                }
                if ($children.Count -gt 0) {
                    Write-Log "  Item $($item.id) '$($item.fields.'System.Title')' has $($children.Count) children"
                    $orderedChildren = Get-OrderedItemsWithChildren -Items $children
                    $orderedItems += $orderedChildren
                } else {
                    Write-Log "  Item $($item.id) '$($item.fields.'System.Title')' has no children in filtered set" "DEBUG"
                }
            }
        }
        return $orderedItems
    }

    # Find root items (items without parents in our dataset)
    $rootItems = $WorkItems | Where-Object { -not $childParentMap.ContainsKey($_.id) }
    Write-Log "Found $($rootItems.Count) root items (items without parents in filtered set)"

    # Sort root items by type priority (Epic > Feature > User Story), then by title
    $sortedRootItems = $rootItems | Sort-Object @(
        @{Expression={
            switch ($_.fields.'System.WorkItemType') {
                'Epic' { 1 }
                'Feature' { 2 }
                'User Story' { 3 }
                default { 4 }
            }
        }; Ascending=$true},
        @{Expression={$_.fields.'System.Title'}; Ascending=$true}
    )

    $orderedWorkItems = Get-OrderedItemsWithChildren -Items $sortedRootItems
    Write-Log "Hierarchical ordering completed: $($orderedWorkItems.Count) work items ordered maintaining parent-child structure"
    return $orderedWorkItems
}

function Export-ToProjectExcel {
    param([array]$WorkItems, [hashtable]$RelationshipMap, [string]$OutputPath, [hashtable]$RegionalSettings)
    
    Write-Log "Creating Microsoft Project compatible Excel file: $OutputPath"
    Write-Log "Using regional settings - Decimal: '$($RegionalSettings.DecimalSeparator)', List: '$($RegionalSettings.ListSeparator)'" "DEBUG"
    
    try {
        # Remove duplicates
        $uniqueWorkItems = @{}
        $deduplicatedWorkItems = @()
        
        foreach ($item in $WorkItems) {
            if (-not $uniqueWorkItems.ContainsKey($item.id)) {
                $uniqueWorkItems[$item.id] = $true
                $deduplicatedWorkItems += $item
            }
        }
        
        # Sort hierarchically with proper parent-child relationship
        $sortedWorkItems = Get-HierarchicallyOrderedWorkItems -WorkItems $deduplicatedWorkItems
        
        Write-Log "Creating Excel data for $($sortedWorkItems.Count) work items..."
        
        $excelData = @()
        $taskId = 1
        
        # Create lookup for work item IDs to task IDs
        $workItemToTaskId = @{}
        foreach ($workItem in $sortedWorkItems) {
            $workItemToTaskId[$workItem.id] = $taskId
            $taskId++
        }
        
        $taskId = 1
        foreach ($workItem in $sortedWorkItems) {
            $fields = $workItem.fields
            $workItemType = $fields.'System.WorkItemType'
            $workItemId = $workItem.id
            $outlineLevel = Get-OutlineLevel -WorkItemType $workItemType
            
            # Calculate effort-based duration with safe null handling
            $effort = $fields.'Microsoft.VSTS.Scheduling.OriginalEstimate'
            if (-not $effort -or $null -eq $effort) {
                $effort = $fields.'Microsoft.VSTS.Scheduling.RemainingWork'
            }            if (-not $effort -or $null -eq $effort) { 
                $effort = 8  # Default 8 hours
            }
            
            # Safe priority handling
            $priorityValue = $fields.'Microsoft.VSTS.Common.Priority'
            $priorityFormatted = if ($priorityValue -and $null -ne $priorityValue) { 
                Format-NumberForRegion -Number $priorityValue -RegionalSettings $RegionalSettings
            } else { 
                "500" 
            }
              # Build predecessors string with proper number formatting (no thousands separators)
            $predecessorsString = ""
            if ($RelationshipMap.ContainsKey($workItemId)) {
                $predecessorIds = @()
                foreach ($predecessorWorkItemId in $RelationshipMap[$workItemId]) {
                    if ($workItemToTaskId.ContainsKey($predecessorWorkItemId)) {
                        $predecessorIds += $workItemToTaskId[$predecessorWorkItemId]
                    }
                }
                
                if ($predecessorIds.Count -gt 0) {
                    # Format each ID without thousands separators and join with configured list separator
                    $formattedIds = $predecessorIds | Sort-Object | ForEach-Object { Format-NumberForRegion -Number $_ -RegionalSettings $RegionalSettings }
                    $predecessorsString = $formattedIds -join $RegionalSettings.ListSeparator
                }
            }
            
            # Generate Azure DevOps URLs
            $workItemDirectUrl = "$($config.AdoOrganizationUrl)/$($config.AdoProjectName)/_workitems/edit/$workItemId"
            $boardUrl = "$($config.AdoOrganizationUrl)/$($config.AdoProjectName)/_boards/board/t/$($config.AdoProjectName)%20Team/Stories"
            $backlogUrl = "$($config.AdoOrganizationUrl)/$($config.AdoProjectName)/_backlogs/backlog/$($config.AdoProjectName)%20Team/Stories"
            
            # Use Microsoft Project standard field names for proper import
            $excelRow = [PSCustomObject]@{
                "Unique ID" = $taskId
                "Name" = if ($fields.'System.Title') { $fields.'System.Title' } else { "Untitled" }
                "Duration" = Format-NumberForRegion -Number (Convert-EffortToDuration -EffortHours $effort) -RegionalSettings $RegionalSettings
                "Start" = Format-DateForProject -DateString $fields.'Microsoft.VSTS.Scheduling.StartDate'
                "Finish" = Format-DateForProject -DateString $fields.'Microsoft.VSTS.Scheduling.TargetDate'                "Predecessors" = $predecessorsString
                "Resource Names" = if ($fields.'System.AssignedTo') { $fields.'System.AssignedTo'.displayName } else { "" }
                "Outline Level" = $outlineLevel
                "Work" = "${effort}h"
                "Priority" = $priorityFormatted
                "% Complete" = if ($fields.'System.State' -eq 'Done' -or $fields.'System.State' -eq 'Closed') { "100%" } else { "0%" }
                "Task Mode" = "Auto Scheduled"
                "WBS" = Format-NumberForRegion -Number $taskId -RegionalSettings $RegionalSettings
                "ADO ID" = Format-NumberForRegion -Number $workItemId -RegionalSettings $RegionalSettings
                "Work Item Type" = if ($workItemType) { $workItemType } else { "Unknown" }
                "Text3" = if ($fields.'System.State') { $fields.'System.State' } else { "" }
                "Text4" = if ($fields.'System.AreaPath') { $fields.'System.AreaPath' } else { "" }
                "Text5" = Format-DateForProject -DateString $fields.'System.CreatedDate'
                "ADO Link" = $workItemDirectUrl
                "Text7" = $boardUrl
                "Text8" = $backlogUrl
                "Text9" = if ($fields.'System.Tags') { $fields.'System.Tags' } else { "" }
                "Text10" = if ($priorityValue) { "Priority: $priorityValue" } else { "" }
            }
            
            $excelData += $excelRow
            $taskId++
        }
        
        # Export to Excel - Create clean simplified version only for Microsoft Project import
        try {
            Import-Module ImportExcel -ErrorAction Stop
            
            # Remove any existing file to ensure clean start
            if (Test-Path $OutputPath) {
                Remove-Item $OutputPath -Force
                Write-Log "Removed existing file to ensure clean export" "DEBUG"
            }            # Create simplified Excel file with essential fields for Microsoft Project
            $simplifiedData = $excelData | Select-Object "Unique ID", "Name", "Duration", "Start", "Finish", "Predecessors", "Resource Names", "Outline Level", "ADO ID", "Work Item Type", "ADO Link"
            
            # Ensure no null values that could cause Export-Excel to fail
            $cleanedData = $simplifiedData | ForEach-Object {
                $row = $_
                $cleanRow = [PSCustomObject]@{}
                foreach ($prop in $row.PSObject.Properties) {
                    $value = $prop.Value
                    if ($null -eq $value -or $value -eq "") {
                        $cleanRow | Add-Member -NotePropertyName $prop.Name -NotePropertyValue ""
                    } else {
                        $cleanRow | Add-Member -NotePropertyName $prop.Name -NotePropertyValue $value.ToString()
                    }
                }
                $cleanRow
            }
              # Export to single clean worksheet
            $cleanedData | Export-Excel -Path $OutputPath -WorksheetName "Tasks" -AutoSize -BoldTopRow
            
            Write-Log "Successfully created Excel file: $OutputPath" "SUCCESS"
        } catch {
            # Fallback to CSV with semicolon delimiter for European regional settings
            $csvPath = $OutputPath -replace '\.xlsx$', '.csv'            # Use the same cleaned data for CSV export
            if (-not $cleanedData) {
                $simplifiedData = $excelData | Select-Object "Unique ID", "Name", "Duration", "Start", "Finish", "Predecessors", "Resource Names", "Outline Level", "ADO ID", "Work Item Type", "ADO Link"
                
                # Clean data for CSV as well
                $cleanedData = $simplifiedData | ForEach-Object {
                    $row = $_
                    $cleanRow = [PSCustomObject]@{}
                    foreach ($prop in $row.PSObject.Properties) {
                        $value = $prop.Value
                        if ($null -eq $value -or $value -eq "") {
                            $cleanRow | Add-Member -NotePropertyName $prop.Name -NotePropertyValue ""
                        } else {
                            $cleanRow | Add-Member -NotePropertyName $prop.Name -NotePropertyValue $value.ToString()
                        }
                    }
                    $cleanRow
                }
            }
              # Use custom delimiter-based CSV export for regional compatibility
            $cleanedData | Export-CsvWithSemicolon -Path $csvPath -Encoding UTF8 -Delimiter $RegionalSettings.ListSeparator
            Write-Log "ImportExcel module not available, created CSV file with '$($RegionalSettings.ListSeparator)' delimiter: $csvPath" "WARNING"
            Write-Log "CSV uses '$($RegionalSettings.ListSeparator)' as delimiter for regional settings compatibility" "INFO"
        }
        
        # Display summary
        Write-Log "=== EXPORT SUMMARY ===" "SUCCESS"
        Write-Log "Total work items exported: $($excelData.Count)" "SUCCESS"
        
        $relationshipsApplied = ($excelData | Where-Object { -not [string]::IsNullOrEmpty($_.Predecessors) }).Count
        Write-Log "Work items with predecessor relationships: $relationshipsApplied" "SUCCESS"
        
        $typeDistribution = $excelData | Group-Object "Work Item Type" | Sort-Object Count -Descending
        Write-Log "Work item type distribution:"
        foreach ($type in $typeDistribution) {
            Write-Log "  $($type.Name): $($type.Count)" "INFO"
        }        # Create field mapping guide
        $mappingGuidePath = Join-Path (Split-Path $OutputPath -Parent) "MSProject_Field_Mapping_Guide.txt"
        
        # Ensure OutputPath is not null for the mapping guide
        $outputFileName = if ($OutputPath) { $OutputPath -replace '.*\\', '' } else { 'AzureDevOpsExport_ProjectImport.xlsx' }
          $mappingGuide = @"
MICROSOFT PROJECT IMPORT - FIELD MAPPING GUIDE
==============================================

REGIONAL SETTINGS:
- Format: $($config.RegionalFormat)
- Decimal Separator: $($RegionalSettings.DecimalSeparator)
- List Separator: $($RegionalSettings.ListSeparator)
- Thousands Separator: $($RegionalSettings.ThousandsSeparator)

STEP 1: Open Microsoft Project
STEP 2: Go to File → Open
STEP 3: Select: $outputFileName
STEP 4: Choose 'Tasks' worksheet (for Excel) or specify '$($RegionalSettings.ListSeparator)' delimiter (for CSV)
STEP 5: In Import Wizard, configure TASK MAPPING:

REQUIRED MAPPINGS (Essential - these MUST be mapped):
- Excel Column 'Unique ID' → Project Field 'Unique ID'
- Excel Column 'Name' → Project Field 'Name'  
- Excel Column 'Outline Level' → Project Field 'Outline Level'

RECOMMENDED MAPPINGS (Important for scheduling):
- Excel Column 'Duration' → Project Field 'Duration'
- Excel Column 'Start' → Project Field 'Start'
- Excel Column 'Finish' → Project Field 'Finish'
- Excel Column 'Predecessors' → Project Field 'Predecessors'
- Excel Column 'Resource Names' → Project Field 'Resource Names'

AZURE DEVOPS INTEGRATION FIELDS:
- Excel Column 'ADO ID' → Project Field 'Number1' (Azure DevOps Work Item ID)
- Excel Column 'Work Item Type' → Project Field 'Text1' (Work Item Type)
- Excel Column 'ADO Link' → Project Field 'Text6' (ADO Work Item Direct URL)

NUMBER FORMATTING:
- All numeric fields use '$($RegionalSettings.DecimalSeparator)' as decimal separator
- No thousands separators for clean import
- Predecessor relationships use '$($RegionalSettings.ListSeparator)' as delimiter
- This ensures compatibility with your regional settings

CSV FALLBACK (if Excel module unavailable):
- CSV files use '$($RegionalSettings.ListSeparator)' as field delimiter
- Compatible with your system's regional settings
- Numbers formatted with '$($RegionalSettings.DecimalSeparator)' decimal separator
- Import: File → Open → Select CSV → Specify '$($RegionalSettings.ListSeparator)' as delimiter

CLEAN EXPORT:
This export creates only one clean file with essential fields optimized for Microsoft Project import.
Only Epic, Feature, and User Story work items are included for cleaner hierarchy.

TROUBLESHOOTING:
- If "Map does not map any fields" error: Ensure Unique ID, Name, and Outline Level are mapped
- If hierarchy is lost: Verify Outline Level field is correctly mapped
- If dates don't import: Check date format in Start/Finish columns
- If resources missing: Verify Resource Names field mapping
- If CSV import issues: Ensure '$($RegionalSettings.ListSeparator)' is specified as delimiter
- Select 'Tasks' worksheet when the Import Wizard prompts for worksheet selection

Generated: $((Get-Date).ToString())
"@
        $mappingGuide | Out-File -FilePath $mappingGuidePath -Encoding UTF8
        Write-Log "Created mapping guide: $mappingGuidePath" "SUCCESS"
        
        return $true
        
    } catch {
        Write-Log "Error creating Excel file: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

# =============================================================================
# MAIN EXECUTION
# =============================================================================

Write-Log "=== Azure DevOps to Microsoft Project Excel Export Started ===" "SUCCESS"
Write-Log "Organization: $($config.AdoOrganizationUrl)"
Write-Log "Project: $($config.AdoProjectName)"
Write-Log "Output: $($config.OutputExcelPath)"

# Initialize regional settings
$regionalSettings = Get-RegionalSettings -Config $config
Write-Log "Regional format: $($config.RegionalFormat)" "INFO"

# Ensure output directory exists
$outputDir = Split-Path $config.OutputExcelPath -Parent
if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
    Write-Log "Created output directory: $outputDir"
}

# Setup authentication
$base64Auth = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$($config.PersonalAccessToken)"))
$headers = @{
    Authorization = "Basic $base64Auth"
    "Content-Type" = "application/json"
}

# Test connection
if (-not (Test-AdoConnection -Headers $headers -OrgUrl $config.AdoOrganizationUrl -ProjectName $config.AdoProjectName)) {
    Write-Log "Connection test failed. Exiting." "ERROR"
    exit 1
}

# Get work item IDs
$workItemIds = Get-WorkItemIds -Headers $headers -OrgUrl $config.AdoOrganizationUrl -ProjectName $config.AdoProjectName -Query $config.WiqlQuery

if ($workItemIds.Count -eq 0) {
    Write-Log "No work items found. Exiting." "WARNING"
    exit 0
}

# Apply test mode limit if configured
if ($config.TestModeLimit -and $config.TestModeLimit -gt 0 -and $workItemIds.Count -gt $config.TestModeLimit) {
    Write-Log "TEST MODE: Limiting to first $($config.TestModeLimit) work items for testing" "WARNING"
    $workItemIds = $workItemIds | Select-Object -First $config.TestModeLimit
}

# Get work item details
$workItems = Get-WorkItemDetails -Headers $headers -OrgUrl $config.AdoOrganizationUrl -ProjectName $config.AdoProjectName -WorkItemIds $workItemIds -Fields $config.FieldsToFetch -BatchSize $config.BatchSize

if ($workItems.Count -eq 0) {
    Write-Log "No work item details retrieved. Exiting." "WARNING"
    exit 0
}

# Get relationships if enabled
$workItemRelationships = @{}
if ($config.ProcessRelationships) {
    $workItemRelationships = Get-WorkItemRelationships -WorkItems $workItems
}

# Export to Excel
$success = Export-ToProjectExcel -WorkItems $workItems -RelationshipMap $workItemRelationships -OutputPath $config.OutputExcelPath -RegionalSettings $regionalSettings

if ($success) {
    Write-Log "=== EXPORT COMPLETED SUCCESSFULLY ===" "SUCCESS"
    Write-Log "Excel file created: $($config.OutputExcelPath)" "SUCCESS"
    Write-Log ""
    Write-Log "NEXT STEPS:" "INFO"
    Write-Log "1. Open Microsoft Project" "INFO"
    Write-Log "2. Go to File > Open" "INFO"
    Write-Log "3. Select the Excel file: $($config.OutputExcelPath)" "INFO"
    Write-Log "4. Choose 'Tasks' worksheet when prompted" "INFO"
    Write-Log "5. Follow the Import Wizard to map columns" "INFO"
    Write-Log "6. Refer to the mapping guide: MSProject_Field_Mapping_Guide.txt" "INFO"
} else {
    Write-Log "Export failed!" "ERROR"
    exit 1
}

Write-Log "Script completed successfully." "SUCCESS"
