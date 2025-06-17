#Requires -Version 5.1

<#
.SYNOPSIS
    ADO2MPP - Azure DevOps to Microsoft Project Bridge

.DESCRIPTION
    A comprehensive PowerShell tool that exports work items from Azure DevOps to Excel format, 
    optimized for Microsoft Project import with full hierarchical structure, task dependencies, 
    and rich metadata integration including direct ADO links and comprehensive field mapping.

.PARAMETER ConfigPath
    Path to the configuration file. If not specified, uses the default config.ps1 in the same directory.

.PARAMETER OutputPath
    Override the output Excel file path specified in configuration.

.PARAMETER AreaPath
    Override the area path filter to export work items from a specific area.

.PARAMETER WorkItemTypes
    Override work item types to export (comma-separated). Default: Epic,Feature,User Story,Task,Bug,Dependency,Milestone

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
    $config.WiqlQuery = "SELECT [System.Id] FROM WorkItems WHERE [System.TeamProject] = '$($config.AdoProjectName)' AND [System.WorkItemType] IN ('Epic', 'Feature', 'User Story', 'Task', 'Bug', 'Dependency', 'Milestone') AND [System.AreaPath] UNDER '$($config.AdoProjectName)\$AreaPath'"
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
    
    # Skip debug messages unless debug logging is enabled
    if ($Level -eq "DEBUG" -and -not $config.EnableDebugLogging) {
        return
    }
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $color = switch ($Level) {
        "ERROR" { "Red" }
        "WARNING" { "Yellow" }  
        "SUCCESS" { "Green" }
        "DEBUG" { "Gray" }
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

function Get-ProgressValue {
    param($Fields)
    
    # Simplified progress calculation - only use state-based logic to avoid null issues
    if (-not $Fields -or -not $Fields.'System.State') {
        return "0%"
    }
    
    $state = $Fields.'System.State'
    switch ($state) {
        'Done' { return "100%" }
        'Closed' { return "100%" }
        'Resolved' { return "100%" }
        'Active' { return "50%" }
        'In Progress' { return "50%" }
        'Committed' { return "25%" }
        'New' { return "0%" }
        'To Do' { return "0%" }
        default { return "0%" }
    }
}

function Get-OutlineLevel {
    param(
        [string]$WorkItemType,
        [hashtable]$WorkItemsById = @{},
        [hashtable]$ChildParentMap = @{},
        [int]$WorkItemId = 0
    )
    
    # If we have relationship information, calculate outline level based on hierarchy
    if ($WorkItemId -gt 0 -and $ChildParentMap.ContainsKey($WorkItemId) -and $WorkItemsById.ContainsKey($WorkItemId)) {
        $level = 1
        $currentId = $WorkItemId
        
        # Traverse up the hierarchy to count levels
        while ($ChildParentMap.ContainsKey($currentId)) {
            $level++
            $currentId = $ChildParentMap[$currentId]
            
            # Prevent infinite loops
            if ($level -gt 10) {
                Write-Log "Warning: Hierarchy depth exceeded for work item $WorkItemId, using type-based level" "WARNING"
                break
            }
        }
        
        return $level
    }
    
    # Fallback to type-based outline levels for items without clear hierarchy
    switch ($WorkItemType) {
        'Epic' { return 1 }
        'Feature' { return 2 }
        'User Story' { return 3 }
        'Task' { return 4 }
        'Bug' { return 4 }
        'Dependency' { return 4 }  # Default level, but can be adjusted based on hierarchy
        'Milestone' { return 4 }   # Default level, but can be adjusted based on hierarchy
        default { return 5 }
    }
}

function Convert-EffortToDuration {
    param($EffortHours)
    
    # Simplified duration calculation to avoid null issues
    if (-not $EffortHours -or $EffortHours -le 0) {
        return 1  # Default to 1 day
    }
    
    # Simple conversion - always return at least 1 day
    $hoursPerDay = 8  # Fixed value to avoid config lookup issues
    $days = [Math]::Ceiling($EffortHours / $hoursPerDay)
    return [Math]::Max(1, $days)
}

function Format-NumberForRegion {
    param($Number, $RegionalSettings = $null)
    
    # Simplified number formatting to avoid null issues
    if (-not $Number -or $Number -eq 0) {
        return "0"
    }
    
    # Simple numeric conversion without complex formatting
    try {
        $numericValue = [double]$Number
        return $numericValue.ToString("0.##")
    } catch {
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
    param(
        [array]$WorkItems,
        [hashtable]$Headers,
        [string]$OrgUrl,
        [string]$ProjectName,
        [array]$Fields
    )

    Write-Log "Ordering work items hierarchically to maintain parent-child relationships..."

    # Create lookup maps
    $workItemsById = @{}
    $parentChildMap = @{}
    $childParentMap = @{}
    $missingParentIds = @()

    # Build lookup maps first
    foreach ($item in $WorkItems) {
        $workItemsById[$item.id] = $item
    }
    
    Write-Log "Building parent-child relationships from work item links..."
    $relationshipCount = 0
    
    # First pass: collect missing parent IDs
    foreach ($item in $WorkItems) {
        if ($item.relations) {
            Write-Log "  Checking $($item.relations.Count) relations for work item $($item.id)" "DEBUG"
            foreach ($relation in $item.relations) {
                Write-Log "    Relation type: $($relation.rel)" "DEBUG"
                if ($relation.rel -eq "System.LinkTypes.Hierarchy-Reverse") {
                    if ($relation.url -match '/(\d+)$') {
                        $parentId = [int]$matches[1]
                        # Track missing parents that are not in our work item set
                        if (-not $workItemsById.ContainsKey($parentId)) {
                            if ($missingParentIds -notcontains $parentId) {
                                $missingParentIds += $parentId
                                Write-Log "  Found missing parent $parentId for work item $($item.id)" "DEBUG"
                            }
                        }
                    }
                }
            }
        }
    }
      # Fetch missing parents if any were found
    if ($missingParentIds.Count -gt 0) {
        Write-Log "Found $($missingParentIds.Count) missing parent work items. Fetching them to maintain hierarchy integrity..." "INFO"
        $missingParents = Get-MissingParents -Headers $Headers -OrgUrl $OrgUrl -ProjectName $ProjectName -MissingParentIds $missingParentIds -Fields $Fields
        
        # Add missing parents to our work items collection and lookup map
        # Only add parents that are not already in our collection to avoid duplicates
        $addedParents = @()
        foreach ($parent in $missingParents) {
            if (-not $workItemsById.ContainsKey($parent.id)) {
                $workItemsById[$parent.id] = $parent
                $addedParents += $parent
                Write-Log "  Added missing parent: $($parent.id) '$($parent.fields.'System.Title')' ($($parent.fields.'System.WorkItemType'))" "INFO"
            } else {
                Write-Log "  Skipped duplicate parent: $($parent.id) '$($parent.fields.'System.Title')' (already in collection)" "DEBUG"
            }
        }
        
        # Update the WorkItems array to include only the newly added missing parents
        if ($addedParents.Count -gt 0) {
            $WorkItems = @($WorkItems) + @($addedParents)
            Write-Log "Total work items after adding $($addedParents.Count) missing parents: $($WorkItems.Count)" "INFO"
        } else {
            Write-Log "No new missing parents to add - all were already in collection" "INFO"
        }
    }
    
    # Second pass: build parent-child relationships with complete item set
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
                            Write-Log "  Skipping child $childId of $($item.id) - not in complete set" "DEBUG"
                        }
                    }
                }
                elseif ($relation.rel -eq "System.LinkTypes.Hierarchy-Reverse") {
                    if ($relation.url -match '/(\d+)$') {
                        $parentId = [int]$matches[1]
                        # Now we should have the parent in our complete work item set
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
                            Write-Log "  Skipping parent $parentId of $($item.id) - still not available" "WARNING"
                        }
                    }
                }
            }
        }
    }

    Write-Log "Found $relationshipCount total hierarchy relationships: $($parentChildMap.Keys.Count) items with children, $($childParentMap.Keys.Count) items with parents"

    # If no hierarchy relationships found, use type-based grouping
    if ($relationshipCount -eq 0) {
        Write-Log "No explicit hierarchy relationships found. Using type-based hierarchical ordering..." "WARNING"        # Group by work item type and sort hierarchically - ensure empty arrays instead of null
        $epics = @($WorkItems | Where-Object { $_.fields.'System.WorkItemType' -eq 'Epic' } | Sort-Object { $_.fields.'System.Title' })
        $features = @($WorkItems | Where-Object { $_.fields.'System.WorkItemType' -eq 'Feature' } | Sort-Object { $_.fields.'System.Title' })
        $userStories = @($WorkItems | Where-Object { $_.fields.'System.WorkItemType' -eq 'User Story' } | Sort-Object { $_.fields.'System.Title' })
        $tasks = @($WorkItems | Where-Object { $_.fields.'System.WorkItemType' -eq 'Task' } | Sort-Object { $_.fields.'System.Title' })
        $bugs = @($WorkItems | Where-Object { $_.fields.'System.WorkItemType' -eq 'Bug' } | Sort-Object { $_.fields.'System.Title' })
        $dependencies = @($WorkItems | Where-Object { $_.fields.'System.WorkItemType' -eq 'Dependency' } | Sort-Object { $_.fields.'System.Title' })
        $milestones = @($WorkItems | Where-Object { $_.fields.'System.WorkItemType' -eq 'Milestone' } | Sort-Object { $_.fields.'System.Title' })
        $others = @($WorkItems | Where-Object { $_.fields.'System.WorkItemType' -notin @('Epic', 'Feature', 'User Story', 'Task', 'Bug', 'Dependency', 'Milestone') } | Sort-Object { $_.fields.'System.Title' })
        
        Write-Log "Type-based grouping: $($epics.Count) Epics, $($features.Count) Features, $($userStories.Count) User Stories, $($tasks.Count) Tasks, $($bugs.Count) Bugs, $($dependencies.Count) Dependencies, $($milestones.Count) Milestones, $($others.Count) Others"
          $orderedWorkItems = @()
        $orderedWorkItems += $epics
        $orderedWorkItems += $features
        $orderedWorkItems += $userStories
        $orderedWorkItems += $tasks
        $orderedWorkItems += $bugs
        $orderedWorkItems += $dependencies
        $orderedWorkItems += $milestones
        $orderedWorkItems += $others
          Write-Log "Type-based hierarchical ordering completed: $($orderedWorkItems.Count) work items ordered by type hierarchy"
        return @{
            OrderedWorkItems = $orderedWorkItems
            WorkItemsById = $workItemsById
            ChildParentMap = $childParentMap
            ParentChildMap = $parentChildMap
        }
    }    function Get-OrderedItemsWithChildren {
        param($Items, $ProcessedItems = @{})
        
        # Ensure we have a proper array, even if empty
        if (-not $Items) {
            return @()
        }
        
        $orderedItems = @()
        $sortedItems = @($Items | Sort-Object { $_.fields.'System.Title' })
        
        foreach ($item in $sortedItems) {            # Skip if we've already processed this item to avoid duplicates
            if ($ProcessedItems.ContainsKey($item.id)) {
                continue
            }
            
            # Mark this item as processed
            $ProcessedItems[$item.id] = $true
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
                    $orderedChildren = @(Get-OrderedItemsWithChildren -Items $children -ProcessedItems $ProcessedItems)
                    $orderedItems += $orderedChildren
                } else {
                    Write-Log "  Item $($item.id) '$($item.fields.'System.Title')' has no children in filtered set" "DEBUG"
                }
            }
        }
        return $orderedItems
    }    # Find root items (items without parents in our dataset)
    $rootItems = @($WorkItems | Where-Object { -not $childParentMap.ContainsKey($_.id) })
    Write-Log "Found $($rootItems.Count) root items (items without parents in filtered set)"
    
    # Handle case where no root items are found
    if ($rootItems.Count -eq 0) {
        Write-Log "No root items found! This could indicate circular references or all items have external parents." "WARNING"
        Write-Log "Falling back to type-based ordering for all work items..."
        
        # Use type-based grouping as fallback
        $epics = @($WorkItems | Where-Object { $_.fields.'System.WorkItemType' -eq 'Epic' } | Sort-Object { $_.fields.'System.Title' })
        $features = @($WorkItems | Where-Object { $_.fields.'System.WorkItemType' -eq 'Feature' } | Sort-Object { $_.fields.'System.Title' })
        $userStories = @($WorkItems | Where-Object { $_.fields.'System.WorkItemType' -eq 'User Story' } | Sort-Object { $_.fields.'System.Title' })
        $tasks = @($WorkItems | Where-Object { $_.fields.'System.WorkItemType' -eq 'Task' } | Sort-Object { $_.fields.'System.Title' })
        $bugs = @($WorkItems | Where-Object { $_.fields.'System.WorkItemType' -eq 'Bug' } | Sort-Object { $_.fields.'System.Title' })
        $dependencies = @($WorkItems | Where-Object { $_.fields.'System.WorkItemType' -eq 'Dependency' } | Sort-Object { $_.fields.'System.Title' })
        $milestones = @($WorkItems | Where-Object { $_.fields.'System.WorkItemType' -eq 'Milestone' } | Sort-Object { $_.fields.'System.Title' })
        $others = @($WorkItems | Where-Object { $_.fields.'System.WorkItemType' -notin @('Epic', 'Feature', 'User Story', 'Task', 'Bug', 'Dependency', 'Milestone') } | Sort-Object { $_.fields.'System.Title' })
        
        $orderedWorkItems = @()
        $orderedWorkItems += $epics
        $orderedWorkItems += $features
        $orderedWorkItems += $userStories
        $orderedWorkItems += $tasks
        $orderedWorkItems += $bugs
        $orderedWorkItems += $dependencies
        $orderedWorkItems += $milestones
        $orderedWorkItems += $others
        
        Write-Log "Fallback type-based hierarchical ordering completed: $($orderedWorkItems.Count) work items ordered by type hierarchy"
        return @{
            OrderedWorkItems = $orderedWorkItems
            WorkItemsById = $workItemsById
            ChildParentMap = $childParentMap
            ParentChildMap = $parentChildMap
        }
    }
    
    # Sort root items by type priority (Epic > Feature > User Story > Task > Bug > Dependency > Milestone), then by title
    $sortedRootItems = @($rootItems | Sort-Object @(
        @{Expression={
            switch ($_.fields.'System.WorkItemType') {
                'Epic' { 1 }
                'Feature' { 2 }
                'User Story' { 3 }
                'Task' { 4 }
                'Bug' { 4 }
                'Dependency' { 5 }
                'Milestone' { 6 }
                default { 7 }
            }
        }; Ascending=$true},
        @{Expression={$_.fields.'System.Title'}; Ascending=$true}
    ))
    
    $orderedWorkItems = @(Get-OrderedItemsWithChildren -Items $sortedRootItems -ProcessedItems @{})
    Write-Log "Hierarchical ordering completed: $($orderedWorkItems.Count) work items ordered maintaining parent-child structure"
    
    # Return both the ordered work items and the hierarchy maps for outline level calculation
    return @{
        OrderedWorkItems = $orderedWorkItems
        WorkItemsById = $workItemsById
        ChildParentMap = $childParentMap
        ParentChildMap = $parentChildMap
    }
}

function Export-ToProjectExcel {
    param(
        [Parameter(Mandatory=$true)] [array]$WorkItems,
        [hashtable]$RelationshipMap,
        [Parameter(Mandatory=$true)][ValidateNotNullOrEmpty()] [string]$OutputPath,
        [Parameter(Mandatory=$true)] [hashtable]$RegionalSettings,
        [hashtable]$ChildParentMap = @{},
        [hashtable]$WorkItemsById = @{},
        [string]$AdoOrganizationUrl = "",
        [string]$AdoProjectName = ""
    )

    # Basic parameter validation
    if (-not $WorkItems -or $WorkItems.Count -eq 0) { Write-Log "ERROR: WorkItems parameter missing" "ERROR"; return $false }
    if ([string]::IsNullOrEmpty($OutputPath))       { Write-Log "ERROR: OutputPath missing" "ERROR"; return $false }
    if (-not $RegionalSettings)                    { Write-Log "ERROR: RegionalSettings missing" "ERROR"; return $false }    # Resolve and normalize path
    try { 
        # Convert relative path to absolute path
        if (-not [System.IO.Path]::IsPathRooted($OutputPath)) {
            $OutputPath = [System.IO.Path]::GetFullPath((Join-Path (Get-Location) $OutputPath))
        }
        # Normalize the path
        $OutputPath = [System.IO.Path]::GetFullPath($OutputPath)
    } catch { 
        Write-Log "Warning: Could not fully resolve output path, using as-is: $OutputPath" "WARNING"
    }
    # Ensure directory
    $dir = Split-Path $OutputPath -Parent
    if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

    # Build Excel data
    $excelData = @()
    $taskId = 1
    $lookup = @{}
    foreach ($item in $WorkItems) {
        $lookup[$item.id] = $taskId
        $taskId++
    }    $taskId = 1
    foreach ($item in $WorkItems) {
        $fields = $item.fields
        
        # Extract name with proper null handling
        $name = if ($fields.'System.Title') { $fields.'System.Title'.ToString() } else { "Untitled Work Item" }
          # Extract resource name with proper null handling - use Owner first, then AssignedTo
        $resourceName = ""
        
        # Check for Owner field first (if available in your ADO setup)
        if ($fields.'Custom.Owner') {
            if ($fields.'Custom.Owner'.displayName) {
                $resourceName = $fields.'Custom.Owner'.displayName.ToString()
            } elseif ($fields.'Custom.Owner'.ToString()) {
                $resourceName = $fields.'Custom.Owner'.ToString()
            }
        }
        # Fallback to AssignedTo if Owner not available
        elseif ($fields.'System.AssignedTo') {
            if ($fields.'System.AssignedTo'.displayName) {
                $resourceName = $fields.'System.AssignedTo'.displayName.ToString()
            } elseif ($fields.'System.AssignedTo'.ToString()) {
                $resourceName = $fields.'System.AssignedTo'.ToString()
            }
        }
        
        # Build predecessors string if RelationshipMap exists
        $predecessorsString = ""
        if ($RelationshipMap -and $RelationshipMap.ContainsKey($item.id)) {
            $predecessorIds = @()
            foreach ($predecessorWorkItemId in $RelationshipMap[$item.id]) {
                if ($lookup.ContainsKey($predecessorWorkItemId)) {
                    $predecessorIds += $lookup[$predecessorWorkItemId]
                }
            }
            if ($predecessorIds.Count -gt 0) {
                $predecessorsString = ($predecessorIds | Sort-Object) -join ";"
            }
        }        # Extract Start and Finish dates with priority logic
        # Start: Use StartDate if available
        $startDate = Format-DateForProject -DateString $fields.'Microsoft.VSTS.Scheduling.StartDate'
        
        # Finish: Use revised due date (TargetDate) if present, otherwise original due date (DueDate)
        $finishDate = ""
        if ($fields.'Microsoft.VSTS.Scheduling.TargetDate') {
            $finishDate = Format-DateForProject -DateString $fields.'Microsoft.VSTS.Scheduling.TargetDate'
        } elseif ($fields.'Microsoft.VSTS.Scheduling.DueDate') {
            $finishDate = Format-DateForProject -DateString $fields.'Microsoft.VSTS.Scheduling.DueDate'
        } elseif ($fields.'Microsoft.VSTS.Scheduling.TargetDate') {
            # Fallback to TargetDate even if DueDate doesn't exist
            $finishDate = Format-DateForProject -DateString $fields.'Microsoft.VSTS.Scheduling.TargetDate'
        }
        $adoUrl = ""
        if ($item.url) {
            # Use the provided URL from the API response
            $adoUrl = $item.url
        } elseif (-not [string]::IsNullOrEmpty($AdoOrganizationUrl) -and -not [string]::IsNullOrEmpty($AdoProjectName)) {
            # Fallback to construct URL from organization and project
            $adoUrl = "$AdoOrganizationUrl/$AdoProjectName/_workitems/edit/$($item.id)"
        } else {
            # Last fallback - just show the work item ID
            $adoUrl = "Work Item ID: $($item.id)"
        }        
        # Create Excel row with native fields plus Text fields for ADO metadata
        $excelData += [PSCustomObject]@{
            'Unique ID'     = $taskId
            'Name'          = $name
            'Outline Level' = (Get-OutlineLevel -WorkItemType $fields.'System.WorkItemType' -WorkItemsById $WorkItemsById -ChildParentMap $ChildParentMap -WorkItemId $item.id)
            '% Complete'    = (Get-ProgressValue -Fields $fields)
            'Start'         = $startDate
            'Finish'        = $finishDate
            'Duration'      = "1"
            'Predecessors'  = $predecessorsString
            'Resource Names'= $resourceName
            'Text1'         = if ($fields.'System.WorkItemType') { $fields.'System.WorkItemType'.ToString() } else { "" }
            'Text2'         = if ($fields.'System.State') { $fields.'System.State'.ToString() } else { "" }
            'Text3'         = $adoUrl
            'Number1'       = $item.id
            'Notes'         = if ($fields.'System.Description') { $fields.'System.Description'.ToString() } else { "" }
        }
        $taskId++
    }    # Export using ImportExcel
    Import-Module ImportExcel -ErrorAction Stop
    
    # Enhanced file handling with retry logic
    $maxRetries = 3
    $retryCount = 0
    $fileRemoved = $false
    
    while ($retryCount -lt $maxRetries -and -not $fileRemoved) {
        try {
            if (Test-Path $OutputPath) {
                # Try to remove the file
                Remove-Item $OutputPath -Force -ErrorAction Stop
                Write-Log "Existing file removed successfully" "DEBUG"
            }
            $fileRemoved = $true
        }
        catch {
            $retryCount++
            Write-Log "Attempt $retryCount to remove existing file failed: $($_.Exception.Message)" "WARNING"
            
            if ($retryCount -lt $maxRetries) {
                Write-Log "Waiting 2 seconds before retry..." "DEBUG"
                Start-Sleep -Seconds 2
                
                # Try to close any Excel processes that might have the file open
                try {
                    $excelProcesses = Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue
                    if ($excelProcesses) {
                        Write-Log "Found $($excelProcesses.Count) Excel processes running. You may need to close Excel manually." "WARNING"
                    }
                }
                catch {
                    # Ignore errors when checking for Excel processes
                }
            }
            else {
                Write-Log "Failed to remove existing file after $maxRetries attempts. The file may be open in Excel." "ERROR"
                Write-Log "Please close the Excel file and try again, or choose a different output path." "ERROR"
                  # Generate alternative filename
                $directory = Split-Path $OutputPath -Parent
                $filenameWithExt = Split-Path $OutputPath -Leaf
                $filename = [System.IO.Path]::GetFileNameWithoutExtension($filenameWithExt)
                $extension = [System.IO.Path]::GetExtension($filenameWithExt)
                $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
                $alternativeOutputPath = Join-Path $directory "$filename`_$timestamp$extension"
                
                Write-Log "Using alternative output path: $alternativeOutputPath" "WARNING"
                $OutputPath = $alternativeOutputPath
                $fileRemoved = $true
            }
        }
    }
    
    # Export to Excel with error handling
    try {
        $excelData | Export-Excel -Path $OutputPath -WorksheetName 'Tasks' -AutoSize -BoldTopRow
        Write-Log "Excel file created successfully: $OutputPath" "SUCCESS"
        return $true
    }
    catch {
        Write-Log "Failed to create Excel file: $($_.Exception.Message)" "ERROR"
        
        # Try CSV export as fallback
        try {
            $csvPath = $OutputPath -replace '\.xlsx$', '.csv'
            $excelData | Export-Csv -Path $csvPath -NoTypeInformation -Delimiter ';'
            Write-Log "Fallback CSV file created: $csvPath" "WARNING"
            return $true
        }
        catch {
            Write-Log "Failed to create CSV fallback: $($_.Exception.Message)" "ERROR"
            return $false
        }
    }
}

function Get-MissingParents {
    param(
        [hashtable]$Headers, 
        [string]$OrgUrl, 
        [string]$ProjectName, 
        [array]$MissingParentIds,
        [array]$Fields
    )
    
    if ($MissingParentIds.Count -eq 0) {
        return @()
    }
    
    Write-Log "Fetching $($MissingParentIds.Count) missing parent work items..."
    
    $missingParents = @()
    
    try {
        # Use batch API to fetch missing parents efficiently
        $batchApiUrl = "$OrgUrl/$ProjectName/_apis/wit/workitemsbatch?api-version=7.1"
        $batchRequest = @{
            ids = $MissingParentIds
            fields = $Fields
        }
        $batchBody = $batchRequest | ConvertTo-Json -Depth 3
        
        $response = Invoke-RestMethod -Uri $batchApiUrl -Method Post -Headers $Headers -Body $batchBody -TimeoutSec $config.ApiTimeout
        
        if ($response.value) {
            $missingParents = $response.value
            Write-Log "Successfully fetched $($missingParents.Count) missing parent work items" "SUCCESS"
            
            # Also fetch relationships for the missing parents if relationship processing is enabled
            if ($config.ProcessRelationships) {
                Write-Log "Fetching relationships for missing parents..."
                for ($i = 0; $i -lt $missingParents.Count; $i++) {
                    try {
                        $workItemApiUrl = "$OrgUrl/$ProjectName/_apis/wit/workitems/$($missingParents[$i].id)?`$expand=Relations&api-version=7.1"
                        $workItemWithRelations = Invoke-RestMethod -Uri $workItemApiUrl -Method Get -Headers $Headers -TimeoutSec $config.ApiTimeout
                        
                        if ($workItemWithRelations.relations) {
                            $missingParents[$i] = $workItemWithRelations
                            Write-Log "Missing parent $($missingParents[$i].id) has $($workItemWithRelations.relations.Count) relations" "DEBUG"
                        }
                    } catch {
                        Write-Log "Failed to get relationships for missing parent $($missingParents[$i].id): $($_.Exception.Message)" "DEBUG"
                    }
                }
            }
        }
    } catch {
        Write-Log "Error fetching missing parents: $($_.Exception.Message)" "ERROR"
    }
    
    return $missingParents
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

# Order work items hierarchically to maintain parent-child relationships
Write-Log "Ordering work items hierarchically for proper Excel export..." "INFO"
$hierarchyResult = Get-HierarchicallyOrderedWorkItems -WorkItems $workItems -Headers $headers -OrgUrl $config.AdoOrganizationUrl -ProjectName $config.AdoProjectName -Fields $config.FieldsToFetch
$orderedWorkItems = $hierarchyResult.OrderedWorkItems
Write-Log "Hierarchical ordering complete: $($orderedWorkItems.Count) work items ordered" "INFO"

# Export to Excel
# DEBUG: Validate parameters before calling Export-ToProjectExcel
Write-Log "=== PRE-EXPORT PARAMETER VALIDATION ===" "DEBUG"
Write-Log "OrderedWorkItems count: $($orderedWorkItems.Count)" "DEBUG"
Write-Log "WorkItemRelationships count: $($workItemRelationships.Count)" "DEBUG"
Write-Log "OutputPath from config: '$($config.OutputExcelPath)'" "DEBUG"
Write-Log "OutputPath type: $($config.OutputExcelPath.GetType().FullName)" "DEBUG"
Write-Log "RegionalSettings: $($regionalSettings | ConvertTo-Json -Depth 2)" "DEBUG"

# Validate critical parameters
if (-not $orderedWorkItems -or $orderedWorkItems.Count -eq 0) {
    Write-Log "ERROR: OrderedWorkItems parameter is null or empty" "ERROR"
    exit 1
}

if (-not $config.OutputExcelPath -or $config.OutputExcelPath -eq "") {
    Write-Log "ERROR: OutputExcelPath is null or empty" "ERROR"
    Write-Log "Config.OutputExcelPath value: '$($config.OutputExcelPath)'" "ERROR"
    exit 1
}

if (-not $regionalSettings) {
    Write-Log "ERROR: RegionalSettings parameter is null" "ERROR"
    exit 1
}

Write-Log "Parameter validation passed. Calling Export-ToProjectExcel..." "DEBUG"
$success = Export-ToProjectExcel -WorkItems $orderedWorkItems -RelationshipMap $workItemRelationships -OutputPath $config.OutputExcelPath -RegionalSettings $regionalSettings -ChildParentMap $hierarchyResult.ChildParentMap -WorkItemsById $hierarchyResult.WorkItemsById -AdoOrganizationUrl $config.AdoOrganizationUrl -AdoProjectName $config.AdoProjectName

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
