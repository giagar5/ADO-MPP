# Check Tags in Azure DevOps Work Items
# This script helps identify what tags are currently used in your project

param(
    [string]$ConfigPath = "..\config\config.ps1",
    [switch]$DebugMode = $false
)

# Load configuration
if (Test-Path $ConfigPath) {
    . $ConfigPath
} else {
    Write-Error "Configuration file not found: $ConfigPath"
    exit 1
}

Write-Host "=== Azure DevOps Tags Analysis ===" -ForegroundColor Green

# Setup headers for Azure DevOps API
$headers = @{
    'Authorization' = "Basic $([Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$($ProductionConfig.PersonalAccessToken)")))"
    'Content-Type' = 'application/json'
    'User-Agent' = 'PowerShell-TagAnalysis/1.0'
}

# Query for all Milestones and Dependencies with their tags
$allItemsQuery = @"
SELECT [System.Id], [System.Title], [System.WorkItemType], [System.State], [System.Tags]
FROM WorkItems 
WHERE [System.TeamProject] = '$($ProductionConfig.AdoProjectName)' 
AND [System.WorkItemType] IN ('Milestone', 'Dependency')
AND [System.State] <> 'Removed'
AND ([System.AreaPath] UNDER 'Azure-Cloud-Transformation-Program\Workstream D- Data Estate Modernization (E07)' OR [System.Id] = 2449)
ORDER BY [System.WorkItemType], [System.Id]
"@

try {
    Write-Host "Querying all Milestones and Dependencies..." -ForegroundColor Yellow
    
    # Execute WIQL query
    $wiqlUrl = "$($ProductionConfig.AdoOrganizationUrl)/$($ProductionConfig.AdoProjectName)/_apis/wit/wiql?api-version=7.1"
    $wiqlRequest = @{ query = $allItemsQuery }
    $wiqlBody = $wiqlRequest | ConvertTo-Json -Depth 3
    
    $queryResponse = Invoke-RestMethod -Uri $wiqlUrl -Headers $headers -Method "POST" -Body $wiqlBody -TimeoutSec 30
    
    if (-not $queryResponse.workItems -or $queryResponse.workItems.Count -eq 0) {
        Write-Host "No Milestones or Dependencies found in the specified scope." -ForegroundColor Red
        exit 0
    }
    
    Write-Host "Found $($queryResponse.workItems.Count) items. Getting detailed information..." -ForegroundColor Green
    
    # Get detailed work item information
    $workItemIds = $queryResponse.workItems | ForEach-Object { $_.id }
    $batchIds = $workItemIds -join ","
    
    $detailUrl = "$($ProductionConfig.AdoOrganizationUrl)/$($ProductionConfig.AdoProjectName)/_apis/wit/workitems?ids=$batchIds&`$expand=All&api-version=7.1"
    $detailResponse = Invoke-RestMethod -Uri $detailUrl -Headers $headers -TimeoutSec 30
    
    # Analyze tags
    $tagAnalysis = @{}
    $itemsWithTags = @()
    $itemsWithoutTags = @()
    
    foreach ($workItem in $detailResponse.value) {
        $id = $workItem.id
        $title = if ($workItem.fields.'System.Title') { $workItem.fields.'System.Title' } else { "No Title" }
        $type = if ($workItem.fields.'System.WorkItemType') { $workItem.fields.'System.WorkItemType' } else { "Unknown" }
        $state = if ($workItem.fields.'System.State') { $workItem.fields.'System.State' } else { "Unknown" }
        $tags = if ($workItem.fields.'System.Tags') { $workItem.fields.'System.Tags' } else { "" }
        
        $itemInfo = [PSCustomObject]@{
            ID = $id
            Title = $title
            Type = $type
            State = $state
            Tags = $tags
        }
        
        if ($tags) {
            $itemsWithTags += $itemInfo
            # Parse individual tags
            $individualTags = $tags -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
            foreach ($tag in $individualTags) {
                if ($tagAnalysis.ContainsKey($tag)) {
                    $tagAnalysis[$tag]++
                } else {
                    $tagAnalysis[$tag] = 1
                }
            }
        } else {
            $itemsWithoutTags += $itemInfo
        }
    }
    
    # Display results
    Write-Host "`n=== SUMMARY ===" -ForegroundColor Cyan
    Write-Host "Total Milestones and Dependencies: $($detailResponse.value.Count)" -ForegroundColor White
    Write-Host "Items with tags: $($itemsWithTags.Count)" -ForegroundColor Green
    Write-Host "Items without tags: $($itemsWithoutTags.Count)" -ForegroundColor Yellow
    
    if ($tagAnalysis.Count -gt 0) {
        Write-Host "`n=== TAG USAGE ANALYSIS ===" -ForegroundColor Cyan
        $sortedTags = $tagAnalysis.GetEnumerator() | Sort-Object Value -Descending
        foreach ($tag in $sortedTags) {
            $isCritical = $tag.Key.ToLower().Contains("critical")
            $color = if ($isCritical) { "Green" } else { "White" }
            Write-Host "  $($tag.Key): $($tag.Value) items" -ForegroundColor $color
        }
        
        # Check for Critical tag specifically
        $criticalTags = $sortedTags | Where-Object { $_.Key.ToLower().Contains("critical") }
        if ($criticalTags) {
            Write-Host "`n=== CRITICAL TAGS FOUND ===" -ForegroundColor Green
            foreach ($criticalTag in $criticalTags) {
                Write-Host "  '$($criticalTag.Key)' - used in $($criticalTag.Value) items" -ForegroundColor Green
            }
        } else {
            Write-Host "`n=== NO CRITICAL TAGS FOUND ===" -ForegroundColor Red
            Write-Host "None of your items have tags containing 'Critical'" -ForegroundColor Yellow
        }
    } else {
        Write-Host "`nNo tags found in any work items." -ForegroundColor Yellow
    }
    
    # Display items with tags
    if ($itemsWithTags.Count -gt 0) {
        Write-Host "`n=== ITEMS WITH TAGS ===" -ForegroundColor Cyan
        $itemsWithTags | Format-Table -AutoSize
    }
    
    # Display items without tags
    if ($itemsWithoutTags.Count -gt 0) {
        Write-Host "`n=== ITEMS WITHOUT TAGS ===" -ForegroundColor Yellow
        Write-Host "These items could be tagged as 'Critical' if they are important for your timeline:" -ForegroundColor White
        $itemsWithoutTags | Format-Table -AutoSize
    }
    
    # Provide recommendations
    Write-Host "`n=== RECOMMENDATIONS ===" -ForegroundColor Magenta
    if ($tagAnalysis.Count -eq 0) {
        Write-Host "1. No tags are currently used in your Milestones/Dependencies" -ForegroundColor White
        Write-Host "2. Consider adding a 'Critical' tag to important items in Azure DevOps" -ForegroundColor White
        Write-Host "3. Or modify the export script to use different criteria" -ForegroundColor White
    } elseif (-not ($tagAnalysis.Keys | Where-Object { $_.ToLower().Contains("critical") })) {
        Write-Host "1. You have tags but none contain 'Critical'" -ForegroundColor White
        Write-Host "2. Consider adding 'Critical' tag to important items" -ForegroundColor White
        Write-Host "3. Or modify the script to use existing tags like: $($tagAnalysis.Keys -join ', ')" -ForegroundColor White
    } else {
        Write-Host "1. You have Critical tags! The export script should work." -ForegroundColor Green
        Write-Host "2. Make sure the export script is using the correct tag matching logic" -ForegroundColor White
    }
    
} catch {
    Write-Host "Error occurred: $($_.Exception.Message)" -ForegroundColor Red
    if ($DebugMode) {
        Write-Host "Full error: $_" -ForegroundColor Red
    }
    exit 1
}
