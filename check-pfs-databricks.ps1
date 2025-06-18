# Diagnostic script to check PFS for databricks epic finish date issue
# Load configuration
. .\config.ps1

Write-Host "=== Checking PFS for databricks epic finish date issue ===" -ForegroundColor Green

# Use the ProductionConfig object from config.ps1
$config = $ProductionConfig

# Setup headers for Azure DevOps API
$headers = @{
    'Authorization' = "Basic $([Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$($config.PersonalAccessToken)")))"
    'Content-Type' = 'application/json'
}

# Search for the specific "PFS for Databricks" work item (ID 12421)
$searchQuery = @"
SELECT [System.Id], [System.Title], [System.WorkItemType], [System.State]
FROM WorkItems 
WHERE [System.TeamProject] = '$($config.AdoProjectName)' 
AND [System.Id] = 12421
"@

try {
    Write-Host "Searching for work items with 'databricks' in title..." -ForegroundColor Yellow
    
    # Debug URL construction
    $orgUrl = $config.AdoOrganizationUrl
    $projectName = $config.AdoProjectName
    Write-Host "Organization URL: '$orgUrl'" -ForegroundColor Cyan
    Write-Host "Project Name: '$projectName'" -ForegroundColor Cyan
    
    $wiqlUrl = "$orgUrl/$projectName/_apis/wit/wiql?api-version=7.1"
    Write-Host "Full WIQL URL: '$wiqlUrl'" -ForegroundColor Cyan
    
    $wiqlRequest = @{ query = $searchQuery }
    $wiqlBody = $wiqlRequest | ConvertTo-Json -Depth 3
    
    Write-Host "Making API request..." -ForegroundColor Yellow
    $wiqlResponse = Invoke-RestMethod -Uri $wiqlUrl -Method Post -Headers $headers -Body $wiqlBody
    
    if ($wiqlResponse.workItems -and $wiqlResponse.workItems.Count -gt 0) {
        Write-Host "Found $($wiqlResponse.workItems.Count) work items with 'databricks' in title" -ForegroundColor Green
        
        # Get detailed information for each work item
        $workItemIds = $wiqlResponse.workItems | ForEach-Object { $_.id }
          foreach ($id in $workItemIds) {
            Write-Host "`n--- Work Item $id ---" -ForegroundColor Cyan
            # Get detailed work item information including all date fields
            $detailUrl = "$orgUrl/$projectName/_apis/wit/workitems/$id" + "?`$expand=All&api-version=7.1"
            
            $workItem = Invoke-RestMethod -Uri $detailUrl -Method Get -Headers $headers
            
            Write-Host "Title: $($workItem.fields.'System.Title')" -ForegroundColor White
            Write-Host "Type: $($workItem.fields.'System.WorkItemType')" -ForegroundColor White
            Write-Host "State: $($workItem.fields.'System.State')" -ForegroundColor White
            
            # Check all relevant date fields
            Write-Host "`nDate Fields:" -ForegroundColor Yellow
            
            $dateFields = @(
                'Microsoft.VSTS.Scheduling.StartDate',
                'Microsoft.VSTS.Scheduling.FinishDate',
                'Microsoft.VSTS.Scheduling.TargetDate',
                'Microsoft.VSTS.Scheduling.DueDate',
                'Custom.OriginalDueDate',
                'Custom.RevisedDueDate'
            )
            
            foreach ($field in $dateFields) {
                $value = $workItem.fields.$field
                if ($value) {
                    Write-Host "  $field`: $value" -ForegroundColor Green
                } else {
                    Write-Host "  $field`: [NOT SET]" -ForegroundColor Red
                }
            }
            
            # Show which field would be used for finish date based on our script logic
            Write-Host "`nFinish Date Logic (from export script):" -ForegroundColor Yellow
            $finishDate = ""
            if ($workItem.fields.'Microsoft.VSTS.Scheduling.RevisedDueDate') {
                $finishDate = $workItem.fields.'Microsoft.VSTS.Scheduling.RevisedDueDate'
                Write-Host "  Using RevisedDueDate: $finishDate" -ForegroundColor Green
            } elseif ($workItem.fields.'Microsoft.VSTS.Scheduling.OriginalDueDate') {
                $finishDate = $workItem.fields.'Microsoft.VSTS.Scheduling.OriginalDueDate'
                Write-Host "  Using OriginalDueDate: $finishDate" -ForegroundColor Green
            } elseif ($workItem.fields.'Microsoft.VSTS.Scheduling.TargetDate') {
                $finishDate = $workItem.fields.'Microsoft.VSTS.Scheduling.TargetDate'
                Write-Host "  Using TargetDate: $finishDate" -ForegroundColor Green
            } else {
                Write-Host "  No finish date available - none of the expected fields are set" -ForegroundColor Red
            }
        }
    } else {
        Write-Host "No work items found with 'databricks' in the title" -ForegroundColor Red
        
        # Try broader search for "PFS"
        Write-Host "`nTrying broader search for 'PFS'..." -ForegroundColor Yellow
        
        $pfsQuery = @"
SELECT [System.Id], [System.Title], [System.WorkItemType], [System.State]
FROM WorkItems 
WHERE [System.TeamProject] = '$($config.AdoProjectName)' 
AND [System.Title] CONTAINS 'PFS'
"@
        
        $wiqlRequest2 = @{ query = $pfsQuery }
        $wiqlBody2 = $wiqlRequest2 | ConvertTo-Json -Depth 3
          $wiqlResponse2 = Invoke-RestMethod -Uri $wiqlUrl -Method Post -Headers $headers -Body $wiqlBody2
        if ($wiqlResponse2.workItems -and $wiqlResponse2.workItems.Count -gt 0) {
            Write-Host "Found $($wiqlResponse2.workItems.Count) work items with 'PFS' in title:" -ForegroundColor Green
            foreach ($item in $wiqlResponse2.workItems) {
                $detailUrl = "$orgUrl/$projectName/_apis/wit/workitems/$($item.id)?api-version=7.1"
                $detail = Invoke-RestMethod -Uri $detailUrl -Method Get -Headers $headers
                Write-Host "  ID $($item.id): $($detail.fields.'System.Title')" -ForegroundColor White
            }
        }
    }
    
} catch {
    Write-Host "Error occurred: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Full error: $_" -ForegroundColor Red
}

Write-Host "`n=== Diagnostic complete ===" -ForegroundColor Green
