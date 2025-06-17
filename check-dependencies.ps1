# Dependency Relationship Diagnostic Script
# This script checks specific work items for their dependency relationships

# Load configuration
. .\config.ps1

$base64Auth = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$PERSONAL_ACCESS_TOKEN"))
$headers = @{
    Authorization = "Basic $base64Auth"
    "Content-Type" = "application/json"
}

Write-Host "=== Dependency Relationship Diagnostic ===" -ForegroundColor Green

# Check specific work items mentioned
$workItems = @(10013, 12418, 21372)

foreach ($workItemId in $workItems) {
    Write-Host "`nChecking work item $workItemId..." -ForegroundColor Yellow
    
    try {
        # Get work item with relations
        $workItemUrl = "$ORGANIZATION_URL/$PROJECT_NAME/_apis/wit/workitems/$workItemId" + "?`$expand=relations&api-version=7.1"
        $workItem = Invoke-RestMethod -Uri $workItemUrl -Method Get -Headers $headers -TimeoutSec 30
        
        Write-Host "  Title: $($workItem.fields.'System.Title')" -ForegroundColor White
        Write-Host "  Type: $($workItem.fields.'System.WorkItemType')" -ForegroundColor White
        
        if ($workItem.relations) {
            Write-Host "  Relations found: $($workItem.relations.Count)" -ForegroundColor Cyan
            
            foreach ($relation in $workItem.relations) {
                Write-Host "    - Type: $($relation.rel)" -ForegroundColor Gray
                
                if ($relation.rel -like "*Dependency*") {
                    # Extract target work item ID
                    if ($relation.url -match '/(\d+)$') {
                        $targetId = $matches[1]
                        Write-Host "      → Target Work Item: $targetId" -ForegroundColor Green
                        
                        # Get target work item title
                        try {
                            $targetUrl = "$ORGANIZATION_URL/$PROJECT_NAME/_apis/wit/workitems/$targetId" + "?api-version=7.1"
                            $targetWorkItem = Invoke-RestMethod -Uri $targetUrl -Method Get -Headers $headers -TimeoutSec 10
                            Write-Host "      → Target Title: $($targetWorkItem.fields.'System.Title')" -ForegroundColor Green
                        } catch {
                            Write-Host "      → Could not fetch target details" -ForegroundColor Red
                        }
                    }
                }
            }
        } else {
            Write-Host "  No relations found" -ForegroundColor Red
        }
    } catch {
        Write-Host "  Error: $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Host "`n=== SUMMARY ===" -ForegroundColor Yellow
Write-Host "This diagnostic checked the specific work items you mentioned." -ForegroundColor White
Write-Host "If dependency relationships are found but not showing in the export," -ForegroundColor White
Write-Host "the issue might be that the related work items are outside the WIQL query scope." -ForegroundColor White
