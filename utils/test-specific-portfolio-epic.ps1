# Test script to check specific Portfolio Epic 2450 and Epic 15358
# Load configuration
. "$PSScriptRoot\..\config.ps1"

Write-Host "=== Testing Portfolio Epic 2450 and Epic 15358 ===" -ForegroundColor Green

# Setup headers for Azure DevOps API
$headers = @{
    'Authorization' = "Basic $([Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(":$($ProductionConfig.PersonalAccessToken)")))"
    'Content-Type' = 'application/json'
}

try {
    # Test 1: Check if Portfolio Epic 2450 exists
    Write-Host "`nTesting Portfolio Epic 2450..." -ForegroundColor Yellow
    $portfolioEpicUrl = "$($ProductionConfig.AdoOrganizationUrl)/_apis/wit/workitems/2450?`$expand=Relations&api-version=7.1"
    Write-Host "URL: $portfolioEpicUrl" -ForegroundColor Gray
    $portfolioEpic = Invoke-RestMethod -Uri $portfolioEpicUrl -Method Get -Headers $headers
    Write-Host "✓ Portfolio Epic 2450 found: $($portfolioEpic.fields.'System.Title')" -ForegroundColor Green
    Write-Host "  Area Path: $($portfolioEpic.fields.'System.AreaPath')" -ForegroundColor Gray
    Write-Host "  Work Item Type: $($portfolioEpic.fields.'System.WorkItemType')" -ForegroundColor Gray
    
    # Test 2: Check if Epic 15358 exists
    Write-Host "`nTesting Epic 15358..." -ForegroundColor Yellow
    $epicUrl = "$($ProductionConfig.AdoOrganizationUrl)/_apis/wit/workitems/15358?api-version=7.1"
    Write-Host "URL: $epicUrl" -ForegroundColor Gray
    $epic = Invoke-RestMethod -Uri $epicUrl -Method Get -Headers $headers
    Write-Host "✓ Epic 15358 found: $($epic.fields.'System.Title')" -ForegroundColor Green
    Write-Host "  Area Path: $($epic.fields.'System.AreaPath')" -ForegroundColor Gray
    Write-Host "  Work Item Type: $($epic.fields.'System.WorkItemType')" -ForegroundColor Gray
    
    # Test 3: Check relationships on Portfolio Epic 2450
    Write-Host "`nChecking relationships on Portfolio Epic 2450..." -ForegroundColor Yellow
    if ($portfolioEpic.relations) {
        Write-Host "Found $($portfolioEpic.relations.Count) relations:" -ForegroundColor Cyan
        foreach ($relation in $portfolioEpic.relations) {
            $relatedId = $relation.url.Split('/')[-1]
            Write-Host "  - $($relation.rel): Work item $relatedId" -ForegroundColor White
            if ($relatedId -eq "15358") {
                Write-Host "    ✓ Found Epic 15358 as a related item!" -ForegroundColor Green
            }
        }
    } else {
        Write-Host "No relations found on Portfolio Epic 2450" -ForegroundColor Red
    }
    
} catch {
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Full error: $_" -ForegroundColor Red
}

Write-Host "`nTest Complete" -ForegroundColor Green
