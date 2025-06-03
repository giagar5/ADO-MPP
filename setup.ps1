# Azure DevOps to Microsoft Project Export Tool - Setup Script
# This script helps you set up the tool safely without exposing sensitive data

Write-Host "=== Azure DevOps to Microsoft Project Export Tool Setup ===" -ForegroundColor Cyan
Write-Host ""

# Check if config.ps1 already exists
if (Test-Path "config.ps1") {
    Write-Host "⚠️  config.ps1 already exists!" -ForegroundColor Yellow
    $overwrite = Read-Host "Do you want to overwrite it? (y/N)"
    if ($overwrite -ne "y" -and $overwrite -ne "Y") {
        Write-Host "Setup cancelled. Your existing config.ps1 was not modified." -ForegroundColor Green
        exit 0
    }
}

# Check if example config exists
if (-not (Test-Path "config.example.ps1")) {
    Write-Host "❌ config.example.ps1 not found!" -ForegroundColor Red
    Write-Host "Please ensure you have the complete tool package." -ForegroundColor Red
    exit 1
}

# Copy example to config
Copy-Item "config.example.ps1" "config.ps1"
Write-Host "✅ Created config.ps1 from example template" -ForegroundColor Green

# Check for ImportExcel module
Write-Host ""
Write-Host "🔍 Checking PowerShell modules..." -ForegroundColor Cyan

if (Get-Module -ListAvailable -Name ImportExcel) {
    Write-Host "✅ ImportExcel module is installed" -ForegroundColor Green
} else {
    Write-Host "⚠️  ImportExcel module not found" -ForegroundColor Yellow
    $install = Read-Host "Do you want to install it now? (Y/n)"
    if ($install -ne "n" -and $install -ne "N") {
        try {
            Install-Module -Name ImportExcel -Force -Scope CurrentUser
            Write-Host "✅ ImportExcel module installed successfully" -ForegroundColor Green
        } catch {
            Write-Host "❌ Failed to install ImportExcel module: $_" -ForegroundColor Red
            Write-Host "You can install it manually later with: Install-Module -Name ImportExcel -Force" -ForegroundColor Yellow
        }
    }
}

Write-Host ""
Write-Host "📝 Next steps:" -ForegroundColor Cyan
Write-Host "1. Edit config.ps1 with your Azure DevOps details:" -ForegroundColor White
Write-Host "   - Organization URL (line ~11)" -ForegroundColor Gray
Write-Host "   - Project name (line ~14)" -ForegroundColor Gray
Write-Host "   - Personal Access Token (line ~19)" -ForegroundColor Gray
Write-Host "   - WIQL query for work item selection (line ~26)" -ForegroundColor Gray
Write-Host ""
Write-Host "2. Generate your Personal Access Token:" -ForegroundColor White
Write-Host "   - Go to Azure DevOps > User Settings > Personal Access Tokens" -ForegroundColor Gray
Write-Host "   - Create new token with 'Work Items (Read)' permission" -ForegroundColor Gray
Write-Host ""
Write-Host "3. Run the export:" -ForegroundColor White
Write-Host "   .\export-ado-workitems.ps1" -ForegroundColor Gray
Write-Host ""
Write-Host "🔒 Security reminder: config.ps1 is ignored by git to protect your credentials" -ForegroundColor Yellow
Write-Host ""
Write-Host "Setup complete! Happy exporting! 🚀" -ForegroundColor Green
