# Azure DevOps Export Launcher
# Simple launcher script for the Azure DevOps to Microsoft Project export tool

Write-Host "=====================================" -ForegroundColor Cyan
Write-Host "Azure DevOps to Microsoft Project Export Tool" -ForegroundColor Cyan
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host ""

# Check if configuration exists
$configPath = Join-Path $PSScriptRoot "config.ps1"
if (-not (Test-Path $configPath)) {
    Write-Host "ERROR: Configuration file not found!" -ForegroundColor Red
    Write-Host "Please ensure 'config.ps1' exists in the same directory." -ForegroundColor Red
    Write-Host "Edit config.ps1 to set your Azure DevOps connection details." -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit 1
}

# Load and validate configuration
try {
    . $configPath
    if (-not $ProductionConfig) {
        throw "Configuration not loaded properly"
    }
    
    Write-Host "Configuration loaded successfully!" -ForegroundColor Green
    Write-Host "Organization: $($ProductionConfig.AdoOrganizationUrl)" -ForegroundColor White
    Write-Host "Project: $($ProductionConfig.AdoProjectName)" -ForegroundColor White
    Write-Host "Output: $($ProductionConfig.OutputExcelPath)" -ForegroundColor White
    Write-Host ""
} catch {
    Write-Host "ERROR: Failed to load configuration: $($_.Exception.Message)" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

# Check for ImportExcel module
try {
    Import-Module ImportExcel -ErrorAction Stop
    Write-Host "ImportExcel module is available" -ForegroundColor Green
} catch {
    Write-Host "ImportExcel module not found - will fallback to CSV export" -ForegroundColor Yellow
    Write-Host "To install: Install-Module -Name ImportExcel -Force" -ForegroundColor Gray
}

Write-Host ""
Write-Host "Ready to export Azure DevOps work items!" -ForegroundColor Green
Write-Host ""

# Offer options
Write-Host "Select an option:" -ForegroundColor Cyan
Write-Host "1. Run with default settings" -ForegroundColor White
Write-Host "2. Specify custom output path" -ForegroundColor White
Write-Host "3. Specify custom area path" -ForegroundColor White
Write-Host "4. Advanced options" -ForegroundColor White
Write-Host "5. Exit" -ForegroundColor White
Write-Host ""

$choice = Read-Host "Enter your choice (1-5)"

$scriptPath = Join-Path $PSScriptRoot "export-ado-workitems.ps1"

switch ($choice) {
    "1" {
        Write-Host "Starting export with default settings..." -ForegroundColor Green
        & $scriptPath
    }
    "2" {
        $outputPath = Read-Host "Enter custom output path"
        if ($outputPath) {
            Write-Host "Starting export to: $outputPath" -ForegroundColor Green
            & $scriptPath -OutputPath $outputPath
        }
    }
    "3" {
        $areaPath = Read-Host "Enter area path"
        if ($areaPath) {
            Write-Host "Starting export from area: $areaPath" -ForegroundColor Green
            & $scriptPath -AreaPath $areaPath
        }
    }
    "4" {
        Write-Host ""
        Write-Host "Advanced Options:" -ForegroundColor Cyan
        $outputPath = Read-Host "Output path (leave empty for default)"
        $areaPath = Read-Host "Area path (leave empty for default)"
        $workItemTypes = Read-Host "Work item types (leave empty for default)"
        
        $params = @{}
        if ($outputPath) { $params.OutputPath = $outputPath }
        if ($areaPath) { $params.AreaPath = $areaPath }
        if ($workItemTypes) { $params.WorkItemTypes = $workItemTypes }
        
        Write-Host "Starting export with custom settings..." -ForegroundColor Green
        & $scriptPath @params
    }
    "5" {
        Write-Host "Exiting..." -ForegroundColor Yellow
        exit 0
    }
    default {
        Write-Host "Invalid choice. Exiting..." -ForegroundColor Red
        exit 1
    }
}

Write-Host ""
Write-Host "Export completed. Press any key to exit..." -ForegroundColor Green
Read-Host "Press Enter to exit"
