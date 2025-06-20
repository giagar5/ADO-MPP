# PowerShell Module Installer for ADO2MPP
# This script ensures all required PowerShell modules are installed

param(
    [switch]$Force = $false
)

Write-Host "=== ADO2MPP Module Dependencies Installer ===" -ForegroundColor Green

# List of required modules
$requiredModules = @(
    @{
        Name = "ImportExcel"
        MinVersion = "7.0.0"
        Description = "Excel import/export functionality"
        Required = $true
    }
)

function Test-ModuleAvailable {
    param(
        [string]$ModuleName,
        [string]$MinVersion = "0.0.0"
    )
    
    try {
        $module = Get-Module -Name $ModuleName -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1
        if ($module) {
            if ([version]$module.Version -ge [version]$MinVersion) {
                return $true
            } else {
                Write-Host "Module $ModuleName is installed but version $($module.Version) is below required $MinVersion" -ForegroundColor Yellow
                return $false
            }
        }
        return $false
    } catch {
        return $false
    }
}

function Install-RequiredModule {
    param(
        [string]$ModuleName,
        [string]$MinVersion,
        [bool]$Force = $false
    )
    
    try {
        Write-Host "Installing $ModuleName (minimum version: $MinVersion)..." -ForegroundColor Yellow
        
        $installParams = @{
            Name = $ModuleName
            MinimumVersion = $MinVersion
            Scope = "CurrentUser"
            Force = $Force
            AllowClobber = $true
        }
        
        Install-Module @installParams
        Write-Host "Successfully installed $ModuleName" -ForegroundColor Green
        return $true
    } catch {
        Write-Host "Failed to install $ModuleName`: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Check and install each required module
$allModulesReady = $true

foreach ($moduleInfo in $requiredModules) {
    Write-Host "`nChecking module: $($moduleInfo.Name)" -ForegroundColor Cyan
    Write-Host "Description: $($moduleInfo.Description)" -ForegroundColor Gray
    
    if (Test-ModuleAvailable -ModuleName $moduleInfo.Name -MinVersion $moduleInfo.MinVersion) {
        Write-Host "✓ Module $($moduleInfo.Name) is available and meets requirements" -ForegroundColor Green
    } else {
        Write-Host "✗ Module $($moduleInfo.Name) needs to be installed" -ForegroundColor Red
        
        if ($moduleInfo.Required) {
            $success = Install-RequiredModule -ModuleName $moduleInfo.Name -MinVersion $moduleInfo.MinVersion -Force $Force
            if (-not $success) {
                $allModulesReady = $false
            }
        } else {
            Write-Host "Module $($moduleInfo.Name) is optional and will be skipped" -ForegroundColor Yellow
        }
    }
}

Write-Host "`n" + "="*60 -ForegroundColor Green

if ($allModulesReady) {
    Write-Host "✓ All required modules are installed and ready!" -ForegroundColor Green
    Write-Host "`nYou can now run the ADO2MPP export scripts:" -ForegroundColor White
    Write-Host "  - export-ado-workitems.ps1 (Main export)" -ForegroundColor Cyan
    Write-Host "  - export-critical-timeline.ps1 (Critical timeline for Office Timeline)" -ForegroundColor Cyan
    Write-Host "  - utils/check-portfolio-epics.ps1 (Portfolio analysis)" -ForegroundColor Cyan
} else {
    Write-Host "✗ Some required modules failed to install" -ForegroundColor Red
    Write-Host "`nPlease resolve the installation issues before running the export scripts." -ForegroundColor Yellow
    Write-Host "You may need to run as Administrator or with different execution policies." -ForegroundColor Yellow
}

Write-Host "`nModule installation check complete." -ForegroundColor Green
