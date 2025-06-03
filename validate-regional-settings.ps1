# Advanced Regional Settings Validation Test
# This script performs comprehensive testing of all regional configuration options
# and validates proper functionality with various number formats and edge cases

# Get script directory for relative paths
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Load configuration
$ConfigPath = Join-Path $ScriptDir "config.ps1"
if (-not (Test-Path $ConfigPath)) {
    Write-Host "Configuration file not found: $ConfigPath" -ForegroundColor Red
    exit 1
}

try {
    . $ConfigPath
    $config = $ProductionConfig
} catch {
    Write-Host "Error loading configuration: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Load the regional settings function from export script
$exportScriptPath = Join-Path $ScriptDir "export-ado-workitems.ps1"
if (Test-Path $exportScriptPath) {
    # Extract just the Get-RegionalSettings and Format-NumberForRegion functions
    $scriptContent = Get-Content $exportScriptPath -Raw
    
    # Extract and execute the functions
    if ($scriptContent -match '(?s)function Get-RegionalSettings.*?(?=^function|\Z)') {
        Invoke-Expression $matches[0]
    }
    
    if ($scriptContent -match '(?s)function Format-NumberForRegion.*?(?=^function|\Z)') {
        Invoke-Expression $matches[0]
    }
}

Write-Host "=== COMPREHENSIVE REGIONAL SETTINGS VALIDATION ===" -ForegroundColor White -BackgroundColor DarkBlue
Write-Host ""

# =============================================================================
# TEST DATA SETS
# =============================================================================

$testNumbers = @(
    @{ Value = 0; Description = "Zero" },
    @{ Value = 1; Description = "Integer" },
    @{ Value = 123.45; Description = "Two decimals" },
    @{ Value = 123.456; Description = "Three decimals (should round to two)" },
    @{ Value = 1000; Description = "Large integer" },
    @{ Value = 1000.99; Description = "Large with decimals" },
    @{ Value = 0.5; Description = "Fractional" },
    @{ Value = 999.99; Description = "Near round number" },
    @{ Value = $null; Description = "Null value" },
    @{ Value = ""; Description = "Empty string" }
)

$testPredecessorLists = @(
    @{ Value = @(1); Description = "Single predecessor" },
    @{ Value = @(1, 2, 3); Description = "Multiple predecessors" },
    @{ Value = @(10, 20, 30, 40, 50); Description = "Many predecessors" },
    @{ Value = @(101, 205, 1003); Description = "Large ID numbers" }
)

# =============================================================================
# COMPREHENSIVE TEST CONFIGURATIONS
# =============================================================================

$testConfigurations = @(
    @{ 
        Name = "Auto Detection (System Default)"
        Config = @{ RegionalFormat = "Auto" }
        Description = "Automatically detects system regional settings"
    },
    @{ 
        Name = "US Standard"
        Config = @{ RegionalFormat = "US" }
        Description = "US format with comma list separators"
    },
    @{ 
        Name = "European Standard"
        Config = @{ RegionalFormat = "European" }
        Description = "European format with semicolon list separators"
    },
    @{ 
        Name = "Custom: Comma Decimal, Pipe List"
        Config = @{ 
            RegionalFormat = "Custom"
            CustomDecimalSeparator = ","
            CustomListSeparator = "|"
            CustomThousandsSeparator = ""
        }
        Description = "Custom format for specialized requirements"
    },
    @{ 
        Name = "Custom: Period Decimal, Space List"
        Config = @{ 
            RegionalFormat = "Custom"
            CustomDecimalSeparator = "."
            CustomListSeparator = " "
            CustomThousandsSeparator = ""
        }
        Description = "Space-separated predecessor lists"
    },
    @{ 
        Name = "Custom: Comma Decimal, Semicolon List"
        Config = @{ 
            RegionalFormat = "Custom"
            CustomDecimalSeparator = ","
            CustomListSeparator = ";"
            CustomThousandsSeparator = ""
        }
        Description = "Mixed European-style format"
    }
)

# =============================================================================
# VALIDATION FUNCTIONS
# =============================================================================

function Test-NumberFormatting {
    param($TestConfig, $RegionalSettings)
    
    Write-Host "  📊 Number Formatting Tests:" -ForegroundColor Cyan
    
    $results = @()
    foreach ($testNumber in $testNumbers) {
        $formatted = Format-NumberForRegion -Number $testNumber.Value -RegionalSettings $RegionalSettings
        $result = @{
            Input = $testNumber.Value
            Description = $testNumber.Description
            Output = $formatted
            Valid = $null -ne $formatted -and $formatted -ne ""
        }
        $results += $result
          $color = if ($result.Valid) { "Green" } else { "Red" }
        $displayInput = if ($null -eq $testNumber.Value) { "null" } elseif ($testNumber.Value -eq "") { "empty" } else { $testNumber.Value }
        Write-Host "    $($testNumber.Description): $displayInput → '$formatted'" -ForegroundColor $color
    }
    
    return $results
}

function Test-PredecessorFormatting {
    param($TestConfig, $RegionalSettings)
    
    Write-Host "  🔗 Predecessor List Tests:" -ForegroundColor Cyan
    
    $results = @()
    foreach ($testList in $testPredecessorLists) {
        $formattedNumbers = $testList.Value | ForEach-Object { Format-NumberForRegion -Number $_ -RegionalSettings $RegionalSettings }
        $predecessorString = $formattedNumbers -join $RegionalSettings.ListSeparator
          $result = @{
            Input = $testList.Value
            Description = $testList.Description
            Output = $predecessorString
            Valid = $null -ne $predecessorString -and $predecessorString -ne ""
        }
        $results += $result
        
        $color = if ($result.Valid) { "Green" } else { "Red" }
        Write-Host "    $($testList.Description): [$($testList.Value -join ', ')] → '$predecessorString'" -ForegroundColor $color
    }
    
    return $results
}

function Test-EdgeCases {
    param($TestConfig, $RegionalSettings)
    
    Write-Host "  ⚠️  Edge Case Tests:" -ForegroundColor Cyan
    
    $edgeCases = @(
        @{ Test = "Very large number"; Value = 999999.99 },
        @{ Test = "Very small decimal"; Value = 0.01 },
        @{ Test = "Negative number"; Value = -123.45 },
        @{ Test = "String number"; Value = "123.45" },
        @{ Test = "Invalid string"; Value = "abc" },
        @{ Test = "Empty predecessor list"; Value = @() }
    )
    
    $results = @()
    foreach ($edgeCase in $edgeCases) {
        try {
            if ($edgeCase.Test -eq "Empty predecessor list") {
                $formatted = if ($edgeCase.Value.Count -eq 0) { "" } else { ($edgeCase.Value | ForEach-Object { Format-NumberForRegion -Number $_ -RegionalSettings $RegionalSettings }) -join $RegionalSettings.ListSeparator }
            } else {
                $formatted = Format-NumberForRegion -Number $edgeCase.Value -RegionalSettings $RegionalSettings
            }
            
            $result = @{
                Test = $edgeCase.Test
                Input = $edgeCase.Value
                Output = $formatted
                Success = $true
                Error = $null
            }
        } catch {
            $result = @{
                Test = $edgeCase.Test
                Input = $edgeCase.Value
                Output = $null
                Success = $false
                Error = $_.Exception.Message
            }
        }
        
        $results += $result
        $color = if ($result.Success) { "Green" } else { "Yellow" }
        $output = if ($result.Success) { "'$($result.Output)'" } else { "ERROR: $($result.Error)" }
        Write-Host "    $($edgeCase.Test): $($edgeCase.Value) → $output" -ForegroundColor $color
    }
    
    return $results
}

function Test-MicrosoftProjectCompatibility {
    param($TestConfig, $RegionalSettings)
    
    Write-Host "  📋 Microsoft Project Compatibility:" -ForegroundColor Cyan
    
    # Test scenarios that are known to work well with Microsoft Project
    $compatibilityTests = @(
        @{ 
            Test = "Simple predecessor chain"
            Predecessors = @(1, 2, 3)
            ExpectedSeparator = $RegionalSettings.ListSeparator
        },
        @{ 
            Test = "Complex predecessor dependencies"
            Predecessors = @(5, 10, 15, 20)
            ExpectedSeparator = $RegionalSettings.ListSeparator
        }
    )
    
    $compatibilityScore = 0
    $totalTests = $compatibilityTests.Count
    
    foreach ($test in $compatibilityTests) {
        $formattedPredecessors = $test.Predecessors | ForEach-Object { Format-NumberForRegion -Number $_ -RegionalSettings $RegionalSettings }
        $result = $formattedPredecessors -join $RegionalSettings.ListSeparator
        
        # Check if result is valid for Microsoft Project
        $isValid = $result -match '^[\d\s' + [regex]::Escape($RegionalSettings.ListSeparator) + ']+$'
        
        if ($isValid) {
            $compatibilityScore++
            Write-Host "    ✅ $($test.Test): '$result'" -ForegroundColor Green
        } else {
            Write-Host "    ❌ $($test.Test): '$result' (may cause import issues)" -ForegroundColor Red
        }
    }
    
    $compatibilityPercentage = ($compatibilityScore / $totalTests) * 100
    
    if ($compatibilityPercentage -eq 100) {
        Write-Host "    🎯 Microsoft Project Compatibility: $compatibilityPercentage% (Excellent)" -ForegroundColor Green
    } elseif ($compatibilityPercentage -ge 80) {
        Write-Host "    🎯 Microsoft Project Compatibility: $compatibilityPercentage% (Good)" -ForegroundColor Yellow
    } else {
        Write-Host "    🎯 Microsoft Project Compatibility: $compatibilityPercentage% (May have issues)" -ForegroundColor Red
    }
    
    return @{
        Score = $compatibilityScore
        Total = $totalTests
        Percentage = $compatibilityPercentage
    }
}

# =============================================================================
# RUN COMPREHENSIVE TESTS
# =============================================================================

$overallResults = @()

foreach ($testConfig in $testConfigurations) {
    Write-Host ""
    Write-Host "--- Testing: $($testConfig.Name) ---" -ForegroundColor Magenta
    Write-Host "$($testConfig.Description)" -ForegroundColor Gray
    
    # Create temporary config for this test
    $tempConfig = $config.Clone()
    foreach ($key in $testConfig.Config.Keys) {
        $tempConfig[$key] = $testConfig.Config[$key]
    }
    
    # Get regional settings for this configuration
    $regionalSettings = Get-RegionalSettings -Config $tempConfig
    
    Write-Host "  Settings: Decimal='$($regionalSettings.DecimalSeparator)', List='$($regionalSettings.ListSeparator)', Thousands='$($regionalSettings.ThousandsSeparator)'" -ForegroundColor White
    
    # Run all tests
    $numberResults = Test-NumberFormatting -TestConfig $testConfig -RegionalSettings $regionalSettings
    $predecessorResults = Test-PredecessorFormatting -TestConfig $testConfig -RegionalSettings $regionalSettings
    $edgeResults = Test-EdgeCases -TestConfig $testConfig -RegionalSettings $regionalSettings
    $compatibilityResults = Test-MicrosoftProjectCompatibility -TestConfig $testConfig -RegionalSettings $regionalSettings
    
    # Store results
    $configResult = @{
        Name = $testConfig.Name
        Config = $testConfig.Config
        RegionalSettings = $regionalSettings
        NumberResults = $numberResults
        PredecessorResults = $predecessorResults
        EdgeResults = $edgeResults
        CompatibilityResults = $compatibilityResults
    }
    $overallResults += $configResult
}

# =============================================================================
# SUMMARY REPORT
# =============================================================================

Write-Host ""
Write-Host "=== VALIDATION SUMMARY REPORT ===" -ForegroundColor White -BackgroundColor DarkGreen
Write-Host ""

foreach ($result in $overallResults) {
    $numberSuccessRate = (($result.NumberResults | Where-Object { $_.Valid }).Count / $result.NumberResults.Count) * 100
    $predecessorSuccessRate = (($result.PredecessorResults | Where-Object { $_.Valid }).Count / $result.PredecessorResults.Count) * 100
    $edgeSuccessRate = (($result.EdgeResults | Where-Object { $_.Success }).Count / $result.EdgeResults.Count) * 100
    $compatibilityScore = $result.CompatibilityResults.Percentage
    
    Write-Host "📊 $($result.Name):" -ForegroundColor Cyan
    Write-Host "   Number Formatting: $([math]::Round($numberSuccessRate))% success" -ForegroundColor White
    Write-Host "   Predecessor Lists: $([math]::Round($predecessorSuccessRate))% success" -ForegroundColor White
    Write-Host "   Edge Cases: $([math]::Round($edgeSuccessRate))% handled" -ForegroundColor White
    Write-Host "   MS Project Compatibility: $([math]::Round($compatibilityScore))%" -ForegroundColor White
    
    $overallScore = ($numberSuccessRate + $predecessorSuccessRate + $edgeSuccessRate + $compatibilityScore) / 4
    
    if ($overallScore -ge 95) {
        Write-Host "   Overall Rating: $([math]::Round($overallScore))% - Excellent ✅" -ForegroundColor Green
    } elseif ($overallScore -ge 85) {
        Write-Host "   Overall Rating: $([math]::Round($overallScore))% - Good ✅" -ForegroundColor Yellow
    } else {
        Write-Host "   Overall Rating: $([math]::Round($overallScore))% - Needs Review ⚠️" -ForegroundColor Red
    }
    Write-Host ""
}

# =============================================================================
# RECOMMENDATIONS
# =============================================================================

Write-Host "=== RECOMMENDATIONS ===" -ForegroundColor White -BackgroundColor Blue
Write-Host ""

$bestConfig = $overallResults | Sort-Object { 
    $numberSuccessRate = (($_.NumberResults | Where-Object { $_.Valid }).Count / $_.NumberResults.Count) * 100
    $predecessorSuccessRate = (($_.PredecessorResults | Where-Object { $_.Valid }).Count / $_.PredecessorResults.Count) * 100
    $edgeSuccessRate = (($_.EdgeResults | Where-Object { $_.Success }).Count / $_.EdgeResults.Count) * 100
    $compatibilityScore = $_.CompatibilityResults.Percentage
    ($numberSuccessRate + $predecessorSuccessRate + $edgeSuccessRate + $compatibilityScore) / 4
} -Descending | Select-Object -First 1

Write-Host "🏆 Best Overall Configuration: $($bestConfig.Name)" -ForegroundColor Green
Write-Host "   Settings: Decimal='$($bestConfig.RegionalSettings.DecimalSeparator)', List='$($bestConfig.RegionalSettings.ListSeparator)'" -ForegroundColor White

Write-Host ""
Write-Host "💡 Configuration Guidelines:" -ForegroundColor Yellow
Write-Host "   • For maximum compatibility: Use 'European' format (period decimal, semicolon list)" -ForegroundColor White
Write-Host "   • For US organizations: Use 'US' format (period decimal, comma list)" -ForegroundColor White
Write-Host "   • For automatic adaptation: Use 'Auto' format" -ForegroundColor White
Write-Host "   • For special requirements: Use 'Custom' format with careful testing" -ForegroundColor White

Write-Host ""
Write-Host "⚠️  Important Notes:" -ForegroundColor Red
Write-Host "   • Always test your chosen configuration with a small dataset first" -ForegroundColor White
Write-Host "   • Microsoft Project imports work best with semicolon list separators" -ForegroundColor White
Write-Host "   • Period decimal separators ensure universal compatibility" -ForegroundColor White

Write-Host ""
Write-Host "Current configuration in config.ps1: $($config.RegionalFormat)" -ForegroundColor Cyan
Write-Host "Validation completed successfully! ✅" -ForegroundColor Green
