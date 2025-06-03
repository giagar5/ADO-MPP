# Regional Settings Configuration Guide

## Overview

The Azure DevOps to Microsoft Project export tool includes comprehensive regional settings support to ensure proper delimiter configuration for optimal Microsoft Project compatibility across different geographical regions and system configurations.

## 🌍 Regional Settings Options

### 1. Auto Detection (`"Auto"`)
**Best for**: Most users who want automatic system compatibility
```powershell
$REGIONAL_FORMAT = "Auto"
```

**What it does**:
- Automatically detects your system's regional settings
- Uses your system's culture (e.g., en-US, de-DE, fr-FR, it-IT)
- Adapts list separators based on system configuration
- Always uses period (.) as decimal separator for Microsoft Project compatibility

**Auto-detection logic**:
- If system list separator is comma (,) → Uses semicolon (;) for predecessor lists
- If system list separator is semicolon (;) → Uses semicolon (;) for predecessor lists
- Always uses period (.) for decimal separator to ensure Microsoft Project compatibility

### 2. US Format (`"US"`)
**Best for**: US-based teams and standard US regional settings
```powershell
$REGIONAL_FORMAT = "US"
```

**Configuration**:
- Decimal separator: `.` (period)
- List separator: `,` (comma)
- Thousands separator: none

**Example output**:
- Number: `123.45` → `123.45`
- Predecessor list: `[1, 2, 3]` → `1,2,3`

### 3. European Format (`"European"`)
**Best for**: European teams and most international users
```powershell
$REGIONAL_FORMAT = "European"
```

**Configuration**:
- Decimal separator: `.` (period)
- List separator: `;` (semicolon)
- Thousands separator: none

**Example output**:
- Number: `123.45` → `123.45`
- Predecessor list: `[1, 2, 3]` → `1;2;3`

### 4. Custom Format (`"Custom"`)
**Best for**: Organizations with specific requirements
```powershell
$REGIONAL_FORMAT = "Custom"
$CUSTOM_DECIMAL_SEPARATOR = ","
$CUSTOM_LIST_SEPARATOR = "|"
$CUSTOM_THOUSANDS_SEPARATOR = ""
```

**Configuration**: User-defined separators
**Example output**:
- Number: `123.45` → `123,45`
- Predecessor list: `[1, 2, 3]` → `1|2|3`

## 🔧 Configuration Instructions

### Step 1: Open Configuration File
Edit `config.ps1` in your ADO-MPP directory.

### Step 2: Set Regional Format
Find the `REGIONAL SETTINGS` section and modify:

```powershell
# =============================================================================
# REGIONAL SETTINGS
# =============================================================================

# Choose one: 'Auto', 'US', 'European', or 'Custom'
$REGIONAL_FORMAT = "European"  # ← Change this value

# Custom settings (only used when REGIONAL_FORMAT = 'Custom')
$CUSTOM_DECIMAL_SEPARATOR = "."
$CUSTOM_LIST_SEPARATOR = ";"
$CUSTOM_THOUSANDS_SEPARATOR = ""
```

### Step 3: Save and Test
```powershell
# Test your configuration
.\test-regional-settings.ps1
```

## 📊 Regional Format Comparison

| Format | Decimal Sep | List Sep | Use Case | Microsoft Project Compatibility |
|--------|-------------|----------|----------|--------------------------------|
| Auto | `.` | System-based | Automatic detection | ✅ Excellent |
| US | `.` | `,` | US organizations | ✅ Good |
| European | `.` | `;` | European organizations | ✅ Excellent |
| Custom | User-defined | User-defined | Special requirements | ⚠️ Depends on settings |

## 🌍 System Culture Examples

### Common Culture Mappings

| System Culture | Country/Region | Auto Detection Result |
|---|---|---|
| `en-US` | United States | List: `;` (adapted), Decimal: `.` |
| `en-GB` | United Kingdom | List: `;`, Decimal: `.` |
| `de-DE` | Germany | List: `;`, Decimal: `.` |
| `fr-FR` | France | List: `;`, Decimal: `.` |
| `it-IT` | Italy | List: `;`, Decimal: `.` |
| `es-ES` | Spain | List: `;`, Decimal: `.` |
| `pt-BR` | Brazil | List: `;`, Decimal: `.` |
| `ja-JP` | Japan | List: `;`, Decimal: `.` |

## 🔍 Testing Your Configuration

### Run the Test Script
```powershell
.\test-regional-settings.ps1
```

### Expected Output Example
```
=== TESTING REGIONAL CONFIGURATION SYSTEM ===

--- Testing European Configuration ---
Determining regional settings for format: European
Using European regional format
Final Settings - Decimal: '.', List: ';', Thousands: ''

Number formatting examples:
  123.45 → '123.45'
  1000 → '1000'
  1.2345 → '1.23'

Predecessor list: [1, 2, 3, 15, 25] → '1;2;3;15;25'
```

## 📋 Best Practices

### 1. Recommended Settings by Region

**🇺🇸 North America**: 
```powershell
$REGIONAL_FORMAT = "US"  # or "Auto"
```

**🇪🇺 Europe**: 
```powershell
$REGIONAL_FORMAT = "European"  # or "Auto"
```

**🌏 Asia-Pacific**: 
```powershell
$REGIONAL_FORMAT = "Auto"  # or "European"
```

### 2. Microsoft Project Import Considerations

**For Best Compatibility**:
- Use `"European"` format for most reliable imports
- Semicolon (`;`) list separators work consistently across Microsoft Project versions
- Period (`.`) decimal separators are universally supported

**Avoid**:
- Comma (`,`) as both decimal and list separator in same configuration
- Complex thousands separators that may confuse import parsers

### 3. Team Collaboration

**For Mixed Teams**:
- Use `"Auto"` to respect individual system settings
- Or standardize on `"European"` for consistency
- Document your choice in team guidelines

## 🛠️ Troubleshooting

### Common Issues

#### Issue: Numbers not formatting correctly
**Solution**: Check that your test numbers have decimal places
```powershell
# This will show "123" (no decimals)
Format-NumberForRegion -Number 123

# This will show "123.45" or "123,45" depending on separator
Format-NumberForRegion -Number 123.45
```

#### Issue: Predecessor lists not importing correctly
**Solutions**:
1. Try `"European"` format (semicolon separators)
2. Check Microsoft Project import field mapping
3. Ensure "Predecessors" field is correctly mapped during import

#### Issue: Auto-detection not working as expected
**Solutions**:
1. Check your system culture: `[System.Globalization.CultureInfo]::CurrentCulture.Name`
2. Manually set to `"European"` or `"US"` instead
3. Use `"Custom"` for complete control

### Testing Different Configurations

```powershell
# Test US format
$config.RegionalFormat = "US"
$settings = Get-RegionalSettings -Config $config

# Test custom format
$config.RegionalFormat = "Custom"
$config.CustomDecimalSeparator = ","
$config.CustomListSeparator = "|"
$settings = Get-RegionalSettings -Config $config
```

## 📈 Advanced Configuration

### Custom Number Formatting

For special requirements, you can modify the `Format-NumberForRegion` function in the main script:

```powershell
function Format-NumberForRegion {
    param($Number, $RegionalSettings = $null)
    
    # Your custom formatting logic here
    # Example: Force 2 decimal places
    $formatted = $numericValue.ToString("0.00", [System.Globalization.CultureInfo]::InvariantCulture)
    
    # Apply regional separators
    if ($decimalSep -ne "." -and $formatted.Contains(".")) {
        $formatted = $formatted.Replace(".", $decimalSep)
    }
    
    return $formatted
}
```

### Environment-Specific Settings

```powershell
# Development environment
if ($env:COMPUTERNAME -eq "DEV-MACHINE") {
    $REGIONAL_FORMAT = "US"
}

# Production environment
if ($env:ENVIRONMENT -eq "PRODUCTION") {
    $REGIONAL_FORMAT = "Auto"
}
```

## 🔄 Migration Guide

### From Hardcoded to Regional Settings

If you're upgrading from a previous version:

1. **Backup your config.ps1**
2. **Add regional settings section**:
   ```powershell
   $REGIONAL_FORMAT = "European"  # Choose appropriate
   $CUSTOM_DECIMAL_SEPARATOR = "."
   $CUSTOM_LIST_SEPARATOR = ";"
   $CUSTOM_THOUSANDS_SEPARATOR = ""
   ```
3. **Add to ProductionConfig hashtable**:
   ```powershell
   RegionalFormat = $REGIONAL_FORMAT
   CustomDecimalSeparator = $CUSTOM_DECIMAL_SEPARATOR
   CustomListSeparator = $CUSTOM_LIST_SEPARATOR
   CustomThousandsSeparator = $CUSTOM_THOUSANDS_SEPARATOR
   ```
4. **Test with your data**:
   ```powershell
   .\test-regional-settings.ps1
   ```

## 📞 Support

### Validation Steps

Before reporting issues:

1. **Run test script**: `.\test-regional-settings.ps1`
2. **Check system culture**: 
   ```powershell
   [System.Globalization.CultureInfo]::CurrentCulture.Name
   ```
3. **Verify configuration**: Check `config.ps1` regional settings
4. **Test with small dataset**: Use `$TEST_MODE_LIMIT = 10`

### Getting Help

1. Include output from `test-regional-settings.ps1`
2. Specify your system culture and desired format
3. Provide sample numbers that aren't formatting correctly
4. Include Microsoft Project version for import issues

---

## 📝 Summary

The regional settings system provides flexible delimiter configuration that automatically adapts to different geographical and organizational requirements while maintaining optimal Microsoft Project compatibility. Use `"Auto"` for automatic detection, `"European"` for the most reliable imports, or `"Custom"` for specific organizational needs.

**Quick Start**: Set `$REGIONAL_FORMAT = "Auto"` in `config.ps1` and run `.\test-regional-settings.ps1` to verify everything works correctly with your system.
