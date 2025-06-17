# DEBUG MODE CONFIGURATION GUIDE
# ==============================

# The export script now respects the debug configuration setting in config.ps1

# PRODUCTION MODE (Default - Clean Output):
# Set in config.ps1:
$ENABLE_DEBUG_LOGGING = $false

# This provides clean output with only essential information:
# - Connection status
# - Work item counts
# - Export progress
# - Success/error messages

# DEBUG MODE (Troubleshooting):
# Set in config.ps1:
$ENABLE_DEBUG_LOGGING = $true

# This provides detailed information including:
# - API call details
# - Hierarchy building process
# - Work item relationship processing
# - Field extraction details
# - Parameter validation
# - Missing parent detection

# QUICK TOGGLE:
# You can quickly enable/disable debug mode by editing line 108 in config.ps1:
# 
# For production: $ENABLE_DEBUG_LOGGING = $false
# For debugging:  $ENABLE_DEBUG_LOGGING = $true

# The script automatically respects this setting - no need to modify the export script itself.
