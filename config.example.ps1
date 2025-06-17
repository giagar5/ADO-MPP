# Azure DevOps to Microsoft Project Export - Example Configuration
# Copy this file to config.ps1 and update with your actual values
# 
# SECURITY IMPORTANT: Never commit actual credentials to version control!

# =============================================================================
# AZURE DEVOPS CONNECTION SETTINGS
# =============================================================================

# Your Azure DevOps organization URL (without trailing slash)
# Example: "https://dev.azure.com/mycompany"
$ORGANIZATION_URL = "https://dev.azure.com/YOUR-ORGANIZATION"

# Your Azure DevOps project name
# Example: "MyProject"
$PROJECT_NAME = "YOUR-PROJECT-NAME"

# Personal Access Token with Work Items Read permissions
# Generate this in Azure DevOps: User Settings > Personal Access Tokens
# SECURITY NOTE: Never commit actual PAT tokens to version control!
# Consider using environment variables: $env:ADO_PAT
$PERSONAL_ACCESS_TOKEN = "YOUR-PERSONAL-ACCESS-TOKEN-HERE"

# =============================================================================
# WORK ITEM QUERY SETTINGS
# =============================================================================

# WIQL Query to select work items to export
# Customize the area path and work item types for your project
$WIQL_QUERY = @"
SELECT [System.Id] 
FROM WorkItems 
WHERE [System.TeamProject] = '$PROJECT_NAME' 
AND [System.WorkItemType] IN ('Epic', 'Feature', 'User Story', 'Task', 'Bug', 'Dependency', 'Milestone') 
AND [System.AreaPath] UNDER 'YOUR-PROJECT-NAME\YOUR-AREA-PATH'
"@

# =============================================================================
# OUTPUT SETTINGS
# =============================================================================

# Output file path for the Excel file
$OUTPUT_EXCEL_PATH = "C:\temp\AzureDevOpsExport_ProjectImport.xlsx"

# Batch size for API calls (adjust if you encounter timeout issues)
$BATCH_SIZE = 200

# Batch size for relationship processing (smaller to avoid timeouts)
$RELATIONSHIP_BATCH_SIZE = 25

# Working hours per day for duration calculations
$HOURS_PER_DAY = 8

# =============================================================================
# FIELD MAPPING SETTINGS
# =============================================================================

# Azure DevOps fields to retrieve and export
# Add or remove fields based on your requirements
# 
# CUSTOM PROGRESS FIELD:
# Replace "Custom.Progress" with your actual custom field name.
# Common examples:
# - "Custom.Progress" (typical custom field)
# - "Microsoft.VSTS.Scheduling.CompletedWork" (standard completed work)
# - "YourCompany.ProgressPercentage" (company-specific field)
# 
# To find your custom field name:
# 1. Go to Azure DevOps work item
# 2. Right-click and "Inspect Element" on your progress field
# 3. Look for the field name in the HTML attributes
$FIELDS_TO_FETCH = @(
    "System.Id",
    "System.Title", 
    "System.WorkItemType",
    "System.State",
    "System.AssignedTo",
    "System.AreaPath",
    "System.CreatedDate",
    "System.ChangedDate",
    "System.Description",
    "System.Tags",
    "Microsoft.VSTS.Scheduling.OriginalEstimate",
    "Microsoft.VSTS.Scheduling.RemainingWork",
    "Microsoft.VSTS.Scheduling.CompletedWork",
    "Microsoft.VSTS.Scheduling.StartDate",
    "Microsoft.VSTS.Scheduling.TargetDate",
    "Microsoft.VSTS.Common.Priority",
    "progress"                           # Add your custom progress field name here
)

# =============================================================================
# REGIONAL SETTINGS
# =============================================================================

# Regional formatting settings - can be 'Auto', 'US', 'European', or 'Custom'
# 'Auto' - Detect from system regional settings
# 'US' - Use US format (period decimal separator, comma list separator)
# 'European' - Use European format (period decimal separator, semicolon list separator)
# 'Custom' - Use custom settings defined below
$REGIONAL_FORMAT = "European"

# Custom regional settings (used only when REGIONAL_FORMAT = 'Custom')
$CUSTOM_DECIMAL_SEPARATOR = "."
$CUSTOM_LIST_SEPARATOR = ";"
$CUSTOM_THOUSANDS_SEPARATOR = ""

# =============================================================================
# ADVANCED SETTINGS
# =============================================================================

# Enable/disable relationship processing (predecessor/successor links)
$PROCESS_RELATIONSHIPS = $true

# Timeout for API calls (in seconds)
$API_TIMEOUT = 60

# Enable detailed logging
$ENABLE_DEBUG_LOGGING = $false

# Test mode: limit number of work items for testing (0 = no limit)
$TEST_MODE_LIMIT = 0

# =============================================================================
# EXPORT ALL SETTINGS AS A CONFIGURATION OBJECT
# =============================================================================

$ProductionConfig = @{
    AdoOrganizationUrl = $ORGANIZATION_URL
    AdoProjectName = $PROJECT_NAME
    PersonalAccessToken = $PERSONAL_ACCESS_TOKEN
    WiqlQuery = $WIQL_QUERY
    OutputExcelPath = $OUTPUT_EXCEL_PATH
    BatchSize = $BATCH_SIZE
    RelationshipBatchSize = $RELATIONSHIP_BATCH_SIZE
    HoursPerDay = $HOURS_PER_DAY
    FieldsToFetch = $FIELDS_TO_FETCH
    ProcessRelationships = $PROCESS_RELATIONSHIPS
    ApiTimeout = $API_TIMEOUT
    EnableDebugLogging = $ENABLE_DEBUG_LOGGING
    TestModeLimit = $TEST_MODE_LIMIT
    RegionalFormat = $REGIONAL_FORMAT
    CustomDecimalSeparator = $CUSTOM_DECIMAL_SEPARATOR
    CustomListSeparator = $CUSTOM_LIST_SEPARATOR
    CustomThousandsSeparator = $CUSTOM_THOUSANDS_SEPARATOR
}

# Validation function
function Test-ProductionConfig {
    param($Config)
    
    $errors = @()
    
    if ([string]::IsNullOrWhiteSpace($Config.AdoOrganizationUrl)) {
        $errors += "Organization URL is required"
    }
    
    if ([string]::IsNullOrWhiteSpace($Config.AdoProjectName)) {
        $errors += "Project name is required"
    }
    
    if ([string]::IsNullOrWhiteSpace($Config.PersonalAccessToken)) {
        $errors += "Personal Access Token is required"
    }
    
    if ($Config.PersonalAccessToken -eq "YOUR-PERSONAL-ACCESS-TOKEN-HERE") {
        $errors += "Please replace placeholder Personal Access Token with your actual token"
    }
    
    if ([string]::IsNullOrWhiteSpace($Config.WiqlQuery)) {
        $errors += "WIQL Query is required"
    }
    
    if ([string]::IsNullOrWhiteSpace($Config.OutputExcelPath)) {
        $errors += "Output Excel path is required"
    }
    
    if ($errors.Count -gt 0) {
        Write-Host "Configuration Validation Errors:" -ForegroundColor Red
        foreach ($error in $errors) {
            Write-Host "  - $error" -ForegroundColor Red
        }
        return $false
    }
    
    Write-Host "Configuration validation passed!" -ForegroundColor Green
    return $true
}

# Export the configuration
if (Test-ProductionConfig -Config $ProductionConfig) {
    Write-Host "Production configuration loaded successfully!" -ForegroundColor Green
    Write-Host "Organization: $($ProductionConfig.AdoOrganizationUrl)" -ForegroundColor Cyan
    Write-Host "Project: $($ProductionConfig.AdoProjectName)" -ForegroundColor Cyan
    Write-Host "Output: $($ProductionConfig.OutputExcelPath)" -ForegroundColor Cyan
} else {
    Write-Host "Please fix configuration errors before proceeding." -ForegroundColor Red
}
