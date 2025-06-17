# ADO2MPP
- 🎯 **Native Fields Core** - Uses Microsoft Project native fields for seamless import
- 📊 **Structured ADO Metadata** - ADO data in standard Text/Number fields for easy filtering and reporting Azure DevOps to Microsoft Project Bridge

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1+-blue.svg)](https://github.com/PowerShell/PowerShell)
[![Azure DevOps](https://img.shields.io/badge/Azure%20DevOps-Compatible-0078d4.svg)](https://azure.microsoft.com/en-us/products/devops/)
[![Microsoft Project](https://img.shields.io/badge/Microsoft%20Project-Compatible-217346.svg)](https://www.microsoft.com/en-us/microsoft-365/project/project-management-software)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

A comprehensive PowerShell-based tool that seamlessly exports work items from Azure DevOps to Excel format, optimized for Microsoft Project import with **native field compatibility** for effortless integration.

### ✨ **Key Features**
- 🎯 **Native Fields Only** - Uses only Microsoft Project native fields for seamless import
- � **ADO Integration** - All Azure DevOps metadata consolidated in Notes field with clickable URLs
- 🏗️ **Hierarchical Structure** - Maintains Epic → Feature → User Story → Task relationships
- � **Smart Date Logic** - Prioritizes revised due dates over original dates
- 👥 **Resource Assignment** - Prefers Owner field over AssignedTo when available
- � **Clickable URLs** - Direct links to Azure DevOps work items in Text3 field
- 🔄 **Easy Filtering** - Filter and group by work item type, state, and ID
- �️ **Production Ready** - Clean output with configurable debug mode

## 🚀 Quick Start

1. **Configure the tool**: Edit `config.ps1` with your Azure DevOps details
2. **Run the export**: Execute `.\export-ado-workitems.ps1`
3. **Import to Project**: Use the generated Excel file in Microsoft Project

## 📋 Requirements

- **PowerShell 5.1** or later
- **Azure DevOps** access with Personal Access Token
- **ImportExcel PowerShell module** (optional, will fallback to CSV)
- **Microsoft Project** for importing the results

## 🛠️ Installation & Setup

### 1. Install ImportExcel Module (Recommended)
```powershell
Install-Module -Name ImportExcel -Force
```

### 2. Configure Azure DevOps Connection
Copy `config.example.ps1` to `config.ps1` and update these key settings:

```powershell
# Copy the example configuration file
Copy-Item config.example.ps1 config.ps1

# Edit config.ps1 with your actual values
# Your Azure DevOps organization URL
$ORGANIZATION_URL = "https://dev.azure.com/your-organization"

# Your project name
$PROJECT_NAME = "Your-Project-Name"

# Your Personal Access Token (with Work Items Read permission)
$PERSONAL_ACCESS_TOKEN = "your-pat-token-here"
```

**🔒 Security Note**: The `config.ps1` file is ignored by git to prevent accidental commits of sensitive data. Always use `config.example.ps1` as your template.

### 3. Customize Work Item Selection
Modify the WIQL query in `config.ps1` to select the work items you want:

```powershell
# Example: Export all work items from a specific area path
$WIQL_QUERY = @"
SELECT [System.Id] 
FROM WorkItems 
WHERE [System.TeamProject] = '$PROJECT_NAME' 
AND [System.AreaPath] UNDER '$PROJECT_NAME\Your-Area-Path'
"@
```

## 🎯 Usage Examples

### Basic Export
```powershell
.\export-ado-workitems.ps1
```

### Export to Specific File
```powershell
.\export-ado-workitems.ps1 -OutputPath "C:\exports\my-project.xlsx"
```

### Export from Specific Area Path
```powershell
.\export-ado-workitems.ps1 -AreaPath "MyProject\Team Alpha"
```

### Export Specific Work Item Types
```powershell
.\export-ado-workitems.ps1 -WorkItemTypes "Epic,Feature,User Story"
```

### Use Custom Configuration File
```powershell
.\export-ado-workitems.ps1 -ConfigPath "C:\configs\my-config.ps1"
```

## 📊 Output Structure

The tool generates an Excel file with **native Microsoft Project fields** plus **structured ADO metadata**:

| Column | Description | Source |
|--------|-------------|---------|
| **Unique ID** | Task sequence number | Auto-generated |
| **Name** | Work item title | System.Title |
| **Outline Level** | Hierarchy level (1=Epic, 2=Feature, 3=Story, 4=Task) | Based on WorkItemType |
| **% Complete** | Progress percentage | Calculated from work estimates |
| **Start** | Start date | Microsoft.VSTS.Scheduling.StartDate |
| **Finish** | Target/Due date | TargetDate (preferred) or DueDate |
| **Predecessors** | Task dependencies | Relationship processing |
| **Resource Names** | Assigned resource | Custom.Owner (preferred) or System.AssignedTo |
| **Text1** | Work Item Type | Epic/Feature/User Story/Task/Bug |
| **Text2** | Work Item State | New/Active/Resolved/Closed |
| **Text3** | ADO URL | Direct clickable link to work item |
| **Number1** | ADO Work Item ID | Azure DevOps work item identifier |
| **Notes** | Work Item Description | System.Description from ADO |

### Benefits of Structured Approach
- **Easy Filtering**: Filter by Text1 (Type) or Text2 (State)
- **Easy Grouping**: Group by work item type or status
- **Easy Reporting**: Use Text/Number fields in Project reports
- **Clickable Links**: Click Text3 URLs to jump to Azure DevOps
- **Searchable IDs**: Find specific work items using Number1

## 🔄 Microsoft Project Import Guide

### Simple Import Process
1. Open **Microsoft Project**
2. Go to **File → Open**
3. Select your generated Excel file
4. Choose **Tasks** worksheet
5. Follow the **Import Wizard**:
   - All fields map directly to standard Project fields
   - Import wizard will automatically recognize field types
   - Click **Finish** - Done!

### Field Mapping (Automatic)
The import wizard will automatically recognize:
- ✅ **Unique ID** → **Unique ID**
- ✅ **Name** → **Name**
- ✅ **Outline Level** → **Outline Level**
- ✅ **% Complete** → **% Complete**
- ✅ **Start** → **Start**
- ✅ **Finish** → **Finish**
- ✅ **Predecessors** → **Predecessors**
- ✅ **Resource Names** → **Resource Names**
- ✅ **Text1** → **Text1** (Work Item Type)
- ✅ **Text2** → **Text2** (Work Item State)
- ✅ **Text3** → **Text3** (ADO URL)
- ✅ **Number1** → **Number1** (ADO ID)
- ✅ **Notes** → **Notes** (Description)

### Post-Import Usage Examples

#### Filtering by ADO Data:
- **Filter by Work Item Type**: Use Text1 field to show only "User Story" items
- **Filter by State**: Use Text2 field to show only "Active" items
- **Filter by ADO ID Range**: Use Number1 field for specific work item ranges

#### Grouping and Organization:
- **Group by Type**: Group by Text1 to organize Epics, Features, Stories, Tasks
- **Group by State**: Group by Text2 to organize by New, Active, Resolved, Closed
- **Group by Assignee**: Group by Resource Names to organize by team member

#### Reporting and Analysis:
- **Type Breakdown**: Use Text1 in reports to show work item distribution
- **Status Reports**: Use Text2 in reports to show completion status
- **Cross-Reference**: Use Number1 to cross-reference with Azure DevOps

### Benefits of Structured Approach
- ✅ **Native Compatibility**: Core fields work with all Project features
- ✅ **Rich Filtering**: Filter by type, state, ID, or any combination
- ✅ **Flexible Grouping**: Group data any way you need
- ✅ **Standard Reports**: Use all built-in Project reporting features
- ✅ **Clickable Links**: Click Text3 URLs to jump to Azure DevOps
- ✅ **Export Compatibility**: Export back to Excel without issues

## ⚙️ Configuration Options

### Core Settings
```powershell
# Connection settings
$ORGANIZATION_URL = "https://dev.azure.com/your-org"
$PROJECT_NAME = "Your-Project"
$PERSONAL_ACCESS_TOKEN = "your-pat"

# Output settings  
$OUTPUT_EXCEL_PATH = "C:\temp\export.xlsx"
$BATCH_SIZE = 200
$HOURS_PER_DAY = 8
```

### Advanced Settings
```powershell
# Enable relationship processing
$PROCESS_RELATIONSHIPS = $true

# API timeout (seconds)
$API_TIMEOUT = 60

# Enable debug logging
$ENABLE_DEBUG_LOGGING = $false
```

### Optimized Field Selection
The tool now uses only essential fields for native Project compatibility:
```powershell
$FIELDS_TO_FETCH = @(
    "System.Id",
    "System.Title", 
    "System.WorkItemType",
    "System.State",
    "System.AssignedTo",
    "System.Description",
    "Microsoft.VSTS.Scheduling.OriginalEstimate",
    "Microsoft.VSTS.Scheduling.RemainingWork",
    "Microsoft.VSTS.Scheduling.CompletedWork",
    "Microsoft.VSTS.Scheduling.StartDate",
    "Microsoft.VSTS.Scheduling.TargetDate",
    "Microsoft.VSTS.Scheduling.DueDate",
    "Custom.Owner"  # Add if you have Owner field in your ADO setup
)
```

### Date Field Priority Logic
- **Start**: Uses `Microsoft.VSTS.Scheduling.StartDate`
- **Finish**: Uses `Microsoft.VSTS.Scheduling.TargetDate` (revised due date) if available, otherwise `Microsoft.VSTS.Scheduling.DueDate` (original due date)
- **Resource Names**: Uses `Custom.Owner` if available, otherwise `System.AssignedTo`

### All ADO Metadata in Notes Field
All additional Azure DevOps information is automatically consolidated into the Notes field:
- Work Item ID and direct URL
- Work Item Type and State
- Clickable link back to Azure DevOps
- No field mapping conflicts

## 🎨 Customization Examples

### Export from Multiple Area Paths
```powershell
$WIQL_QUERY = @"
SELECT [System.Id] 
FROM WorkItems 
WHERE [System.TeamProject] = '$PROJECT_NAME' 
AND ([System.AreaPath] UNDER '$PROJECT_NAME\Team A' 
     OR [System.AreaPath] UNDER '$PROJECT_NAME\Team B')
"@
```

### Export from Specific Iteration
```powershell
$WIQL_QUERY = @"
SELECT [System.Id] 
FROM WorkItems 
WHERE [System.TeamProject] = '$PROJECT_NAME' 
AND [System.IterationPath] = '$PROJECT_NAME\Sprint 1'
"@
```

### Export by Date Range
```powershell
$WIQL_QUERY = @"
SELECT [System.Id] 
FROM WorkItems 
WHERE [System.TeamProject] = '$PROJECT_NAME' 
AND [System.CreatedDate] >= '2024-01-01'
AND [System.CreatedDate] <= '2024-12-31'
"@
```

## 🐛 Troubleshooting

### Common Issues

#### Connection Errors
- **401 Unauthorized**: Check your Personal Access Token
- **403 Forbidden**: Ensure your PAT has "Work Items (Read)" permission
- **404 Not Found**: Verify organization URL and project name

#### Export Errors
- **No work items found**: Check your WIQL query and area path
- **Timeout errors**: Reduce batch size in configuration
- **Memory issues**: Export smaller batches or specific work item types

#### Import Errors
- **Field mapping issues**: Refer to the detailed mapping guide in `C:\temp\MSProject_EXACT_Field_Mapping_Guide.txt`
- **Hierarchy problems**: Ensure Outline Level field is correctly mapped
- **Missing dependencies**: Verify Predecessors field mapping

### Getting Help
1. Check the generated log output for specific error messages
2. Enable debug logging: `$ENABLE_DEBUG_LOGGING = $true`
3. Verify your Azure DevOps permissions
4. Test with a smaller dataset first

## 📁 File Structure

```
ADO-MPP/
├── export-ado-workitems.ps1    # Main production script
├── config.ps1                  # Configuration file
├── README.md                   # This documentation
└── [Output Files]
    ├── AzureDevOpsExport_ProjectImport.xlsx
    ├── AzureDevOpsExport_SIMPLIFIED.xlsx
    └── MSProject_EXACT_Field_Mapping_Guide.txt
```

## 🔒 Security Considerations

- **⚠️ Never commit PAT tokens**: The `config.ps1` file is automatically ignored by git
- **Use environment variables**: Consider `$env:ADO_PAT` for CI/CD scenarios
- **Minimal permissions**: Only grant "Work Items (Read)" permission to your PAT
- **Token rotation**: Regularly rotate your Personal Access Tokens
- **Template usage**: Always copy from `config.example.ps1` to create your `config.ps1`

## 📈 Features

- ✅ **Hierarchical Export**: Maintains Epic > Feature > User Story > Task > Bug > Dependency > Milestone structure
- ✅ **Task Dependencies**: Exports predecessor/successor relationships
- ✅ **Batch Processing**: Handles large datasets efficiently
- ✅ **Error Handling**: Robust error handling and logging
- ✅ **Flexible Configuration**: Easy to customize for different projects
- ✅ **Multiple Output Formats**: Excel primary, CSV fallback
- ✅ **Microsoft Project Compatible**: Direct import support
- ✅ **Production Ready**: Comprehensive configuration and documentation

## 📝 License

This tool is provided as-is for internal use. Modify and distribute according to your organization's policies.

## 🔄 Version History

- **v2.0** - Production release with external configuration
- **v1.0** - Initial working version with embedded configuration
