# Azure DevOps to Microsoft Project Export Tool

A PowerShell-based tool that exports work items from Azure DevOps to Excel format compatible with Microsoft Project import, maintaining hierarchical structure and task dependencies.

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

The tool generates an Excel file with the following columns:

| Column | Description |
|--------|-------------|
| **Unique ID** | Task sequence number for Microsoft Project |
| **Name** | Work item title |
| **Duration** | Calculated from effort estimates |
| **Start** | Start date from Azure DevOps |
| **Finish** | Target date from Azure DevOps |
| **Predecessors** | Task dependencies (region-specific separator) |
| **Resource Names** | Assigned team members |
| **Outline Level** | Hierarchy level (1=Epic, 2=Feature, 3=Story, 4=Task) |
| **ADO ID** | Azure DevOps work item ID (dedicated field) |
| **Work Item Type** | Work item type (Epic, Feature, User Story, etc.) |
| **Work Item State** | Current work item state (New, Active, Done, Closed, etc.) |
| **Area Path** | Azure DevOps area path (organizational hierarchy) |
| **Tags** | Work item tags (comma-separated) |
| **ADO Link** | Direct link to Azure DevOps work item |

## 🔄 Microsoft Project Import Guide

### Method 1: Using Excel Import
1. Open **Microsoft Project**
2. Go to **File → Open**
3. Select your generated Excel file
4. Choose **Excel Workbook** file type
5. Follow the **Import Wizard**:
   - Select the **Project Import** worksheet
   - In the **Map** step, configure **Task Mapping**:
     - **Unique ID** → **ID** (or Unique ID)
     - **Name** → **Name**
     - **Duration** → **Duration**
     - **Start** → **Start**
     - **Finish** → **Finish**
     - **Predecessors** → **Predecessors**
     - **Resource Names** → **Resource Names**
     - **Outline Level** → **Outline Level**
     - **ADO ID** → **Number1** (optional, for Azure DevOps work item ID)
     - **Work Item Type** → **Text1** (optional, for work item type)
     - **Work Item State** → **Text2** (optional, for work item state)
     - **Area Path** → **Text3** (optional, for Azure DevOps area path)
     - **Tags** → **Text4** (optional, for work item tags)
     - **ADO Link** → **Text5** (optional, for Azure DevOps links)
   - Leave **Resource Mapping** and **Assignment Mapping** blank
6. Click **Finish**

### Method 2: Using CSV Import (Fallback)
If Excel import fails, the tool automatically creates a CSV version:
1. Open **Microsoft Project**
2. Go to **File → Open**
3. Change file type to **CSV**
4. Select the `.csv` file
5. Follow the same field mapping as above

### Troubleshooting Import Issues
- **"Map does not map any fields"**: Ensure you've configured at least the basic mappings (ID, Name, Outline Level) in the Task Mapping tab
- **Missing hierarchy**: Verify the Outline Level field is mapped correctly
- **No dependencies**: Check that Predecessors field is mapped to enable task relationships

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

### Custom Field Selection
Add or remove fields in the `$FIELDS_TO_FETCH` array:
```powershell
$FIELDS_TO_FETCH = @(
    "System.Id",
    "System.Title", 
    "System.WorkItemType",
    "System.State",
    "System.AssignedTo",
    "System.Description",        # Add description
    "System.Tags",              # Add tags
    "Microsoft.VSTS.Common.Priority"  # Add priority
)
```

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

- ✅ **Hierarchical Export**: Maintains Epic > Feature > User Story > Task structure
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
