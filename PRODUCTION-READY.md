\
# PRODUCTION READY - Azure DevOps to Microsoft Project Export Tool

## 🎉 Production Setup Complete!

Your Azure DevOps to Microsoft Project export tool has been prepared for production use with the following improvements:

### ✅ **What's New in Production Version**

1. **External Configuration**: All settings moved to `config.ps1` for easy customization
2. **Enhanced Error Handling**: Robust error handling and detailed logging
3. **Flexible Parameters**: Command-line parameter support for different use cases
4. **Regional Settings Support**: Configurable delimiter and formatting for international use
5. **Clean Workspace**: Removed test files and debug artifacts
6. **Comprehensive Documentation**: Complete README and regional settings guide
7. **User-Friendly Launcher**: Interactive script for easy execution

### 📁 **Current File Structure**

```
ADO-MPP/
├── 🚀 run-export.ps1                           # Interactive launcher (START HERE)
├── ⚙️ config.ps1                              # Configuration file (EDIT THIS)
├── 🔧 export-ado-workitems.ps1               # Main production script
├── 📋 README.md                               # Complete documentation
├── 🌍 REGIONAL-SETTINGS-GUIDE.md              # Regional settings documentation
├── 🛠️ create_msproject_import_template.ps1   # Helper for mapping instructions
└── 📜 PRODUCTION-READY.md                     # This file
```

### 🚀 **Quick Start for Production**

#### Step 1: Configure Your Settings
Edit `config.ps1` and update these key values:
```powershell
$ORGANIZATION_URL = "https://dev.azure.com/your-organization"
$PROJECT_NAME = "Your-Project-Name"  
$PERSONAL_ACCESS_TOKEN = "your-pat-token-here"

# Regional settings for international compatibility
$REGIONAL_FORMAT = "European"  # Options: 'Auto', 'US', 'European', 'Custom'
```

#### Step 2: Run the Export
**Easiest way:** Double-click `run-export.ps1` for interactive menu

**Command line:**
```powershell
.\export-ado-workitems.ps1                    # Default export
.\export-ado-workitems.ps1 -OutputPath "C:\exports\my-project.xlsx"
.\export-ado-workitems.ps1 -AreaPath "MyProject\Team Alpha"
```

#### Step 3: Import to Microsoft Project
1. Open Microsoft Project
2. File → Open → Select your Excel file
3. Follow Import Wizard with field mappings from README.md

### 🔧 **Configuration Examples**

#### Export All Work Items from Project
```powershell
$WIQL_QUERY = "SELECT [System.Id] FROM WorkItems WHERE [System.TeamProject] = '$PROJECT_NAME'"
```

#### Export from Specific Team/Area
```powershell
$WIQL_QUERY = "SELECT [System.Id] FROM WorkItems WHERE [System.TeamProject] = '$PROJECT_NAME' AND [System.AreaPath] UNDER '$PROJECT_NAME\Your-Team'"
```

#### Export Specific Work Item Types
```powershell
$WIQL_QUERY = "SELECT [System.Id] FROM WorkItems WHERE [System.TeamProject] = '$PROJECT_NAME' AND [System.WorkItemType] IN ('Epic', 'Feature', 'User Story')"
```

### 📊 **Expected Output**

The tool will generate:
- **Excel file** with hierarchical work items (Epic → Feature → User Story → Task)
- **Task dependencies** from Azure DevOps relationships
- **Microsoft Project compatible** column structure
- **Detailed logging** during export process

### 🛡️ **Security & Best Practices**

1. **Personal Access Token**: 
   - Create with minimal permissions (Work Items: Read)
   - Store securely, don't commit to version control
   - Rotate regularly

2. **Configuration Management**:
   - Keep sensitive settings in separate config files
   - Use environment variables for CI/CD scenarios
   - Document area paths and queries for team use

3. **Data Handling**:
   - Verify output before sharing
   - Be mindful of confidential work item data
   - Clean up temporary files regularly

### 🎯 **Performance Optimization**

- **Batch Size**: Adjust `$BATCH_SIZE` (default: 200) for your network
- **Field Selection**: Remove unused fields from `$FIELDS_TO_FETCH`
- **Query Optimization**: Use specific area paths and date ranges
- **Large Datasets**: Consider exporting in smaller chunks

### 🔍 **Troubleshooting Quick Reference**

| Issue | Solution |
|-------|----------|
| 401 Unauthorized | Check Personal Access Token |
| 403 Forbidden | Verify PAT permissions |
| No work items found | Check WIQL query and area path |
| Excel import fails | Use CSV fallback or check field mapping |
| Timeout errors | Reduce batch size |
| Missing hierarchy | Verify Outline Level mapping |

### 📞 **Support & Maintenance**

#### For Configuration Issues:
1. Check `config.ps1` syntax
2. Validate Azure DevOps connection
3. Test with simple WIQL query first

#### For Export Issues:
1. Enable debug logging: `$ENABLE_DEBUG_LOGGING = $true`
2. Check Azure DevOps permissions
3. Try smaller datasets

#### For Import Issues:
1. Review field mapping guide in README.md
2. Use simplified Excel file first
3. Fallback to CSV import if needed

### 🔄 **Regular Maintenance Tasks**

- **Monthly**: Rotate Personal Access Tokens
- **As Needed**: Update area paths and queries for new teams
- **Before Major Exports**: Test with small dataset first
- **After Azure DevOps Changes**: Verify field names and permissions

### 📈 **Feature Roadmap Ideas**

Future enhancements you could consider:
- Configuration templates for different teams
- Automated scheduling/CI-CD integration
- Custom field mapping configurations
- Integration with other project management tools
- Bulk update capabilities back to Azure DevOps

---

## 🔧 **RECENT FIX: Microsoft Project Excel Import Issue**

### ✅ **Issue Resolved: "Map does not map any fields" Error**

**Problem**: Microsoft Project Import Wizard was showing "Map does not map any fields" error because Excel column headers didn't match Microsoft Project's expected field names.

**Solution Applied**: Updated Excel export to use Microsoft Project standard field names:
- Changed `"ID"` → `"Unique ID"` (CRITICAL for import recognition)
- Added Microsoft Project standard fields: `"Work"`, `"% Complete"`, `"Priority"`, `"Task Mode"`, `"WBS"`
- Azure DevOps data now mapped to `"Text1"` through `"Text5"` fields and `"Notes"`

### 📋 **New Field Mapping for Microsoft Project Import**

**Essential Fields (MUST be mapped):**
- `Unique ID` → Project Field: Unique ID
- `Name` → Project Field: Name  
- `Outline Level` → Project Field: Outline Level

**Recommended Fields:**
- `Duration` → Project Field: Duration
- `Start` → Project Field: Start
- `Finish` → Project Field: Finish
- `Predecessors` → Project Field: Predecessors
- `Resource Names` → Project Field: Resource Names
- `Work` → Project Field: Work
- `% Complete` → Project Field: % Complete

### 🎯 **Import Instructions Updated**

1. **Use the SIMPLIFIED file first**: `AzureDevOpsExport_ProjectImport_SIMPLIFIED.xlsx`
2. **Open Microsoft Project** → File → Open
3. **Import Wizard should now recognize fields automatically**
4. **Map the essential fields** (Unique ID, Name, Outline Level)
5. **Map optional fields** as needed

The Excel files now use Microsoft Project's expected column names, so the Import Wizard should automatically detect and suggest field mappings instead of showing the "no fields found" error.

---

## ✅ **Production Checklist**

- [x] Configuration externalized
- [x] Debug files removed
- [x] Error handling improved
- [x] Documentation completed
- [x] User-friendly launcher created
- [x] Security considerations documented
- [x] Performance optimization guidelines provided
- [x] Troubleshooting guide included

**Your Azure DevOps to Microsoft Project export tool is now production-ready!**

Start with `run-export.ps1` for the best user experience.
