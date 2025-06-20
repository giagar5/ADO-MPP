# ADO2MPP - Azure DevOps to Microsoft Project Bridge

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1+-blue.svg)](https://github.com/PowerShell/PowerShell)
[![Azure DevOps](https://img.shields.io/badge/Azure%20DevOps-Compatible-0078d4.svg)](https://azure.microsoft.com/en-us/products/devops/)
[![Microsoft Project](https://img.shields.io/badge/Microsoft%20Project-Compatible-217346.svg)](https://www.microsoft.com/en-us/microsoft-365/project/project-management-software)

## 🚀 **Quick Start - Production Ready**

### **Main Export (Azure DevOps → Microsoft Project)**
```cmd
run-main-export.bat
```
**OR**
```powershell
.\export-ado-workitems.ps1 -ConfigPath "config\config.ps1"
```

### **Critical Timeline Export (Azure DevOps → Office Timeline Expert)**
```cmd
run-critical-timeline-export.bat
```
**OR**
```powershell
.\export-critical-timeline.ps1 -ConfigPath "config\config.ps1"
```

## 📁 **Simplified Project Structure**

```
ADO2MPP/
├── 📄 export-ado-workitems.ps1          # Main export script
├── 📄 export-critical-timeline.ps1      # Critical timeline export
├── 📄 run-main-export.bat               # Main export launcher
├── 📄 run-critical-timeline-export.bat  # Timeline export launcher
├── � README.md                         # This file
├── �📁 config/                           # Configuration files
│   ├── config.ps1                       # Main configuration
│   └── config.example.ps1               # Example configuration
└── 📁 utils/                            # Utility scripts (optional)
```

## ⚙️ **Setup (First Time)**

1. **Configure**: Copy `config\config.example.ps1` to `config\config.ps1`
2. **Edit**: Update `config\config.ps1` with your Azure DevOps details
3. **Run**: Execute either batch file or PowerShell script directly

## 🎯 **Key Features**

- **🎯 Main Export**: Complete Azure DevOps work items → Microsoft Project
- **📊 Critical Timeline**: Critical milestones → Office Timeline Expert
- **� Easy Launch**: Simple batch files for one-click execution
- **🧹 Clean Structure**: Minimal folders, maximum clarity
- **⚙️ Simple Setup**: Just configure and run

**Simplified for production use - only essential files included!**

Project reorganized for production: 2025-06-20
