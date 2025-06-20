@echo off
REM ADO2MPP Main Export Batch File
REM This batch file runs the main PowerShell script to export Azure DevOps work items to Microsoft Project

echo ========================================
echo ADO2MPP - Azure DevOps to Microsoft Project Export
echo ========================================
echo.

REM Check if PowerShell is available
powershell -Command "Write-Host 'PowerShell is available'" >nul 2>&1
if errorlevel 1 (
    echo ERROR: PowerShell is not available or not in PATH
    echo Please ensure PowerShell is installed and accessible
    pause
    exit /b 1
)

REM Change to the script directory
cd /d "%~dp0"

REM Run the PowerShell script with execution policy bypass
echo Running ADO2MPP Export...
echo.
powershell -ExecutionPolicy Bypass -File "export-ado-workitems.ps1" -ConfigPath ".\config.ps1"

REM Check if the script ran without errors
if errorlevel 1 (
    echo.
    echo ERROR: Script execution failed
    echo Please check the error messages above
) else (
    echo.
    echo SUCCESS: ADO2MPP Export completed successfully
    echo Check the output folder for your Excel file
)

echo.
echo Press any key to exit...
pause >nul
