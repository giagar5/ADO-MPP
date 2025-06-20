@echo off
echo ========================================
echo Timeline Export for Office Timeline Expert
echo ========================================
echo.
echo This script exports Milestones and Dependencies to Excel
echo for Office Timeline Expert import.
echo.
echo Choose export mode:
echo [1] Critical Only - Items tagged as "Critical" or with critical keywords
echo [2] All Milestones and Dependencies
echo [3] Exit
echo.
echo Press ENTER for Critical Only or choose (1-3):
set /p choice="Your choice: "

REM Default to option 1 if user just pressed ENTER
if "%choice%"=="" set choice=1

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

if "%choice%"=="1" (
    echo.
    echo Exporting CRITICAL ONLY items...
    powershell.exe -ExecutionPolicy Bypass -File ".\export-critical-timeline.ps1" -ConfigPath ".\config.ps1" -PriorityTags "Critical"
) else if "%choice%"=="2" (
    echo.
    echo Exporting ALL Milestones and Dependencies...
    powershell.exe -ExecutionPolicy Bypass -File ".\export-critical-timeline.ps1" -ConfigPath ".\config.ps1" -ExportAll
) else if "%choice%"=="3" (
    echo Exiting...
    exit /b 0
) else (
    echo Invalid choice. Please run the script again.
    pause
    exit /b 1
)

echo.
echo ========================================
echo Export completed!
echo ========================================
echo.
echo The Excel file has been created in C:\temp\
echo You can now import it into PowerPoint Office Timeline Expert.
echo.
pause
