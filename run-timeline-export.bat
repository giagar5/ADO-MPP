@echo off
echo ========================================
echo Timeline Export for Office Timeline Expert
echo ========================================
echo.
echo This script exports Milestones and Dependencies to Excel
echo for Office Timeline Expert import.
echo.
echo Choose export mode:
echo [1] Quick Export - Priority items (MG1, DQT-Phase1, Modernization, DataProduct)
echo [2] Export ALL Milestones and Dependencies  
echo [3] Custom tags (you will be prompted)
echo [4] Exit
echo.
echo Press ENTER for Quick Export or choose (1-4):
set /p choice="Your choice: "

REM Default to option 1 if user just pressed ENTER
if "%choice%"=="" set choice=1

if "%choice%"=="1" (
    echo.
    echo Running Quick Export with priority tags...
    powershell.exe -ExecutionPolicy Bypass -File ".\export-critical-timeline.ps1" -ConfigPath "config\config.ps1" -DebugMode
) else if "%choice%"=="2" (
    echo.
    echo Exporting ALL Milestones and Dependencies...
    powershell.exe -ExecutionPolicy Bypass -File ".\export-critical-timeline.ps1" -ConfigPath "config\config.ps1" -ExportAll -DebugMode
) else if "%choice%"=="3" (
    echo.
    set /p customtags="Enter tags separated by commas (e.g., Cloudera,Teradata,DataProduct): "
    echo Exporting items with custom tags: %customtags%
    powershell.exe -ExecutionPolicy Bypass -Command "& '.\export-critical-timeline.ps1' -ConfigPath 'config\config.ps1' -PriorityTags '%customtags%'.Split(',') -DebugMode"
) else if "%choice%"=="4" (
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
