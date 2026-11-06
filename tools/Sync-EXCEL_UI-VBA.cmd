@echo off
setlocal EnableExtensions
set "REPO_ROOT=%~dp0.."
cd /d "%REPO_ROOT%"

set "SCRIPT=%~dp0Sync-EXCEL_UI-VBA.ps1"

if not exist "%SCRIPT%" (
    echo ERROR: PowerShell script not found:
    echo   %SCRIPT%
    echo.
    pause
    exit /b 1
)

rem Without a dropped workbook, a blank development workbook is created in
rem the ignored test-results folder, replaced on every run. A dropped
rem workbook is synchronized as it is; the script asks for the configuration.
set "WORKBOOK=%~1"
set "BUILD="
if not defined WORKBOOK (
    set "WORKBOOK=%CD%\test-results\EXCEL_UI_Development.xlsm"
    set "BUILD=-Mode Development -NewWorkbook"
)

echo VBA-EXCEL_UI workbook synchronization
echo.

if not defined BUILD if not exist "%WORKBOOK%" (
    echo ERROR: Workbook not found:
    echo   %WORKBOOK%
    echo.
    pause
    exit /b 1
)

echo Repository:
echo   %CD%
echo Workbook:
echo   %WORKBOOK%
echo.
if defined BUILD (
    echo Creating a blank development workbook and importing the modules
    echo of src and test...
) else (
    echo Synchronizing the VBA modules of the dropped workbook...
)
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" -WorkbookPath "%WORKBOOK%" -RepoRoot "%REPO_ROOT%" %BUILD%
set "RC=%ERRORLEVEL%"

echo.
if "%RC%"=="0" (
    echo SUCCESS: VBA synchronization completed.
    if defined BUILD (
        echo Open the workbook and run Debug - Compile VBAProject:
        echo   %WORKBOOK%
    )
) else if "%RC%"=="2" (
    echo NOT CHANGED: nothing was imported; the workbook was not changed.
) else (
    echo FAILED: VBA synchronization returned exit code %RC%.
)
echo.
pause
exit /b %RC%
