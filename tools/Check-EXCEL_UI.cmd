@echo off
setlocal EnableExtensions

rem Runs certification, the regression runners and the demo checks on the
rem checked-out commit in Excel, and writes one result bundle with a summary
rem in a new folder next to the checkout. Double-click it, or run it from a
rem command prompt; arguments are passed on to Check-EXCEL_UI.ps1.
rem See docs\EXCEL_CHECKS.md.

set "SCRIPT=%~dp0Check-EXCEL_UI.ps1"

if not exist "%SCRIPT%" (
    echo ERROR: PowerShell script not found:
    echo   %SCRIPT%
    echo.
    pause
    exit /b 1
)

cd /d "%~dp0.."
echo VBA Excel UI checks
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" %*
set "CODE=%ERRORLEVEL%"
echo.
pause
exit /b %CODE%
