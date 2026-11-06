@echo off
setlocal EnableExtensions

rem Certifies the checked-out commit in Excel: builds a development workbook in
rem a new bundle folder next to the checkout, compiles it, runs
rem Test_EXCEL_UI_RunReleaseCertification and collects the evidence.
rem Double-click it, or run it from a command prompt; arguments are passed on
rem to Certify-EXCEL_UI.ps1. See docs\EXCEL_CERTIFICATION.md.

set "SCRIPT=%~dp0Certify-EXCEL_UI.ps1"

if not exist "%SCRIPT%" (
    echo ERROR: PowerShell script not found:
    echo   %SCRIPT%
    echo.
    pause
    exit /b 1
)

cd /d "%~dp0.."
echo VBA Excel UI certification
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" %*
set "CODE=%ERRORLEVEL%"
echo.
pause
exit /b %CODE%
