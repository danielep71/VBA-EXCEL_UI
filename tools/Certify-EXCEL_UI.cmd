@echo off
setlocal EnableExtensions

rem Certifies the checked-out commit in Excel: builds a development workbook in
rem a new bundle folder next to the checkout, compiles it, runs
rem Test_EXCEL_UI_RunReleaseCertification and collects the evidence.
rem Double-click it, or run it from a command prompt; arguments are passed on
rem to Certify-EXCEL_UI.ps1, so -Unattended runs it without prompts and without
rem the final pause. See docs\EXCEL_CERTIFICATION.md.

set "SCRIPT=%~dp0Certify-EXCEL_UI.ps1"

rem With -Unattended among the arguments the run never prompts, so the window
rem is not held open at the end either.
set "PAUSE_AT_END=1"
:scan
if "%~1"=="" goto scanned
if /i "%~1"=="-Unattended" set "PAUSE_AT_END="
shift /1
goto scan
:scanned

if not exist "%SCRIPT%" (
    echo ERROR: PowerShell script not found:
    echo   %SCRIPT%
    echo.
    if defined PAUSE_AT_END pause
    exit /b 1
)

cd /d "%~dp0.."
echo VBA Excel UI certification
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" %*
set "CODE=%ERRORLEVEL%"
echo.
if defined PAUSE_AT_END pause
exit /b %CODE%
