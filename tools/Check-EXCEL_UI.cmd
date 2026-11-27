@echo off
setlocal EnableExtensions

rem Runs certification, the regression runners and the demo checks on the
rem checked-out commit in Excel, and writes one result bundle with a summary
rem in a new folder next to the checkout. Double-click it, or run it from a
rem command prompt; arguments are passed on to Check-EXCEL_UI.ps1.
rem This launcher passes -Unattended itself, so the run needs no key press and
rem no click and the window closes at the end; the summary is in the bundle.
rem To run attended, start Check-EXCEL_UI.ps1 from a PowerShell prompt without
rem -Unattended. See docs\EXCEL_CHECKS.md.

set "SCRIPT=%~dp0Check-EXCEL_UI.ps1"

rem -Unattended is passed once, even when it is also among the arguments.
set "UNATTENDED_ARG=-Unattended"
:scan
if "%~1"=="" goto scanned
if /i "%~1"=="-Unattended" set "UNATTENDED_ARG="
shift /1
goto scan
:scanned

if not exist "%SCRIPT%" (
    echo ERROR: PowerShell script not found:
    echo   %SCRIPT%
    exit /b 1
)

cd /d "%~dp0.."
echo VBA Excel UI checks
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" %UNATTENDED_ARG% %*
set "CODE=%ERRORLEVEL%"
exit /b %CODE%
