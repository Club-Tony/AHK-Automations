@echo off
setlocal

echo ======================================
echo   ItemVar Discovery Tool
echo ======================================
echo.

:: Prompt for PK#
set /p ASSETID="Enter PK# (e.g., PK438893): "

:: Check if cookies.txt exists
if not exist "%~dp0cookies.txt" (
    echo ERROR: cookies.txt not found
    pause
    exit /b 1
)

:: Run the PowerShell script (opens notepad with results, then closes)
powershell -ExecutionPolicy Bypass -File "%~dp0discover_itemvars.ps1" -AssetId "%ASSETID%" -CookieFile "%~dp0cookies.txt" -OutputFile "%~dp0itemvar_discovery.txt"
