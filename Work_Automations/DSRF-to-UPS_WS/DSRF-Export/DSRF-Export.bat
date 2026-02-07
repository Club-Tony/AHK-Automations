@echo off
setlocal enabledelayedexpansion

echo ======================================
echo   DSRF to WorldShip Export Tool
echo ======================================
echo.

set "SCRIPT_DIR=%~dp0"
set "SQLITE3=%SCRIPT_DIR%sqlite3.exe"
set "IMPORT_DIR=%SCRIPT_DIR%import"
set "PS_SCRIPT=%SCRIPT_DIR%DSRF-Export.ps1"

:: ========================================
:: Pre-flight checks
:: ========================================

:: Check PowerShell script exists
if not exist "%PS_SCRIPT%" (
    echo ERROR: DSRF-Export.ps1 not found.
    echo Expected at: %PS_SCRIPT%
    echo.
    pause
    exit /b 1
)

:: Check sqlite3
if not exist "%SQLITE3%" (
    echo ERROR: sqlite3.exe not found.
    echo.
    echo Run DSRF-ExportSetup.bat first to install dependencies.
    echo.
    pause
    exit /b 1
)

:: Check import folder
if not exist "%IMPORT_DIR%\" (
    mkdir "%IMPORT_DIR%" 2>nul
)

:: ========================================
:: Auto-detect PK# from Firefox history
:: ========================================
set "DETECTED_PK="

for /f "usebackq delims=" %%i in (`powershell -ExecutionPolicy Bypass -File "%PS_SCRIPT%" -DetectOnly`) do (
    set "DETECTED_PK=%%i"
)

:: ========================================
:: Prompt for PK#
:: ========================================
if defined DETECTED_PK (
    echo Detected: !DETECTED_PK! from Firefox browsing history
    echo.
    set /p CONFIRM="Use !DETECTED_PK!? [Y/n]: "
    if /i "!CONFIRM:~0,1!"=="n" (
        set /p ASSETID="Enter PK# (e.g., PK123456): "
    ) else (
        set "ASSETID=!DETECTED_PK!"
    )
) else (
    set /p ASSETID="Enter PK# (e.g., PK123456): "
)

if "%ASSETID%"=="" (
    echo ERROR: No PK# entered.
    pause
    exit /b 1
)

:: ========================================
:: Run PowerShell export script
:: ========================================
echo.
powershell -ExecutionPolicy Bypass -File "%PS_SCRIPT%" -AssetId "%ASSETID%"

if %ERRORLEVEL% neq 0 (
    echo.
    echo Script encountered an error. See messages above.
)

echo.
pause
