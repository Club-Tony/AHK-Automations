@echo off
setlocal enabledelayedexpansion

echo ======================================
echo   DSRF-Export Setup
echo ======================================
echo.

set "SCRIPT_DIR=%~dp0"
set "IMPORT_DIR=%SCRIPT_DIR%import"
set "SQLITE3=%SCRIPT_DIR%sqlite3.exe"
set "ALL_GOOD=1"

:: ========================================
:: Step 1: Check/Download sqlite3.exe
:: ========================================
echo [1/3] Checking sqlite3.exe...

if exist "%SQLITE3%" (
    echo       Found: sqlite3.exe already exists.
) else (
    echo       sqlite3.exe not found. Downloading...
    echo.

    powershell -ExecutionPolicy Bypass -Command ^
        "$ErrorActionPreference = 'Stop'; " ^
        "try { " ^
        "  $zipUrl = 'https://www.sqlite.org/2024/sqlite-tools-win-x64-3470200.zip'; " ^
        "  $zipFile = Join-Path $env:TEMP 'sqlite-tools.zip'; " ^
        "  $extractDir = Join-Path $env:TEMP 'sqlite-tools-extract'; " ^
        "  Write-Host '       Downloading from sqlite.org...'; " ^
        "  [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; " ^
        "  Invoke-WebRequest -Uri $zipUrl -OutFile $zipFile -UseBasicParsing; " ^
        "  Write-Host '       Extracting...'; " ^
        "  if (Test-Path $extractDir) { Remove-Item $extractDir -Recurse -Force }; " ^
        "  Expand-Archive -Path $zipFile -DestinationPath $extractDir -Force; " ^
        "  $sqlite3File = Get-ChildItem $extractDir -Recurse -Filter 'sqlite3.exe' | Select-Object -First 1; " ^
        "  if (-not $sqlite3File) { throw 'sqlite3.exe not found in downloaded archive' }; " ^
        "  Copy-Item $sqlite3File.FullName '%SQLITE3%' -Force; " ^
        "  Remove-Item $zipFile -Force -ErrorAction SilentlyContinue; " ^
        "  Remove-Item $extractDir -Recurse -Force -ErrorAction SilentlyContinue; " ^
        "  Write-Host '       sqlite3.exe installed successfully.'; " ^
        "} catch { " ^
        "  Write-Host ('       ERROR: ' + $_.Exception.Message); " ^
        "  Write-Host ''; " ^
        "  Write-Host '       Manual download: https://www.sqlite.org/download.html'; " ^
        "  Write-Host '       Get: sqlite-tools-win-x64 zip'; " ^
        "  Write-Host '       Extract sqlite3.exe to this folder.'; " ^
        "  exit 1; " ^
        "}"

    if !ERRORLEVEL! neq 0 (
        set "ALL_GOOD=0"
    )
    echo.
)

:: ========================================
:: Step 2: Create import subfolder
:: ========================================
echo [2/3] Checking import folder...

if exist "%IMPORT_DIR%\" (
    echo       Found: import\ folder already exists.
) else (
    mkdir "%IMPORT_DIR%"
    if !ERRORLEVEL! neq 0 (
        echo       ERROR: Could not create import\ folder.
        set "ALL_GOOD=0"
    ) else (
        echo       Created: import\ folder.
    )
)

:: ========================================
:: Step 3: Verify Firefox profile exists
:: ========================================
echo [3/3] Checking Firefox profile...

set "FF_PROFILES=%APPDATA%\Mozilla\Firefox\Profiles"
if exist "%FF_PROFILES%" (
    echo       Found: Firefox profiles directory exists.
) else (
    echo       WARNING: Firefox profiles not found at %FF_PROFILES%
    echo       Make sure Firefox is installed and you've logged into Intra.
)

:: ========================================
:: Summary
:: ========================================
echo.
echo ======================================

if "%ALL_GOOD%"=="1" (
    if exist "%SQLITE3%" (
        echo   Setup complete!
        echo ======================================
        echo.
        echo   You can now run DSRF-Export.bat
        echo.
        echo ======================================
        echo   WorldShip XML Auto-Import Setup
        echo ======================================
        echo.
        echo   To use XML Auto-Import in WorldShip:
        echo.
        echo   1. Open UPS WorldShip
        echo   2. Go to Import/Export ^> Auto Import
        echo   3. Set import directory to:
        echo      %IMPORT_DIR%
        echo   4. Enable auto-processing
        echo.
        echo   After setup, WorldShip will automatically
        echo   pick up shipment XML files from the
        echo   import\ folder when you run DSRF-Export.bat.
        echo.
        echo   Alternatively, use the CSV file generated
        echo   in the same folder for manual batch import.
        echo ======================================
    ) else (
        echo   Setup incomplete - sqlite3.exe missing
        echo ======================================
        echo.
        echo   Please download sqlite3.exe manually:
        echo   https://www.sqlite.org/download.html
        echo   Extract to: %SCRIPT_DIR%
    )
) else (
    echo   Setup had errors - see above
    echo ======================================
)

echo.
pause
