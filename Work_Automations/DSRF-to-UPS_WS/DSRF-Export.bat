@echo off
setlocal enabledelayedexpansion

echo ======================================
echo   DSRF to CSV Export Tool
echo ======================================
echo.

:: Prompt for PK#
set /p ASSETID="Enter PK# (e.g., PK438893): "

:: Check if cookies.txt exists
if not exist "%~dp0cookies.txt" (
    echo.
    echo ERROR: cookies.txt not found in script folder
    echo.
    echo Please create cookies.txt with your Intra session cookies.
    echo See README.txt for instructions.
    echo.
    pause
    exit /b 1
)

:: Read cookies from file (first line only)
set /p COOKIES=<"%~dp0cookies.txt"

:: Create temp files
set "TEMPJSON=%TEMP%\dsrf_response.json"
set "OUTPUTCSV=%~dp0dsrf_export.csv"

echo.
echo Fetching data for %ASSETID%...
echo.

:: Build the SQL query
set "SQL=declare @profileid int = 1; declare @assetid nvarchar(50) = '%ASSETID%'; SELECT TOP 1 ISNULL(iv148.ItemVarValue,'') as company, ISNULL(iv149.ItemVarValue,'') as name, ISNULL(iv150.ItemVarValue,'') as address1, ISNULL(iv151.ItemVarValue,'') as address2, ISNULL(iv152.ItemVarValue,'') as city, ISNULL(iv153.ItemVarValue,'') as state, ISNULL(iv154.ItemVarValue,'') as postal, ISNULL(iv155.ItemVarValue,'') as serviceType FROM Asset a LEFT JOIN assetitemvars iv148 ON iv148.assetid=a.assetid AND iv148.profileid=1 AND iv148.itemvarid=148 LEFT JOIN assetitemvars iv149 ON iv149.assetid=a.assetid AND iv149.profileid=1 AND iv149.itemvarid=149 LEFT JOIN assetitemvars iv150 ON iv150.assetid=a.assetid AND iv150.profileid=1 AND iv150.itemvarid=150 LEFT JOIN assetitemvars iv151 ON iv151.assetid=a.assetid AND iv151.profileid=1 AND iv151.itemvarid=151 LEFT JOIN assetitemvars iv152 ON iv152.assetid=a.assetid AND iv152.profileid=1 AND iv152.itemvarid=152 LEFT JOIN assetitemvars iv153 ON iv153.assetid=a.assetid AND iv153.profileid=1 AND iv153.itemvarid=153 LEFT JOIN assetitemvars iv154 ON iv154.assetid=a.assetid AND iv154.profileid=1 AND iv154.itemvarid=154 LEFT JOIN assetitemvars iv155 ON iv155.assetid=a.assetid AND iv155.profileid=1 AND iv155.itemvarid=155 WHERE a.AssetID=@assetid AND a.ProfileID=@profileid"

:: Write JSON body to temp file for curl
set "BODYJSON=%TEMP%\dsrf_body.json"
echo {"Sql":"%SQL%"} > "%BODYJSON%"

:: Call API using curl (built into Windows 10+)
curl.exe -s -X POST ^
  -H "Content-Type: application/json" ^
  -H "Cookie: %COOKIES%" ^
  -d @"%BODYJSON%" ^
  "https://amazonmailservices.us.spsprod.net/IntraWeb/api/automation/then/executeQuery" ^
  -o "%TEMPJSON%"

if %ERRORLEVEL% neq 0 (
    echo ERROR: API call failed (curl error)
    echo Make sure you have an active Intra session and valid cookies.
    pause
    exit /b 1
)

:: Check if response file exists and has content
if not exist "%TEMPJSON%" (
    echo ERROR: No response received from API
    pause
    exit /b 1
)

:: Parse JSON and output CSV using PowerShell
powershell -Command ^
  "try { ^
     $json = Get-Content '%TEMPJSON%' -Raw | ConvertFrom-Json; ^
     if ($json -eq $null -or $json.Count -eq 0) { ^
       Write-Host 'ERROR: No data returned for this PK#'; ^
       exit 1; ^
     } ^
     $row = $json[0]; ^
     $csv = 'Name,Company,Address1,Address2,City,State,Postal,ServiceType'; ^
     $csv += [Environment]::NewLine; ^
     $name = if ($row.name) { $row.name } else { '' }; ^
     $company = if ($row.company) { $row.company } else { '' }; ^
     $addr1 = if ($row.address1) { $row.address1 } else { '' }; ^
     $addr2 = if ($row.address2) { $row.address2 } else { '' }; ^
     $city = if ($row.city) { $row.city } else { '' }; ^
     $state = if ($row.state) { $row.state } else { '' }; ^
     $postal = if ($row.postal) { $row.postal } else { '' }; ^
     $service = if ($row.serviceType) { $row.serviceType } else { '' }; ^
     $csv += '\"' + $name + '\",\"' + $company + '\",\"' + $addr1 + '\",\"' + $addr2 + '\",\"' + $city + '\",\"' + $state + '\",\"' + $postal + '\",\"' + $service + '\"'; ^
     $csv | Out-File -Encoding UTF8 '%OUTPUTCSV%'; ^
     Write-Host ''; ^
     Write-Host '======================================'; ^
     Write-Host '  SUCCESS: Exported to dsrf_export.csv'; ^
     Write-Host '======================================'; ^
     Write-Host ''; ^
     Write-Host ('Name:        ' + $name); ^
     Write-Host ('Company:     ' + $company); ^
     Write-Host ('Address 1:   ' + $addr1); ^
     Write-Host ('Address 2:   ' + $addr2); ^
     Write-Host ('City:        ' + $city); ^
     Write-Host ('State:       ' + $state); ^
     Write-Host ('Postal:      ' + $postal); ^
     Write-Host ('Service:     ' + $service); ^
   } catch { ^
     Write-Host ('ERROR: ' + $_.Exception.Message); ^
     exit 1; ^
   }"

if %ERRORLEVEL% neq 0 (
    echo.
    echo Check that your cookies are valid and not expired.
    pause
    exit /b 1
)

echo.
pause
