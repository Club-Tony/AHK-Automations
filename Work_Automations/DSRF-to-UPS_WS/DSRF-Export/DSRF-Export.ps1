# DSRF-Export.ps1
# Fetches DSRF shipping data from Intra API and generates WorldShip import files (XML + CSV)
# Called by DSRF-Export.bat — do not run directly (use the .bat for pre-flight checks)

param(
    [string]$AssetId,
    [switch]$DetectOnly
)

$ErrorActionPreference = 'Stop'

# =============================================================================
# Configuration
# =============================================================================
$scriptDir     = $PSScriptRoot
$sqlite3       = Join-Path $scriptDir 'sqlite3.exe'
$cookiesFile   = Join-Path $scriptDir 'cookies.txt'
$importDir     = Join-Path $scriptDir 'import'
$intraDomain   = 'amazonmailservices.us.spsprod.net'
$apiUrl        = "https://$intraDomain/IntraWeb/api/automation/then/executeQuery"

# =============================================================================
# Detect-only mode: query Firefox history for most recent DSRF PK# and exit
# =============================================================================
if ($DetectOnly) {
    $ffProfiles = Join-Path $env:APPDATA 'Mozilla\Firefox\Profiles'
    if (-not (Test-Path $ffProfiles)) { exit }

    # Find default Firefox profile
    $profileDir = $null
    $profilesIni = Join-Path $env:APPDATA 'Mozilla\Firefox\profiles.ini'
    if (Test-Path $profilesIni) {
        $content = Get-Content $profilesIni -Raw
        $sections = $content -split '(?=\[Profile\d+\])'
        foreach ($section in $sections) {
            if ($section -match 'Default=1' -and $section -match 'Path=(.+)') {
                $path = $Matches[1].Trim()
                if ($section -match 'IsRelative=1') {
                    $profileDir = Join-Path (Join-Path $env:APPDATA 'Mozilla\Firefox') $path
                } else {
                    $profileDir = $path
                }
                break
            }
        }
    }
    if (-not $profileDir -or -not (Test-Path $profileDir)) {
        $profileDir = Get-ChildItem $ffProfiles -Directory |
            Where-Object { $_.Name -match '\.default-release$' } |
            Select-Object -First 1 -ExpandProperty FullName
    }
    if (-not $profileDir) { exit }

    $placesDb = Join-Path $profileDir 'places.sqlite'
    if (-not (Test-Path $placesDb)) { exit }

    # Copy to temp (Firefox locks it while running)
    $tempDb = Join-Path $env:TEMP 'dsrf_places_temp.sqlite'
    try {
        Copy-Item $placesDb $tempDb -Force -ErrorAction Stop
    } catch { exit }

    # Query for most recent DSRF URL containing assetId=PK
    $sqlQuery = "SELECT url FROM moz_places WHERE INSTR(url,'assetId=PK') > 0 ORDER BY last_visit_date DESC LIMIT 1;"
    $sqlFile = Join-Path $env:TEMP 'dsrf_detect_query.sql'
    $outputFile = Join-Path $env:TEMP 'dsrf_detect_output.txt'

    Remove-Item $sqlFile -Force -ErrorAction SilentlyContinue
    Remove-Item $outputFile -Force -ErrorAction SilentlyContinue

    $sqlQuery | Out-File -Encoding ASCII -FilePath $sqlFile -NoNewline

    $cmdArgs = "/c `"`"$sqlite3`" `"$tempDb`" < `"$sqlFile`" > `"$outputFile`"`""
    Start-Process -FilePath 'cmd.exe' -ArgumentList $cmdArgs -NoNewWindow -Wait

    Remove-Item $tempDb -Force -ErrorAction SilentlyContinue
    Remove-Item $sqlFile -Force -ErrorAction SilentlyContinue

    if (Test-Path $outputFile) {
        $url = (Get-Content $outputFile -Raw).Trim()
        Remove-Item $outputFile -Force -ErrorAction SilentlyContinue
        if ($url -match 'assetId=(PK\d+)') {
            Write-Output $Matches[1]
        }
    }
    exit
}

# Validate AssetId for normal mode
if (-not $AssetId) {
    Write-Host 'ERROR: No AssetId provided.'
    exit 1
}

# =============================================================================
# Functions
# =============================================================================

function Extract-FirefoxCookies {
    $ffProfiles = Join-Path $env:APPDATA 'Mozilla\Firefox\Profiles'
    if (-not (Test-Path $ffProfiles)) {
        Write-Host 'ERROR: Firefox profiles not found.'
        Write-Host 'Make sure Firefox is installed.'
        return $null
    }

    # Find default profile from profiles.ini
    $profileDir = $null
    $profilesIni = Join-Path $env:APPDATA 'Mozilla\Firefox\profiles.ini'

    if (Test-Path $profilesIni) {
        $content = Get-Content $profilesIni -Raw
        $sections = $content -split '(?=\[Profile\d+\])'

        foreach ($section in $sections) {
            if ($section -match 'Default=1' -and $section -match 'Path=(.+)') {
                $path = $Matches[1].Trim()
                if ($section -match 'IsRelative=1') {
                    $profileDir = Join-Path (Join-Path $env:APPDATA 'Mozilla\Firefox') $path
                } else {
                    $profileDir = $path
                }
                break
            }
        }
    }

    # Fallback: scan for .default-release profile
    if (-not $profileDir -or -not (Test-Path $profileDir)) {
        $profileDir = Get-ChildItem $ffProfiles -Directory |
            Where-Object { $_.Name -match '\.default-release$' } |
            Select-Object -First 1 -ExpandProperty FullName
    }
    if (-not $profileDir) {
        $profileDir = Get-ChildItem $ffProfiles -Directory |
            Where-Object { Test-Path (Join-Path $_.FullName 'cookies.sqlite') } |
            Select-Object -First 1 -ExpandProperty FullName
    }
    if (-not $profileDir) {
        Write-Host 'ERROR: No Firefox profile with cookies found.'
        return $null
    }

    $cookiesDb = Join-Path $profileDir 'cookies.sqlite'
    if (-not (Test-Path $cookiesDb)) {
        Write-Host 'ERROR: cookies.sqlite not found in profile.'
        return $null
    }

    # Copy database (Firefox locks it while running)
    $tempDb = Join-Path $env:TEMP 'firefox_cookies_temp.sqlite'
    try {
        Copy-Item $cookiesDb $tempDb -Force -ErrorAction Stop
    } catch {
        Write-Host 'ERROR: Cannot read Firefox cookies database.'
        Write-Host 'Firefox may have it locked. Close Firefox and retry.'
        return $null
    }

    # Query cookies using sqlite3
    $sqlQuery = "SELECT name || '=' || value FROM moz_cookies WHERE host LIKE '%$intraDomain%' ORDER BY name;"
    $sqlFile = Join-Path $env:TEMP 'cookie_query.sql'
    $outputFile = Join-Path $env:TEMP 'cookie_output.txt'

    # Clean up any previous temp files
    Remove-Item $sqlFile -Force -ErrorAction SilentlyContinue
    Remove-Item $outputFile -Force -ErrorAction SilentlyContinue

    $sqlQuery | Out-File -Encoding ASCII -FilePath $sqlFile -NoNewline

    # Run sqlite3 with cmd.exe redirection (same pattern as AHK version)
    $cmdArgs = "/c `"`"$sqlite3`" `"$tempDb`" < `"$sqlFile`" > `"$outputFile`"`""
    Start-Process -FilePath 'cmd.exe' -ArgumentList $cmdArgs -NoNewWindow -Wait

    # Clean up temp db and sql file
    Remove-Item $tempDb -Force -ErrorAction SilentlyContinue
    Remove-Item $sqlFile -Force -ErrorAction SilentlyContinue

    if (-not (Test-Path $outputFile)) {
        Write-Host 'ERROR: sqlite3 query produced no output.'
        return $null
    }

    $lines = Get-Content $outputFile | Where-Object { $_ -ne '' }
    Remove-Item $outputFile -Force -ErrorAction SilentlyContinue

    if (-not $lines -or $lines.Count -eq 0) {
        Write-Host 'ERROR: No cookies found for Intra domain.'
        Write-Host 'Make sure you are logged into Intra in Firefox.'
        return $null
    }

    $cookieString = $lines -join '; '
    if ($cookieString.Length -lt 100) {
        Write-Host 'ERROR: Extracted cookies seem incomplete.'
        Write-Host 'Log into Intra in Firefox and try again.'
        return $null
    }

    return $cookieString
}

function Invoke-IntraAPI {
    param(
        [string]$Cookies,
        [string]$PkNumber
    )

    $sql = @"
declare @profileid int = 1; declare @assetid nvarchar(50) = '$PkNumber';
SELECT TOP 1
ISNULL(iv148.ItemVarValue,'') as company,
ISNULL(iv149.ItemVarValue,'') as name,
ISNULL(iv150.ItemVarValue,'') as address1,
ISNULL(iv151.ItemVarValue,'') as address2,
ISNULL(iv152.ItemVarValue,'') as city,
ISNULL(iv153.ItemVarValue,'') as state,
ISNULL(iv154.ItemVarValue,'') as postal,
ISNULL(iv155.ItemVarValue,'') as serviceType,
ISNULL(iv162.ItemVarValue,'') as declaredValue,
ISNULL(iv202.ItemVarValue,'') as sfName,
ISNULL(iv203.ItemVarValue,'') as email,
ISNULL(iv38.ItemVarValue,'') as sfPhone,
ISNULL(iv41.ItemVarValue,'') as stPhone,
ISNULL(iv144.ItemVarValue,'') as costCenter,
ISNULL(iv145.ItemVarValue,'') as formType
FROM Asset a
LEFT JOIN assetitemvars iv148 ON iv148.assetid=a.assetid AND iv148.profileid=1 AND iv148.itemvarid=148
LEFT JOIN assetitemvars iv149 ON iv149.assetid=a.assetid AND iv149.profileid=1 AND iv149.itemvarid=149
LEFT JOIN assetitemvars iv150 ON iv150.assetid=a.assetid AND iv150.profileid=1 AND iv150.itemvarid=150
LEFT JOIN assetitemvars iv151 ON iv151.assetid=a.assetid AND iv151.profileid=1 AND iv151.itemvarid=151
LEFT JOIN assetitemvars iv152 ON iv152.assetid=a.assetid AND iv152.profileid=1 AND iv152.itemvarid=152
LEFT JOIN assetitemvars iv153 ON iv153.assetid=a.assetid AND iv153.profileid=1 AND iv153.itemvarid=153
LEFT JOIN assetitemvars iv154 ON iv154.assetid=a.assetid AND iv154.profileid=1 AND iv154.itemvarid=154
LEFT JOIN assetitemvars iv155 ON iv155.assetid=a.assetid AND iv155.profileid=1 AND iv155.itemvarid=155
LEFT JOIN assetitemvars iv162 ON iv162.assetid=a.assetid AND iv162.profileid=1 AND iv162.itemvarid=162
LEFT JOIN assetitemvars iv202 ON iv202.assetid=a.assetid AND iv202.profileid=1 AND iv202.itemvarid=202
LEFT JOIN assetitemvars iv203 ON iv203.assetid=a.assetid AND iv203.profileid=1 AND iv203.itemvarid=203
LEFT JOIN assetitemvars iv38 ON iv38.assetid=a.assetid AND iv38.profileid=1 AND iv38.itemvarid=38
LEFT JOIN assetitemvars iv41 ON iv41.assetid=a.assetid AND iv41.profileid=1 AND iv41.itemvarid=41
LEFT JOIN assetitemvars iv144 ON iv144.assetid=a.assetid AND iv144.profileid=1 AND iv144.itemvarid=144
LEFT JOIN assetitemvars iv145 ON iv145.assetid=a.assetid AND iv145.profileid=1 AND iv145.itemvarid=145
WHERE a.AssetID=@assetid AND a.ProfileID=@profileid
"@

    $body = @{ Sql = $sql } | ConvertTo-Json
    $headers = @{
        'Content-Type' = 'application/json'
        'Cookie'       = $Cookies
    }

    try {
        $response = Invoke-WebRequest -Uri $apiUrl -Method POST -Headers $headers -Body $body -UseBasicParsing
    } catch {
        return @{ Error = "API request failed: $($_.Exception.Message)" }
    }

    $content = $response.Content

    # Check for auth failures (login page, SAML redirect, access denied)
    if ($content -match '<!DOCTYPE' -or $content -match '<html' -or
        $content -match '<form.*login' -or $content -match 'SAMLRequest' -or
        $content -match 'Access Denied') {
        return @{ Error = 'AUTH_FAILED' }
    }

    if ([string]::IsNullOrWhiteSpace($content)) {
        return @{ Error = 'Empty response from API' }
    }

    $json = $content | ConvertFrom-Json
    if ($null -eq $json -or $json.Count -eq 0) {
        return @{ Error = 'No data returned for this PK#' }
    }

    return $json[0]
}

function Get-ServiceCode {
    param([string]$ServiceType)

    if ($ServiceType -match 'Ground')              { return 'GND' }
    if ($ServiceType -match 'Next Day Air Saver')   { return '1DP' }
    if ($ServiceType -match 'Next Day Air')         { return '1DA' }
    if ($ServiceType -match '2nd Day|Second Day')   { return '2DA' }
    if ($ServiceType -match '3.Day|3-Day')          { return '3DS' }
    return ''
}

function Get-Ref1Value {
    param($Data)

    if ($Data.costCenter) { return $Data.costCenter }
    if ($Data.formType -eq 'Personal') { return '9999' }
    return ''
}

function Write-WorldShipXML {
    param($Data, [string]$OutputPath)

    $serviceCode  = Get-ServiceCode $Data.serviceType
    $ref1         = Get-Ref1Value $Data
    $companyOrName = if ($Data.company) { $Data.company } else { $Data.name }

    # XML-escape helper
    function Esc([string]$s) {
        if (-not $s) { return '' }
        return [System.Security.SecurityElement]::Escape($s)
    }

    $xml = @"
<?xml version="1.0" encoding="utf-8"?>
<OpenShipments xmlns="x-schema:OpenShipments.xdr">
  <OpenShipment ShipmentOption="SC" ProcessStatus="">
    <ShipTo>
      <CompanyOrName>$(Esc $companyOrName)</CompanyOrName>
      <Attention>$(Esc $Data.name)</Attention>
      <Address1>$(Esc $Data.address1)</Address1>
      <Address2>$(Esc $Data.address2)</Address2>
      <CountryTerritory>US</CountryTerritory>
      <PostalCode>$(Esc $Data.postal)</PostalCode>
      <CityOrTown>$(Esc $Data.city)</CityOrTown>
      <StateProvinceCounty>$(Esc $Data.state)</StateProvinceCounty>
      <Telephone>$(Esc $Data.stPhone)</Telephone>
      <EmailAddress>$(Esc $Data.email)</EmailAddress>
    </ShipTo>
    <ShipFrom>
      <CompanyOrName>$(Esc $Data.sfName)</CompanyOrName>
      <Attention>$(Esc $Data.sfName)</Attention>
      <Telephone>$(Esc $Data.sfPhone)</Telephone>
    </ShipFrom>
    <ShipmentInformation>
      <ServiceType>$serviceCode</ServiceType>
      <NumberOfPackages>1</NumberOfPackages>
      <PackageType>CP</PackageType>
      <DeclaredValue>
        <Amount>$(if ($Data.declaredValue) { $Data.declaredValue } else { '0' })</Amount>
        <CurrencyCode>USD</CurrencyCode>
      </DeclaredValue>
      <Reference>
        <ReferenceNumber1>$(Esc $ref1)</ReferenceNumber1>
        <ReferenceNumber2>$(Esc $Data.sfName)</ReferenceNumber2>
      </Reference>
    </ShipmentInformation>
  </OpenShipment>
</OpenShipments>
"@

    $xml | Out-File -Encoding UTF8 -FilePath $OutputPath
}

function Write-ExportCSV {
    param($Data, [string]$OutputPath)

    $ref1          = Get-Ref1Value $Data
    $companyOrName = if ($Data.company) { $Data.company } else { $Data.name }

    $header = 'CompanyOrName,Attention,Address1,Address2,CityOrTown,StateProvinceCounty,PostalCode,CountryTerritory,Telephone,EmailAddress,ShipFromCompanyOrName,ShipFromAttention,ShipFromTelephone,ServiceType,DeclaredValue,ReferenceNumber1,ReferenceNumber2'

    $values = @(
        $companyOrName, $Data.name, $Data.address1, $Data.address2,
        $Data.city, $Data.state, $Data.postal, 'US',
        $Data.stPhone, $Data.email,
        $Data.sfName, $Data.sfName, $Data.sfPhone,
        $Data.serviceType, $Data.declaredValue, $ref1, $Data.sfName
    ) | ForEach-Object {
        '"' + ($_ -replace '"', '""') + '"'
    }

    $csv = $header + [Environment]::NewLine + ($values -join ',')
    $csv | Out-File -Encoding UTF8 -FilePath $OutputPath
}

function Show-Results {
    param($Data)

    $ref1 = Get-Ref1Value $Data

    Write-Host ''
    Write-Host '======================================'
    Write-Host '  Shipment Data'
    Write-Host '======================================'
    Write-Host ''
    Write-Host "  Company:       $($Data.company)"
    Write-Host "  Name:          $($Data.name)"
    Write-Host "  Address 1:     $($Data.address1)"
    Write-Host "  Address 2:     $($Data.address2)"
    Write-Host "  City:          $($Data.city)"
    Write-Host "  State:         $($Data.state)"
    Write-Host "  Postal:        $($Data.postal)"
    Write-Host "  ST Phone:      $($Data.stPhone)"
    Write-Host "  Email:         $($Data.email)"
    Write-Host ''
    Write-Host "  Ship From:     $($Data.sfName)"
    Write-Host "  SF Phone:      $($Data.sfPhone)"
    Write-Host ''
    Write-Host "  Service Type:  $($Data.serviceType)"
    Write-Host "  Decl. Value:   $($Data.declaredValue)"
    Write-Host "  Form Type:     $($Data.formType)"
    Write-Host "  Cost Center:   $($Data.costCenter)"
    Write-Host "  -> Ref 1:      $ref1"
    Write-Host "  -> Ref 2:      $($Data.sfName)"
    Write-Host ''
}

# =============================================================================
# Main Execution
# =============================================================================

# Step 1: Get cookies (automatic extraction via sqlite3)
$cookies = $null
$usedCache = $false

# Try cached cookies first
if (Test-Path $cookiesFile) {
    $cachedCookies = (Get-Content $cookiesFile -Raw).Trim()
    if ($cachedCookies.Length -gt 100) {
        Write-Host 'Using cached cookies...'
        $cookies = $cachedCookies
        $usedCache = $true
    }
}

# Extract fresh cookies if no valid cache
if (-not $cookies) {
    Write-Host 'Extracting cookies from Firefox...'
    $cookies = Extract-FirefoxCookies
    if (-not $cookies) { exit 1 }
    $cookies | Out-File -Encoding ASCII -FilePath $cookiesFile -NoNewline
    Write-Host 'Cookies extracted and saved.'
}

# Step 2: Call Intra API
Write-Host ''
Write-Host "Fetching data for $AssetId..."

$data = Invoke-IntraAPI -Cookies $cookies -PkNumber $AssetId

# If auth failed with cached cookies, try fresh extraction
if ($data.Error -eq 'AUTH_FAILED' -and $usedCache) {
    Write-Host 'Cached cookies expired. Re-extracting from Firefox...'
    $cookies = Extract-FirefoxCookies
    if (-not $cookies) { exit 1 }
    $cookies | Out-File -Encoding ASCII -FilePath $cookiesFile -NoNewline
    Write-Host 'Fresh cookies saved.'
    $data = Invoke-IntraAPI -Cookies $cookies -PkNumber $AssetId
}

# Check for errors
if ($data.Error) {
    Write-Host ''
    if ($data.Error -eq 'AUTH_FAILED') {
        Write-Host 'ERROR: Authentication failed - cookies are invalid.'
        Write-Host ''
        Write-Host 'To fix: Log into Intra in Firefox, then run this script again.'
        Write-Host 'The script will automatically extract fresh cookies.'
    } else {
        Write-Host "ERROR: $($data.Error)"
    }
    exit 1
}

# Validate we got actual shipping data
if ([string]::IsNullOrWhiteSpace($data.name) -and
    [string]::IsNullOrWhiteSpace($data.company) -and
    [string]::IsNullOrWhiteSpace($data.address1)) {
    Write-Host ''
    Write-Host 'ERROR: No shipping data found for this PK#.'
    Write-Host 'The form may not have shipping information filled in.'
    exit 1
}

# Step 3: Display results
Show-Results $data

# Step 4: Generate output files
$xmlFile = Join-Path $importDir "$AssetId.xml"
$csvFile = Join-Path $importDir "$AssetId.csv"

Write-WorldShipXML -Data $data -OutputPath $xmlFile
Write-ExportCSV -Data $data -OutputPath $csvFile

Write-Host '======================================'
Write-Host '  Files Generated'
Write-Host '======================================'
Write-Host ''
Write-Host "  XML: $xmlFile"
Write-Host "  CSV: $csvFile"
Write-Host ''
Write-Host '  If WorldShip Auto-Import is configured,'
Write-Host '  the XML file will be picked up automatically.'
Write-Host ''
Write-Host '======================================'
