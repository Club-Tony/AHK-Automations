# discover_itemvars.ps1 - Standalone PowerShell version
param(
    [Parameter(Mandatory=$true)]
    [string]$AssetId,

    [Parameter(Mandatory=$true)]
    [string]$CookieFile,

    [Parameter(Mandatory=$true)]
    [string]$OutputFile
)

$ErrorActionPreference = 'Stop'

Write-Host "======================================"
Write-Host "  ItemVar Discovery Tool (PS)"
Write-Host "======================================"
Write-Host ""

# Read cookies
if (-not (Test-Path $CookieFile)) {
    Write-Host "ERROR: Cookie file not found: $CookieFile"
    exit 1
}

$cookies = (Get-Content $CookieFile -Raw).Trim()
Write-Host "DEBUG: Cookie length = $($cookies.Length)"

# Build SQL using working LEFT JOIN pattern - search itemVars 1-50
$sql = @"
declare @profileid int = 1;
declare @assetid nvarchar(50) = '$AssetId';
SELECT TOP 1
  ISNULL(iv1.ItemVarValue,'') as iv1,
  ISNULL(iv2.ItemVarValue,'') as iv2,
  ISNULL(iv3.ItemVarValue,'') as iv3,
  ISNULL(iv4.ItemVarValue,'') as iv4,
  ISNULL(iv5.ItemVarValue,'') as iv5,
  ISNULL(iv6.ItemVarValue,'') as iv6,
  ISNULL(iv7.ItemVarValue,'') as iv7,
  ISNULL(iv8.ItemVarValue,'') as iv8,
  ISNULL(iv9.ItemVarValue,'') as iv9,
  ISNULL(iv10.ItemVarValue,'') as iv10,
  ISNULL(iv11.ItemVarValue,'') as iv11,
  ISNULL(iv12.ItemVarValue,'') as iv12,
  ISNULL(iv13.ItemVarValue,'') as iv13,
  ISNULL(iv14.ItemVarValue,'') as iv14,
  ISNULL(iv15.ItemVarValue,'') as iv15,
  ISNULL(iv16.ItemVarValue,'') as iv16,
  ISNULL(iv17.ItemVarValue,'') as iv17,
  ISNULL(iv18.ItemVarValue,'') as iv18,
  ISNULL(iv19.ItemVarValue,'') as iv19,
  ISNULL(iv20.ItemVarValue,'') as iv20,
  ISNULL(iv21.ItemVarValue,'') as iv21,
  ISNULL(iv22.ItemVarValue,'') as iv22,
  ISNULL(iv23.ItemVarValue,'') as iv23,
  ISNULL(iv24.ItemVarValue,'') as iv24,
  ISNULL(iv25.ItemVarValue,'') as iv25,
  ISNULL(iv26.ItemVarValue,'') as iv26,
  ISNULL(iv27.ItemVarValue,'') as iv27,
  ISNULL(iv28.ItemVarValue,'') as iv28,
  ISNULL(iv29.ItemVarValue,'') as iv29,
  ISNULL(iv30.ItemVarValue,'') as iv30,
  ISNULL(iv31.ItemVarValue,'') as iv31,
  ISNULL(iv32.ItemVarValue,'') as iv32,
  ISNULL(iv33.ItemVarValue,'') as iv33,
  ISNULL(iv34.ItemVarValue,'') as iv34,
  ISNULL(iv35.ItemVarValue,'') as iv35,
  ISNULL(iv36.ItemVarValue,'') as iv36,
  ISNULL(iv37.ItemVarValue,'') as iv37,
  ISNULL(iv38.ItemVarValue,'') as iv38,
  ISNULL(iv39.ItemVarValue,'') as iv39,
  ISNULL(iv40.ItemVarValue,'') as iv40,
  ISNULL(iv41.ItemVarValue,'') as iv41,
  ISNULL(iv42.ItemVarValue,'') as iv42,
  ISNULL(iv43.ItemVarValue,'') as iv43,
  ISNULL(iv44.ItemVarValue,'') as iv44,
  ISNULL(iv45.ItemVarValue,'') as iv45,
  ISNULL(iv46.ItemVarValue,'') as iv46,
  ISNULL(iv47.ItemVarValue,'') as iv47,
  ISNULL(iv48.ItemVarValue,'') as iv48,
  ISNULL(iv49.ItemVarValue,'') as iv49,
  ISNULL(iv50.ItemVarValue,'') as iv50
FROM Asset a
LEFT JOIN assetitemvars iv1 ON iv1.assetid=a.assetid AND iv1.profileid=1 AND iv1.itemvarid=1
LEFT JOIN assetitemvars iv2 ON iv2.assetid=a.assetid AND iv2.profileid=1 AND iv2.itemvarid=2
LEFT JOIN assetitemvars iv3 ON iv3.assetid=a.assetid AND iv3.profileid=1 AND iv3.itemvarid=3
LEFT JOIN assetitemvars iv4 ON iv4.assetid=a.assetid AND iv4.profileid=1 AND iv4.itemvarid=4
LEFT JOIN assetitemvars iv5 ON iv5.assetid=a.assetid AND iv5.profileid=1 AND iv5.itemvarid=5
LEFT JOIN assetitemvars iv6 ON iv6.assetid=a.assetid AND iv6.profileid=1 AND iv6.itemvarid=6
LEFT JOIN assetitemvars iv7 ON iv7.assetid=a.assetid AND iv7.profileid=1 AND iv7.itemvarid=7
LEFT JOIN assetitemvars iv8 ON iv8.assetid=a.assetid AND iv8.profileid=1 AND iv8.itemvarid=8
LEFT JOIN assetitemvars iv9 ON iv9.assetid=a.assetid AND iv9.profileid=1 AND iv9.itemvarid=9
LEFT JOIN assetitemvars iv10 ON iv10.assetid=a.assetid AND iv10.profileid=1 AND iv10.itemvarid=10
LEFT JOIN assetitemvars iv11 ON iv11.assetid=a.assetid AND iv11.profileid=1 AND iv11.itemvarid=11
LEFT JOIN assetitemvars iv12 ON iv12.assetid=a.assetid AND iv12.profileid=1 AND iv12.itemvarid=12
LEFT JOIN assetitemvars iv13 ON iv13.assetid=a.assetid AND iv13.profileid=1 AND iv13.itemvarid=13
LEFT JOIN assetitemvars iv14 ON iv14.assetid=a.assetid AND iv14.profileid=1 AND iv14.itemvarid=14
LEFT JOIN assetitemvars iv15 ON iv15.assetid=a.assetid AND iv15.profileid=1 AND iv15.itemvarid=15
LEFT JOIN assetitemvars iv16 ON iv16.assetid=a.assetid AND iv16.profileid=1 AND iv16.itemvarid=16
LEFT JOIN assetitemvars iv17 ON iv17.assetid=a.assetid AND iv17.profileid=1 AND iv17.itemvarid=17
LEFT JOIN assetitemvars iv18 ON iv18.assetid=a.assetid AND iv18.profileid=1 AND iv18.itemvarid=18
LEFT JOIN assetitemvars iv19 ON iv19.assetid=a.assetid AND iv19.profileid=1 AND iv19.itemvarid=19
LEFT JOIN assetitemvars iv20 ON iv20.assetid=a.assetid AND iv20.profileid=1 AND iv20.itemvarid=20
LEFT JOIN assetitemvars iv21 ON iv21.assetid=a.assetid AND iv21.profileid=1 AND iv21.itemvarid=21
LEFT JOIN assetitemvars iv22 ON iv22.assetid=a.assetid AND iv22.profileid=1 AND iv22.itemvarid=22
LEFT JOIN assetitemvars iv23 ON iv23.assetid=a.assetid AND iv23.profileid=1 AND iv23.itemvarid=23
LEFT JOIN assetitemvars iv24 ON iv24.assetid=a.assetid AND iv24.profileid=1 AND iv24.itemvarid=24
LEFT JOIN assetitemvars iv25 ON iv25.assetid=a.assetid AND iv25.profileid=1 AND iv25.itemvarid=25
LEFT JOIN assetitemvars iv26 ON iv26.assetid=a.assetid AND iv26.profileid=1 AND iv26.itemvarid=26
LEFT JOIN assetitemvars iv27 ON iv27.assetid=a.assetid AND iv27.profileid=1 AND iv27.itemvarid=27
LEFT JOIN assetitemvars iv28 ON iv28.assetid=a.assetid AND iv28.profileid=1 AND iv28.itemvarid=28
LEFT JOIN assetitemvars iv29 ON iv29.assetid=a.assetid AND iv29.profileid=1 AND iv29.itemvarid=29
LEFT JOIN assetitemvars iv30 ON iv30.assetid=a.assetid AND iv30.profileid=1 AND iv30.itemvarid=30
LEFT JOIN assetitemvars iv31 ON iv31.assetid=a.assetid AND iv31.profileid=1 AND iv31.itemvarid=31
LEFT JOIN assetitemvars iv32 ON iv32.assetid=a.assetid AND iv32.profileid=1 AND iv32.itemvarid=32
LEFT JOIN assetitemvars iv33 ON iv33.assetid=a.assetid AND iv33.profileid=1 AND iv33.itemvarid=33
LEFT JOIN assetitemvars iv34 ON iv34.assetid=a.assetid AND iv34.profileid=1 AND iv34.itemvarid=34
LEFT JOIN assetitemvars iv35 ON iv35.assetid=a.assetid AND iv35.profileid=1 AND iv35.itemvarid=35
LEFT JOIN assetitemvars iv36 ON iv36.assetid=a.assetid AND iv36.profileid=1 AND iv36.itemvarid=36
LEFT JOIN assetitemvars iv37 ON iv37.assetid=a.assetid AND iv37.profileid=1 AND iv37.itemvarid=37
LEFT JOIN assetitemvars iv38 ON iv38.assetid=a.assetid AND iv38.profileid=1 AND iv38.itemvarid=38
LEFT JOIN assetitemvars iv39 ON iv39.assetid=a.assetid AND iv39.profileid=1 AND iv39.itemvarid=39
LEFT JOIN assetitemvars iv40 ON iv40.assetid=a.assetid AND iv40.profileid=1 AND iv40.itemvarid=40
LEFT JOIN assetitemvars iv41 ON iv41.assetid=a.assetid AND iv41.profileid=1 AND iv41.itemvarid=41
LEFT JOIN assetitemvars iv42 ON iv42.assetid=a.assetid AND iv42.profileid=1 AND iv42.itemvarid=42
LEFT JOIN assetitemvars iv43 ON iv43.assetid=a.assetid AND iv43.profileid=1 AND iv43.itemvarid=43
LEFT JOIN assetitemvars iv44 ON iv44.assetid=a.assetid AND iv44.profileid=1 AND iv44.itemvarid=44
LEFT JOIN assetitemvars iv45 ON iv45.assetid=a.assetid AND iv45.profileid=1 AND iv45.itemvarid=45
LEFT JOIN assetitemvars iv46 ON iv46.assetid=a.assetid AND iv46.profileid=1 AND iv46.itemvarid=46
LEFT JOIN assetitemvars iv47 ON iv47.assetid=a.assetid AND iv47.profileid=1 AND iv47.itemvarid=47
LEFT JOIN assetitemvars iv48 ON iv48.assetid=a.assetid AND iv48.profileid=1 AND iv48.itemvarid=48
LEFT JOIN assetitemvars iv49 ON iv49.assetid=a.assetid AND iv49.profileid=1 AND iv49.itemvarid=49
LEFT JOIN assetitemvars iv50 ON iv50.assetid=a.assetid AND iv50.profileid=1 AND iv50.itemvarid=50
WHERE a.AssetID=@assetid AND a.ProfileID=@profileid
"@

# Create JSON body
$body = @{ Sql = $sql } | ConvertTo-Json -Compress
Write-Host "DEBUG: Request body length = $($body.Length)"

# Make API call
$uri = "https://amazonmailservices.us.spsprod.net/IntraWeb/api/automation/then/executeQuery"

Write-Host "DEBUG: Calling API..."
Write-Host ""

try {
    $headers = @{
        'Content-Type' = 'application/json'
        'Cookie' = $cookies
    }

    $response = Invoke-WebRequest -Uri $uri -Method POST -Headers $headers -Body $body -UseBasicParsing

    Write-Host "DEBUG: HTTP Status = $($response.StatusCode)"
    Write-Host "DEBUG: Response length = $($response.Content.Length)"

    $content = $response.Content

} catch {
    Write-Host "ERROR: API call failed"
    Write-Host "Exception: $($_.Exception.Message)"
    Write-Host ""

    if ($_.Exception.Response) {
        $statusCode = [int]$_.Exception.Response.StatusCode
        Write-Host "HTTP Status: $statusCode"

        # Try to read error response body
        try {
            $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
            $errorBody = $reader.ReadToEnd()
            Write-Host "Error Response (first 500 chars):"
            Write-Host $errorBody.Substring(0, [Math]::Min(500, $errorBody.Length))
        } catch {
            Write-Host "Could not read error response body"
        }
    }

    # Write error to output file so user can see it
    $errorOut = @()
    $errorOut += '======================================'
    $errorOut += '  ItemVar Discovery - ERROR'
    $errorOut += "  Asset: $AssetId"
    $errorOut += '======================================'
    $errorOut += ''
    $errorOut += 'API call failed - see console output above'
    $errorOut | Out-File -Encoding UTF8 $OutputFile
    Start-Process notepad $OutputFile
    exit 1
}

# Build output
$out = @()
$out += '======================================'
$out += '  ItemVar Discovery Results'
$out += "  Asset: $AssetId"
$out += '======================================'
$out += ''

# Check for HTML (login page)
if ($content -match '<!DOCTYPE' -or $content -match '<html') {
    $out += 'ERROR: Session expired - cookies invalid'
    $out += 'Please refresh cookies from browser DevTools'
    $out += ''
    $out += 'Raw response (first 500 chars):'
    $out += $content.Substring(0, [Math]::Min(500, $content.Length))
} else {
    # Parse JSON
    try {
        $json = $content | ConvertFrom-Json

        if ($null -eq $json -or $json.Count -eq 0) {
            $out += 'No data found for this PK#'
            $out += ''
            $out += 'The asset may not have any itemVars stored.'
        } else {
            $row = $json[0]
            $out += 'ItemVar values (1-50):'
            $out += ''
            $out += 'ItemVarID | Value'
            $out += '----------|------'

            # Loop through itemVars 1-50
            for ($i = 1; $i -le 50; $i++) {
                $colName = "iv$i"
                $val = $row.$colName
                if ($val -and $val -ne '') {
                    # Truncate long values
                    if ($val.Length -gt 50) {
                        $val = $val.Substring(0, 47) + '...'
                    }
                    $out += "$i | $val"
                }
            }

            $out += ''
            $out += '(Only showing itemVars with values)'
        }
    } catch {
        $out += "ERROR parsing JSON: $($_.Exception.Message)"
        $out += ''
        $out += 'Raw response (first 500 chars):'
        $out += $content.Substring(0, [Math]::Min(500, $content.Length))
    }
}

$out += ''
$out += '======================================'
$out += 'Searching range: 1-50'
$out += 'Looking for: Phone (2063567351), Cost Center (9999)'
$out += '======================================'

# Save to file
$out | Out-File -Encoding UTF8 $OutputFile
Write-Host ""
Write-Host "Results saved to: $OutputFile"
Write-Host ""

# Open in notepad
Start-Process notepad $OutputFile
