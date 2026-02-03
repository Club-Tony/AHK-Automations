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

# Build SQL - Join Asset to assetitemvars using the id column
# The working query joins on a.assetid but that's the string PK#
# Try joining on a.id instead (the numeric internal ID)
$sql = @"
declare @pk nvarchar(50) = '$AssetId';
SELECT iv.itemvarid, iv.ItemVarValue
FROM Asset a
INNER JOIN assetitemvars iv ON iv.assetid = a.id AND iv.profileid = 1
WHERE a.AssetID = @pk AND a.ProfileID = 1
  AND iv.ItemVarValue IS NOT NULL AND iv.ItemVarValue != ''
ORDER BY iv.itemvarid
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
            $out += "Found $($json.Count) itemVars with values:"
            $out += ''
            $out += 'ItemVarID | Value'
            $out += '----------|------'

            # Each row has itemvarid and ItemVarValue
            foreach ($row in $json) {
                $id = $row.itemvarid
                $val = $row.ItemVarValue
                # Truncate long values for display
                if ($val.Length -gt 50) {
                    $val = $val.Substring(0, 47) + '...'
                }
                $out += "$id | $val"
            }

            $out += ''
            $out += '(Showing ALL itemVars with non-empty values)'
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
$out += 'Known mappings:'
$out += '  148=Company, 149=Name, 150=Addr1, 151=Addr2'
$out += '  152=City, 153=State, 154=Postal, 155=Service'
$out += '  162=DeclaredValue, 202=ShipFromName, 203=Email'
$out += ''
$out += 'Looking for: Phone, Cost Center'
$out += '======================================'

# Save to file
$out | Out-File -Encoding UTF8 $OutputFile
Write-Host ""
Write-Host "Results saved to: $OutputFile"
Write-Host ""

# Open in notepad
Start-Process notepad $OutputFile
