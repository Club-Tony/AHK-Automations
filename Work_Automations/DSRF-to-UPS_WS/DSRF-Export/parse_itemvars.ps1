# parse_itemvars.ps1 - Parse itemVar discovery results
param(
    [string]$JsonFile,
    [string]$OutputFile,
    [string]$AssetId
)

$ErrorActionPreference = 'Continue'

try {
    $out = @()
    $out += '======================================'
    $out += '  ItemVar Discovery Results'
    $out += "  Asset: $AssetId"
    $out += '======================================'
    $out += ''

    if (-not (Test-Path $JsonFile)) {
        $out += 'ERROR: Response file not found'
        $out += "Expected: $JsonFile"
        $out | Out-File -Encoding UTF8 $OutputFile
        Write-Host 'Error: Response file not found'
        exit 1
    }

    $content = Get-Content $JsonFile -Raw

    if ([string]::IsNullOrWhiteSpace($content)) {
        $out += 'ERROR: Empty response from API'
        $out += 'Cookies may be invalid.'
        $out | Out-File -Encoding UTF8 $OutputFile
        Write-Host 'Error: Empty response'
        exit 1
    }

    # Check for HTML (login page)
    if ($content -match '<!DOCTYPE' -or $content -match '<html') {
        $out += 'ERROR: Session expired - cookies invalid'
        $out += 'Please refresh cookies from browser DevTools'
        $out += ''
        $out += 'Raw response (first 500 chars):'
        $out += $content.Substring(0, [Math]::Min(500, $content.Length))
        $out | Out-File -Encoding UTF8 $OutputFile
        Write-Host 'Error: Cookies invalid (got login page)'
        exit 1
    }

    # Try to parse JSON
    try {
        $json = $content | ConvertFrom-Json
    } catch {
        $out += 'ERROR: Failed to parse JSON'
        $out += $_.Exception.Message
        $out += ''
        $out += 'Raw response (first 500 chars):'
        $out += $content.Substring(0, [Math]::Min(500, $content.Length))
        $out | Out-File -Encoding UTF8 $OutputFile
        Write-Host 'Error: JSON parse failed'
        exit 1
    }

    if ($null -eq $json -or $json.Count -eq 0) {
        $out += 'No itemVars found for this PK#'
    } else {
        $out += "Found $($json.Count) itemVars:"
        $out += ''
        $out += 'ItemVarID | ItemVarName | Value'
        $out += '----------|-------------|------'
        foreach ($row in $json) {
            $id = $row.itemvarid
            $name = if ($row.ItemVarName) { $row.ItemVarName } else { '(unknown)' }
            $val = if ($row.ItemVarValue) { $row.ItemVarValue } else { '' }
            $out += "$id | $name | $val"
        }
    }

    $out += ''
    $out += '======================================'
    $out += 'Look for these values:'
    $out += '  Phone: 2063567351'
    $out += '  Email: daveyuan@amazon.com'
    $out += '  Declared Value: 100'
    $out += '  Ship From Name: Davey, Anthony'
    $out += '======================================'

    $out | Out-File -Encoding UTF8 $OutputFile
    Write-Host "Results saved to $OutputFile"

} catch {
    "ERROR: $($_.Exception.Message)" | Out-File -Encoding UTF8 $OutputFile
    Write-Host "Unexpected error: $($_.Exception.Message)"
    exit 1
}
