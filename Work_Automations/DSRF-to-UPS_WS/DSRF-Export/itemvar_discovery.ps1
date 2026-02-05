param(
    [string]$AssetId = "PK123456",
    [int]$StartRange = 1,
    [int]$EndRange = 200
)

$cookies = Get-Content "$PSScriptRoot\cookies.txt" -Raw

# Build SQL for itemvars in range
$selectParts = @()
for ($i = $StartRange; $i -le $EndRange; $i++) {
    $selectParts += "ISNULL(iv$i.ItemVarValue,'') as iv$i"
}
$joinParts = @()
for ($i = $StartRange; $i -le $EndRange; $i++) {
    $joinParts += "LEFT JOIN assetitemvars iv$i ON iv$i.assetid=a.assetid AND iv$i.profileid=1 AND iv$i.itemvarid=$i"
}

$sql = "declare @profileid int = 1; declare @assetid nvarchar(50) = '$AssetId'; SELECT TOP 1 " + ($selectParts -join ', ') + " FROM Asset a " + ($joinParts -join ' ') + " WHERE a.AssetID=@assetid AND a.ProfileID=@profileid"

$body = @{ Sql = $sql } | ConvertTo-Json
$headers = @{ 'Content-Type'='application/json'; 'Cookie'=$cookies }

try {
    $response = Invoke-WebRequest -Uri 'https://amazonmailservices.us.spsprod.net/IntraWeb/api/automation/then/executeQuery' -Method POST -Headers $headers -Body $body -UseBasicParsing
    Write-Host "Status: $($response.StatusCode)"
    $content = $response.Content
    if ($content -match '<!DOCTYPE' -or $content -match '<html' -or $content -match 'login') {
        Write-Host "ERROR: Got login page - cookies expired"
        exit 1
    }
    Write-Host "Response length: $($content.Length)"
    Write-Host "Raw response: $content"
    $json = $content | ConvertFrom-Json
    Write-Host "JSON type: $($json.GetType().Name)"
    # Handle both array and single object response
    if ($json -is [array]) {
        $row = $json[0]
    } else {
        $row = $json
    }
    if ($row) {
        Write-Host "======================================"
        Write-Host "  ItemVar Discovery Results ($StartRange-$EndRange)"
        Write-Host "  Asset: $AssetId"
        Write-Host "======================================"
        Write-Host ""
        Write-Host "ItemVarID | Value"
        Write-Host "----------|------"
        for ($i = $StartRange; $i -le $EndRange; $i++) {
            $val = $row."iv$i"
            if ($val -and $val -ne '') {
                Write-Host "$i | $val"
            }
        }
        Write-Host ""
        Write-Host "(Only showing itemVars with values)"
    } else {
        Write-Host "No data returned for asset $AssetId"
    }
} catch {
    Write-Host "API Error: $($_.Exception.Message)"
}
