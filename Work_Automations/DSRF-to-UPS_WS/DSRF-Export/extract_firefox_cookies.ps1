# extract_firefox_cookies.ps1
# Extracts cookies for Intra domain directly from Firefox's cookies.sqlite database
# No browser extension required!

param(
    [string]$Domain = "amazonmailservices.us.spsprod.net",
    [string]$OutputFile = "$PSScriptRoot\cookies.txt"
)

# Find Firefox profile directory
$firefoxProfiles = "$env:APPDATA\Mozilla\Firefox\Profiles"
if (-not (Test-Path $firefoxProfiles)) {
    Write-Host "ERROR: Firefox profiles not found at $firefoxProfiles"
    exit 1
}

# Find the default profile (usually ends with .default-release or .default)
$profileDir = Get-ChildItem $firefoxProfiles -Directory | Where-Object {
    $_.Name -match '\.default-release$|\.default$'
} | Select-Object -First 1

if (-not $profileDir) {
    # Try any profile
    $profileDir = Get-ChildItem $firefoxProfiles -Directory | Select-Object -First 1
}

if (-not $profileDir) {
    Write-Host "ERROR: No Firefox profile found"
    exit 1
}

$cookiesDb = Join-Path $profileDir.FullName "cookies.sqlite"
if (-not (Test-Path $cookiesDb)) {
    Write-Host "ERROR: cookies.sqlite not found at $cookiesDb"
    exit 1
}

Write-Host "Found Firefox profile: $($profileDir.Name)"
Write-Host "Reading cookies from: $cookiesDb"

# Firefox locks the database while running, so we need to copy it first
$tempDb = "$env:TEMP\firefox_cookies_temp.sqlite"
try {
    Copy-Item $cookiesDb $tempDb -Force -ErrorAction Stop
} catch {
    Write-Host "ERROR: Could not copy cookies database. Firefox may have it locked."
    Write-Host "Try closing Firefox or wait a moment and try again."
    exit 1
}

# Use System.Data.SQLite or fall back to sqlite3 CLI
# First, try using ADO.NET with SQLite
$connectionString = "Data Source=$tempDb;Version=3;Read Only=True;"

try {
    # Try loading SQLite assembly
    Add-Type -Path "$env:ProgramFiles\System.Data.SQLite\2010\bin\System.Data.SQLite.dll" -ErrorAction SilentlyContinue
} catch {
    # Assembly not found, we'll use alternative method
}

# Alternative: Use sqlite3 CLI if available, or parse with PowerShell
# For simplicity, let's use a .NET approach that works without external dependencies

# Build SQL query for the domain
$query = @"
SELECT name, value FROM moz_cookies
WHERE host LIKE '%$Domain%' OR baseDomain LIKE '%$Domain%'
ORDER BY name
"@

# Try using sqlite3 command line if available
$sqlite3 = @(
    "$env:ProgramFiles\SQLite\sqlite3.exe",
    "$env:ProgramFiles (x86)\SQLite\sqlite3.exe",
    "sqlite3"  # Check PATH
) | Where-Object {
    if ($_ -eq "sqlite3") {
        Get-Command sqlite3 -ErrorAction SilentlyContinue
    } else {
        Test-Path $_
    }
} | Select-Object -First 1

if ($sqlite3) {
    Write-Host "Using sqlite3 CLI..."
    $result = & $sqlite3 $tempDb "SELECT name || '=' || value FROM moz_cookies WHERE host LIKE '%$Domain%' ORDER BY name;" 2>&1
    if ($LASTEXITCODE -eq 0 -and $result) {
        $cookieString = ($result -join "; ")
    }
} else {
    # Fallback: Use a simple binary read approach for the most common cookies
    # This is less reliable but doesn't require sqlite3
    Write-Host "sqlite3 not found, using binary search method..."

    $dbContent = [System.IO.File]::ReadAllText($tempDb, [System.Text.Encoding]::UTF8)

    # Look for common cookie patterns
    $cookies = @()
    $patterns = @(
        'CurrentProfile=ProfileId=\d+',
        'AWSALBTG=[A-Za-z0-9+/=]+',
        'AWSALBTGCORS=[A-Za-z0-9+/=]+',
        'AWSALB=[A-Za-z0-9+/=]+',
        'AWSALBCORS=[A-Za-z0-9+/=]+',
        '\.AspNet\.ApplicationCookie=[A-Za-z0-9_-]+',
        'Saml2\.[A-Za-z0-9]+=[A-Za-z0-9_-]+',
        'ASP\.NET_SessionId=[A-Za-z0-9]+'
    )

    foreach ($pattern in $patterns) {
        if ($dbContent -match $pattern) {
            $cookies += $Matches[0]
        }
    }

    if ($cookies.Count -gt 0) {
        $cookieString = $cookies -join "; "
    }
}

# Clean up temp file
Remove-Item $tempDb -Force -ErrorAction SilentlyContinue

if (-not $cookieString -or $cookieString.Length -lt 50) {
    Write-Host "ERROR: Could not extract cookies for domain $Domain"
    Write-Host "Make sure you're logged into Intra in Firefox."
    exit 1
}

# Validate we got session cookies
$hasSession = $cookieString -match 'ASP\.NET_SessionId|AWSALB|Saml2'
if (-not $hasSession) {
    Write-Host "WARNING: Extracted cookies may be incomplete"
}

# Save to output file
$cookieString | Out-File -FilePath $OutputFile -Encoding ASCII -NoNewline

Write-Host ""
Write-Host "SUCCESS: Cookies extracted!"
Write-Host "Length: $($cookieString.Length) characters"
Write-Host "Saved to: $OutputFile"
Write-Host ""
Write-Host "Preview: $($cookieString.Substring(0, [Math]::Min(100, $cookieString.Length)))..."
