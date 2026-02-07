# DSRF-Export CLAUDE.md

Project-specific context for the DSRF-to-UPS_WS automation tools.

## Overview

Tools for fetching shipping data from Intra Shipping Request Forms (DSRF) and pasting to UPS WorldShip.

## API Details

- **Endpoint**: `https://amazonmailservices.us.spsprod.net/IntraWeb/api/automation/then/executeQuery`
- **Method**: POST with JSON body `{"Sql": "..."}`
- **Auth**: Cookie header from browser session (stored in `cookies.txt`)

## ItemVar ID Mappings

### Currently Used
| ItemVar ID | Field | WorldShip Location |
|------------|-------|-------------------|
| 148 | Company | Ship To tab |
| 149 | Name (Attention) | Ship To tab |
| 150 | Address Line 1 | Ship To tab |
| 151 | Address Line 2 | Ship To tab |
| 152 | City | Ship To tab (auto-fill) |
| 153 | State | Ship To tab (auto-fill) |
| 154 | Postal Code | Ship To tab |
| 155 | Service Type | Not currently pasted |
| 162 | Declared Value | Service tab |
| 202 | Ship From Name | Service tab (Ref2) |
| 203 | Email | Ship To tab |

### Other Discovered (metadata/flags)
| ItemVar ID | Value Found | Likely Field |
|------------|-------------|--------------|
| 143 | 1 | Unknown |
| 145 | Personal | Shipment Type |
| 147 | Residential | Address Type |
| 156-160 | boolean flags | Various |
| 190 | Any Other Domestic Location | Origin |
| 200 | Yes | Unknown |
| 211 | Generated | Status |
| 216/217 | AD | Unknown |
| 220 | DSRF for testing | Notes |

### Not Found (searched ranges 50-320)
- **Phone** (value: 2063567351) - not found in itemVars 50-320
- Ship From Company
- Ship From Address
- Cost Center

## SQL Query Pattern

Must use LEFT JOIN from Asset table (direct assetitemvars queries don't work reliably):

```sql
declare @profileid int = 1;
declare @assetid nvarchar(50) = 'PK######';
SELECT TOP 1
  ISNULL(iv148.ItemVarValue,'') as company,
  ...
FROM Asset a
LEFT JOIN assetitemvars iv148 ON iv148.assetid=a.assetid AND iv148.profileid=1 AND iv148.itemvarid=148
...
WHERE a.AssetID=@assetid AND a.ProfileID=@profileid
```

## WorldShip Coordinates (3440x1440)

See `ahk\DSRF-Export.ahk` for current coordinate definitions.

## Files

### AHK Version (in `ahk\` subfolder)
- `ahk\DSRF-Export.ahk` - Main AHK script (Ctrl+Alt+D hotkey, direct paste to WorldShip)
- `ahk\Cookie_Extractor.ahk` - Standalone cookie extraction (Ctrl+Alt+K)

### Standalone Version (no AHK required)
- `DSRF-Export.bat` - Launcher (prompts for PK#, runs .ps1)
- `DSRF-Export.ps1` - Core logic: cookie extraction, API call, XML/CSV generation
- `DSRF-ExportSetup.bat` - One-time setup (downloads sqlite3, creates folders)

### Shared
- `cookies.txt` - Session cookies (auto-generated, cached)
- `sqlite3.exe` - SQLite CLI tool
- `discover_itemvars.ps1` - Tool for finding itemVar IDs
- `discover_itemvars.bat` - Launcher for discovery tool
