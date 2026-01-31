# DSRF-to-UPS_WS-Paste.ahk

AutoHotkey script that fetches shipping data from Intra Shipping Request Form via API and pastes directly into UPS WorldShip.

## Requirements

- AutoHotkey v1
- Windows 10+ (uses PowerShell for API calls)
- Active Intra session in browser (Firefox, Chrome, or Edge)
- UPS WorldShip running
- Valid session cookies in `cookies.txt`

## Setup

### 1. Create cookies.txt

1. Open your browser and log into Intra
2. Navigate to any Shipping Request Form page
3. Press F12 to open Developer Tools
4. Go to the **Network** tab
5. Refresh the page (F5)
6. Click on any request in the list
7. Find **Request Headers** > **Cookie:**
8. Copy the entire cookie string
9. Paste into `cookies.txt` in this folder (single line, no line breaks)

**Example cookie format:**
```
CurrentProfile=ProfileId=1; AWSALB=xxx...; .AspNet.ApplicationCookie=xxx...
```

### 2. Run the Script

Double-click `DSRF-to-UPS_WS-Paste.ahk` to start.

## Usage

| Hotkey | Action |
|--------|--------|
| `Ctrl+Alt+D` | Fetch DSRF data and paste to WorldShip |

### Workflow

1. Open a Shipping Request Form in your browser (URL contains `assetId=PK######`)
2. Have UPS WorldShip open
3. Press `Ctrl+Alt+D`
4. Script will:
   - Detect the active DSRF window (Firefox, Chrome, or Edge)
   - Extract PK# from browser URL
   - Call Intra API to fetch shipping data
   - Switch to WorldShip and paste all fields

### Multiple DSRF Windows

If multiple Shipping Request Form windows are open with different PK#s, the script will prompt you to select which one to use.

## Supported Browsers

- Mozilla Firefox
- Google Chrome
- Microsoft Edge

## Field Mapping

| API Field | WorldShip Field |
|-----------|-----------------|
| itemVar149 | Recipient Name (Attention) |
| itemVar148 | Company |
| itemVar150 | Address Line 1 |
| itemVar151 | Address Line 2 |
| itemVar154 | Postal Code |
| (auto) | City/State (auto-filled by postal) |

## Troubleshooting

### "No Shipping Request Form windows found"
- Make sure a Shipping Request Form page is open in Firefox, Chrome, or Edge
- The window title must contain "Intra: Shipping Request Form"

### "Could not extract PK# from URL"
- URL must contain `assetId=PK######`
- Try refreshing the DSRF page

### "Could not get session cookies"
- Create `cookies.txt` in the script folder with your session cookies
- Or set `INTRA_COOKIES` environment variable

### "Cookie Error: API call failed - cookies are likely invalid or expired"
This is the most common error. The script validates the API response and detects:
- HTML login pages (session expired)
- Error responses from the server
- Empty or malformed data

**To fix:**
1. Open Intra in your browser and log in
2. Press F12 to open DevTools
3. Go to Network tab, refresh the page
4. Click any request, find the **Cookie:** header
5. Copy the entire cookie string
6. Paste into `cookies.txt` (replace all existing content)

### "Could not activate UPS WorldShip window"
- Make sure UPS WorldShip is running

## Cookie Validation

The script now includes robust cookie validation:
- Checks HTTP response status
- Detects HTML login/error pages in response
- Validates JSON format before parsing
- Verifies shipping data fields are present
- Validates postal code format

If cookies are invalid, the script will show an error message **before** attempting to paste anything to WorldShip.

## Notes

- Cookies expire periodically - refresh them when you see cookie errors
- Script only triggers when a valid DSRF window is detected
- The script validates all data before pasting to prevent garbage data in WorldShip
