# DSRF-Export

Tools for exporting DSRF (Desktop Shipping Request Form) data from Intra and pasting into UPS WorldShip.

## Required Programs

| Program | Purpose | Location |
|---------|---------|----------|
| **sqlite3.exe** | Read Firefox cookies database | `DSRF-Export\sqlite3.exe` |
| **AutoHotkey v1** | Run automation scripts | System-wide install |
| **Firefox** | Browser with Intra session cookies | N/A |

### sqlite3.exe

Required for `Cookie_Extractor.ahk` to read session cookies directly from Firefox's `cookies.sqlite` database.

**Download:** https://www.sqlite.org/download.html (Precompiled Binaries for Windows - sqlite-tools-win-x64)

Extract `sqlite3.exe` from the zip and place it in this folder.

## Scripts

### Cookie_Extractor.ahk

Extracts Intra session cookies from Firefox and saves to `cookies.txt`.

**Hotkey:** `Ctrl+Alt+K`

**How it works:**
1. Locates Firefox profile folder
2. Copies `cookies.sqlite` to temp (Firefox locks it while running)
3. Queries for cookies matching `amazonmailservices.us.spsprod.net`
4. Formats and saves to `cookies.txt`

**Note:** If Firefox has the database locked, close Firefox and retry.

### DSRF-to-UPS_WS-Paste.ahk

Fetches DSRF data from Intra API using cookies and pastes into UPS WorldShip.

**Hotkey:** `Ctrl+Alt+D`

**Requires:** Valid `cookies.txt` (run Cookie_Extractor first)

### ItemVar Discovery Tools

For discovering additional itemvars from the Intra API:

- `discover_itemvars.bat` - Interactive batch file
- `itemvar_discovery.ps1` - PowerShell script
- `ItemVar_Reference.txt` - Documentation of discovered itemvars

## Workflow

1. Log into Intra in Firefox
2. Press `Ctrl+Alt+K` to extract cookies
3. Open a DSRF form in Intra, copy the PK#
4. Open UPS WorldShip with a new shipment
5. Press `Ctrl+Alt+D` to fetch and paste data

## Files

| File | Description |
|------|-------------|
| `cookies.txt` | Session cookies (auto-generated) |
| `cookies.txt.bak` | Backup of previous cookies |
| `sqlite3.exe` | SQLite CLI tool |
| `ItemVar_Reference.txt` | ItemVar field mappings |
