# DSRF-Export

Tools for exporting DSRF (Desktop Shipping Request Form) data from Intra and importing into UPS WorldShip.

## Two Versions

### Standalone Version (No AHK Required)
Uses `DSRF-Export.bat` + `DSRF-Export.ps1` to fetch data and generate WorldShip-compatible XML/CSV files.

### AHK Version (Direct Paste)
Uses `ahk\DSRF-Export.ahk` to fetch data and paste directly into WorldShip fields via automation.

---

## Standalone Version Setup

### Requirements

| Program | Purpose | Install |
|---------|---------|---------|
| **sqlite3.exe** | Read Firefox cookies database | Run `DSRF-ExportSetup.bat` (auto-downloads) |
| **Firefox** | Browser with Intra session | Already installed |

### First-Time Setup

1. Run `DSRF-ExportSetup.bat` — downloads sqlite3.exe and creates the `import\` folder
2. Configure WorldShip Auto-Import (one-time):
   - In WorldShip: Import/Export > Auto Import
   - Set import directory to the `import\` subfolder
   - Enable auto-processing

### Standalone Workflow

1. Log into Intra in Firefox
2. Double-click `DSRF-Export.bat`
3. Enter the PK# when prompted
4. Script auto-extracts cookies, fetches data, generates files in `import\`
5. WorldShip picks up the XML automatically (if Auto-Import configured)
   - Or use the CSV file for manual batch import

### Output Files

Generated in `import\` subfolder, named by PK#:
- `PK######.xml` — WorldShip XML Auto-Import format (no field mapping needed)
- `PK######.csv` — Standard CSV for batch import (requires one-time field mapping)

---

## AHK Version Setup

### Requirements

| Program | Purpose | Install |
|---------|---------|---------|
| **sqlite3.exe** | Read Firefox cookies database | `DSRF-Export\sqlite3.exe` |
| **AutoHotkey v1** | Run automation scripts | System-wide install |
| **Firefox** | Browser with Intra session | Already installed |

### AHK Workflow

1. Log into Intra in Firefox
2. Open a DSRF form in Intra
3. Open UPS WorldShip with a new shipment
4. Press `Ctrl+Alt+D` to fetch and paste data directly into WorldShip
5. Press `Esc` to cancel/reload if needed

---

## Scripts

| File | Description |
|------|-------------|
| `DSRF-Export.bat` | Standalone launcher (prompts for PK#, runs .ps1) |
| `DSRF-Export.ps1` | Core logic: cookie extraction, API call, XML/CSV generation |
| `DSRF-ExportSetup.bat` | One-time setup: downloads sqlite3, creates folders |
| `ahk\DSRF-Export.ahk` | AHK version: direct paste into WorldShip |
| `ahk\Cookie_Extractor.ahk` | Standalone cookie extraction (AHK) |

## Reference Files

| File | Description |
|------|-------------|
| `ItemVar_Reference.txt` | ItemVar field mappings (IDs to field names) |
| `sqlite3.exe` | SQLite CLI tool (auto-installed by setup) |
| `cookies.txt` | Session cookies (auto-generated, cached) |

## Discovery Tools

| File | Description |
|------|-------------|
| `discover_itemvars.bat` | Interactive itemvar discovery |
| `itemvar_discovery.ps1` | PowerShell itemvar discovery |
| `parse_itemvars.ps1` | Parse itemvar results |
| `extract_firefox_cookies.ps1` | Standalone PowerShell cookie extraction |
