# AGENTS.md

Guidance for coding agents working in this DSRF-Export package.

## Scope
- This folder is a standalone API-based package for DSRF data export and WorldShip automation.
- Two versions:
  - **Standalone** (no AHK): `DSRF-Export.bat` + `DSRF-Export.ps1` → generates WorldShip XML/CSV
  - **AHK**: `ahk\DSRF-Export.ahk` → direct paste into WorldShip via hotkey
- Shared files: `sqlite3.exe`, `cookies.txt`, `ItemVar_Reference.txt`
- Discovery tools: `discover_itemvars.ps1` + `discover_itemvars.bat`
- `cookies.txt` and any copied cookie headers may contain live session data. Treat as secret and never print, log, or commit real values.

## Current behavior
- Both main tools call `.../api/automation/then/executeQuery` with SQL built from `Asset` plus `LEFT JOIN assetitemvars`.
- Shared shipping mapping:
  - itemVar148 -> company
  - itemVar149 -> name
  - itemVar150 -> address1
  - itemVar151 -> address2
  - itemVar152 -> city
  - itemVar153 -> state
  - itemVar154 -> postal
  - itemVar155 -> serviceType
- `DSRF-Export.bat` + `DSRF-Export.ps1` (Standalone):
  - Prompts for PK#.
  - Auto-extracts cookies from Firefox via sqlite3, caches in `cookies.txt`.
  - Calls API with all 15 fields, generates XML + CSV in `import\` subfolder.
- `ahk\DSRF-Export.ahk` (AHK version):
  - Hotkey `Ctrl+Alt+D` (context-sensitive: only when Intra/WorldShip active).
  - Scans Firefox/Chrome/Edge windows titled "Intra: Shipping Request Form".
  - Activates candidate windows, copies URL, extracts `assetId=PK######`, and prompts when multiple PK windows exist.
  - Auto-extracts cookies from Firefox via sqlite3, falls back to cached `cookies.txt`.
  - Pastes all fields into WorldShip tabs and uses postal-code entry to trigger city/state autofill.

## ItemVar discovery workflow
- `discover_itemvars.bat` launches `discover_itemvars.ps1` and outputs `itemvar_discovery.txt`.
- Discovery currently scans itemVars 220-270 using the same `Asset` + `LEFT JOIN` query pattern.
- `itemvar_discovery.txt` is generated analysis output and should not be treated as permanent source-of-truth mapping data.

## Workflow notes
- Do not run network actions or tests unless requested.
- Keep AutoHotkey v1 style with `#Warn`; declare locals in functions.
- If you change SQL or field mappings, update both main scripts (`ahk\DSRF-Export.ahk` and `DSRF-Export.ps1`) and README.md in the same change.
- Preserve response validation safeguards (login-page detection, JSON checks, empty-data checks) before paste/export.
- If you add or change hotkeys or launcher integration, update parent tooltip/launcher files too.
