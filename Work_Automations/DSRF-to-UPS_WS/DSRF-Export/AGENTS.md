# AGENTS.md

Guidance for coding agents working in this DSRF-Export package.

## Scope
- This folder is a standalone API-based package for DSRF data export and WorldShip paste automation.
- Primary files:
  - `DSRF-Export.bat` (standalone CSV export using curl + PowerShell parsing)
  - `DSRF-to-UPS_WS-Paste.ahk` (AHK hotkey script that fetches API data then pastes into WorldShip)
  - `discover_itemvars.ps1` + `discover_itemvars.bat` (helper workflow for finding unmapped itemVar IDs)
  - `README.txt` and `README(AHK ver).md`
  - `cookies.txt.template`
- `cookies.txt`, `curl_test.bat`, and any copied cookie headers may contain live session data. Treat as secret and never print, log, or commit real values.

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
- `DSRF-Export.bat`:
  - Prompts for PK#.
  - Reads first line of `cookies.txt`.
  - Writes request/response temp files in `%TEMP%`, validates response format, and writes `dsrf_export.csv` in this folder.
- `DSRF-to-UPS_WS-Paste.ahk`:
  - Hotkey `Ctrl+Alt+D`.
  - Scans Firefox/Chrome/Edge windows titled "Intra: Shipping Request Form".
  - Activates candidate windows, copies URL, extracts `assetId=PK######`, and prompts when multiple PK windows exist.
  - Cookie lookup order: `INTRA_COOKIES` env var first, then `cookies.txt`.
  - Pastes Ship To fields in WorldShip and uses postal-code entry to trigger city/state autofill.

## ItemVar discovery workflow
- `discover_itemvars.bat` launches `discover_itemvars.ps1` and outputs `itemvar_discovery.txt`.
- Discovery currently scans itemVars 220-270 using the same `Asset` + `LEFT JOIN` query pattern.
- `itemvar_discovery.txt` is generated analysis output and should not be treated as permanent source-of-truth mapping data.

## Workflow notes
- Do not run network actions or tests unless requested.
- Keep AutoHotkey v1 style with `#Warn`; declare locals in functions.
- If you change SQL or field mappings, update both main scripts and both README files in the same change.
- Preserve response validation safeguards (login-page detection, JSON checks, empty-data checks) before paste/export.
- If you add or change hotkeys or launcher integration, update parent tooltip/launcher files too.
