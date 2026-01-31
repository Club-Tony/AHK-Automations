# AGENTS.md

Guidance for coding agents working in this DSRF-Export package.

## Scope
- This folder is a standalone distribution package for DSRF API exports and the AHK paste helper.
- Primary files:
  - DSRF-Export.bat (standalone CSV export using curl + PowerShell)
  - DSRF-to-UPS_WS-Paste.ahk (AHK hotkey to paste into WorldShip)
  - README.txt and README(AHK ver).md
  - cookies.txt.template
- cookies.txt holds live session cookies. Treat as secret; do not log, print, or commit.

## How it works
- Both tools call the Intra executeQuery API and map itemVar fields to shipping fields.
- DSRF-Export.bat:
  - Prompts for PK#.
  - Reads cookies from cookies.txt (first line only).
  - Writes dsrf_export.csv in this folder.
- DSRF-to-UPS_WS-Paste.ahk:
  - Hotkey Ctrl+Alt+D.
  - Scans Firefox/Chrome/Edge windows for "Intra: Shipping Request Form".
  - Extracts assetId=PK###### from URL, calls API via PowerShell, then pastes into WorldShip.
  - Uses postal code to trigger city/state autofill.

## Workflow notes
- Do not run network actions or tests unless requested.
- Keep AutoHotkey v1 style with #Warn; declare locals in functions.
- If you change SQL or field mappings, update both scripts and the README files.
- If you add or change hotkeys or launcher integration, update the parent tooltip/launcher files.

## External reference
- C:\Users\daveyuan\Desktop\curl_test.bat is a verbose curl example with full headers/cookies for debugging only.
