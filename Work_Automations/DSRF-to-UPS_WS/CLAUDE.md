# DSRF-to-UPS_WS Scripts

**Location:** `Work_Automations/DSRF-to-UPS_WS/`

This folder contains AutoHotkey v1 scripts that automate transferring data from the Intra: Shipping Request Form (DSRF) web page to UPS WorldShip desktop application.

## Launcher Integration

These scripts are launched via `Launcher_Script(WorkAutos).ahk` in the parent `Work_Automations/` directory:

| Hotkey | Action | Script Launched |
|--------|--------|-----------------|
| `Ctrl+Alt+C` | Launch DSRF transfer (auto-detects browser) | `DSRF-to-UPS_WS.ahk` (Firefox) or `DSRF-to-UPS_WS-Legacy.ahk` (Chrome/Edge) |
| `Ctrl+Alt+U` | Launch Super-Speed version | `DSRF-to-UPS_WS(Super-Speed).ahk` (Firefox) or Legacy (Chrome/Edge) |

**Launcher behavior:**
1. Checks if Intra: Shipping Request Form is open in Firefox/Chrome/Edge (priority order)
2. Closes any previously running DSRF scripts via `CloseWorldShipScripts()`
3. Focuses and resizes the Intra window to standard dimensions (1917, 0, 1530, 1399)
4. Scrolls form to top with `WheelUp 25`
5. Launches appropriate script based on browser (Firefox uses zoom-dependent scripts, Chrome/Edge use Legacy)

## Goal

The primary goal is to develop `DSRF-to-UPS_WS-New_Approach.ahk` which aims to:
- Extract Intra form data into a `.bat` file (or similar intermediate storage)
- Pull form data from there and enter into UPS WorldShip
- Avoid direct copy/paste between windows, reducing timing dependencies and failures

## Scripts Overview

### DSRF-to-UPS_WS.ahk (Most Successful - Reference Implementation)
The primary working script. Uses a "collect-then-paste" architecture:
1. Collects ALL fields from Intra into a payload object first
2. Then pastes everything to WorldShip in sequence
3. Supports both Business (`Ctrl+Alt+B`) and Personal (`Ctrl+Alt+P`) form modes

**Key Success Factors:**
- Browser zoom to 60% (`Ctrl+0` then 4x `Ctrl+WheelDown`) for consistent coordinates
- Window normalization to fixed size (1917, 0, 1530, 1399)
- Neutral click + `Ctrl+Home` before field operations to reset scroll/focus
- Clipboard preservation pattern (save → clear → copy → restore)

### DSRF-to-UPS_WS(Super-Speed).ahk
Optimized version (~11 sec faster) using:
- `FastSleep()` function with configurable scale factor (0.75x)
- `WaitForAddressFill()` polling instead of fixed delays
- `CaptureAddressSnapshot()` to detect when WorldShip autofill completes
- Shorter ClipWait timeouts with retry logic

### DSRF-to-UPS_WS-Legacy.ahk
For Chrome/Edge when Firefox is unavailable (no zoom dependency). Uses different coordinate offsets and scrolling strategy.

### DSRF-to-UPS_WS-NewScale(Firefox+Legacy).ahk
Alternative approach with explicit scrolling between field groups. Contains more verbose, step-by-step code structure.

## Critical Patterns & Techniques

### Window Management
```autohotkey
; Multi-browser support (priority order)
intraWinExes := ["firefox.exe", "chrome.exe", "msedge.exe"]

GetIntraWindowTitle() {
    Loop % intraWinExes.Length() {
        candidate := intraWinTitle " ahk_exe " intraWinExes[A_Index]
        if (WinExist(candidate))
            return candidate
    }
    return intraWinTitle
}

; Fixed window positioning for coordinate consistency
EnsureIntraWindow() {
    WinMove, %title%,, 1917, 0, 1530, 1399
}
```

### Clipboard Operations
```autohotkey
CopyFieldAt(x, y) {
    ClipSaved := ClipboardAll  ; Save full clipboard (text + formats)
    Clipboard :=               ; Clear clipboard
    MouseClick, left, %x%, %y%
    Sleep 150
    SendInput, ^a
    Sleep 80
    SendInput, ^c
    ClipWait, 0.5             ; Wait for clipboard with timeout
    text := Clipboard
    Clipboard := ClipSaved     ; Restore original clipboard
    ClipSaved := ""
    return text
}
```

### WorldShip Field Paste (Ctrl+A doesn't work)
```autohotkey
PasteFieldAt(x, y, text) {
    if (text = "")
        return
    MouseClick, left, %x%, %y%
    Sleep 150
    SendInput, {Home}          ; WorldShip ignores Ctrl+A
    Sleep 80
    SendInput, +{End}          ; So use Home → Shift+End → Delete
    Sleep 80
    SendInput, {Delete}
    Sleep 120
    SendInput, ^v
}
```

### WorldShip Address Book Delays
WorldShip has an address book autofill that triggers after pasting Company name:
- If company exists in address book: ~2000ms delay sufficient
- If new/unknown company: ~5000ms delay needed
- Super-Speed version polls City field to detect completion

### Electronic Scale Disable
```autohotkey
DisableWorldShipScale() {
    global scaleOffClick, scaleClickDone
    if (scaleClickDone)
        return
    MouseClick, left, % scaleOffClick.x, scaleOffClick.y
    scaleClickDone := true
}
```

### Neutral Click Pattern
Always click a neutral area and scroll to top before field operations:
```autohotkey
NeutralAndHome() {
    global neutralClickR
    WinGetPos,,, winW, winH, A
    targetX := Floor(winW * neutralClickR.x)
    targetY := Floor(winH * neutralClickR.y)
    MouseClick, left, %targetX%, %targetY%
    Sleep 200
    SendInput, ^{Home}
    Sleep 200
}
```

### Declared Value Field Navigation
The Declared Value field can't be reliably clicked directly. Tab from PostalCode:
```autohotkey
MouseClick, left, % fields.PostalCode.x, fields.PostalCode.y
Loop 6 {
    Sleep 50
    SendInput, {Tab}
}
; Validate with regex - retry if wrong field
if (!RegExMatch(value, "^\d{0,5}$")) {
    SendInput, {Tab}
    value := CopyCaretValue()
}
```

## Coordinate Reference

### Intra Fields (at 60% zoom, 1530x1399 window)
| Field | Business X,Y | Personal X,Y |
|-------|-------------|--------------|
| CostCenter | 553, 394 | N/A |
| Alias | 551, 578 | 552, 524 |
| SFName | 754, 572 | 749, 518 |
| SFPhone | 934, 517 | 920, 464 |
| STName | 587, 863 | 598, 809 |
| Company | 887, 863 | 917, 809 |
| Address1 | 587, 1031 | 589, 976 |
| Address2 | 543, 1084 | 548, 1030 |
| STPhone | 878, 1031 | 891, 976 |
| PostalCode | 969, 1085 | 1000, 1031 |

### WorldShip Fields
| Field | X, Y |
|-------|------|
| Company | 78, 241 |
| Name/Attn | 85, 280 |
| Address1 | 85, 323 |
| Address2 | 85, 364 |
| PostalCode | 215, 403 |
| City | 85, 445 |
| Phone | 85, 485 |
| Email | 210, 485 |
| Ref1 | 721, 309 |
| Ref2 | 721, 345 |
| DeclVal | 721, 273 |

### WorldShip Tabs
| Tab | X, Y |
|-----|------|
| ShipTo | 47, 162 |
| ShipFrom | 99, 162 |
| Service | 323, 162 |
| Options | 372, 162 |
| QVN | 381, 282 |
| Recipients | 560, 253 |

## Common Issues & Solutions

1. **Coordinates off**: Verify window size is exactly 1530x1399 and zoom is 60%
2. **Clipboard empty**: Increase Sleep after Ctrl+C, check ClipWait timeout
3. **WorldShip autofill interference**: Wait for address book lookup (5s for new, 2s for known)
4. **Scale lag**: Call `DisableWorldShipScale()` once per session
5. **Field focus wrong**: Use neutral click + Ctrl+Home before each operation

## New Approach Ideas (for DSRF-to-UPS_WS-New_Approach.ahk)

Potential strategies to explore:
1. **File-based data transfer**: Extract all fields to temp file, read from file when pasting
2. **COM automation**: Use WorldShip's COM interface if available
3. **OCR-based extraction**: Screen capture + OCR instead of clipboard
4. **Browser automation**: Use browser DevTools protocol to read form values directly
5. **Intermediate GUI**: Show captured data for verification before pasting

## Hotkeys

| Hotkey | Action |
|--------|--------|
| `Ctrl+Alt+B` | Business form transfer |
| `Ctrl+Alt+P` | Personal form transfer |
| `Esc` | Exit script |

---

## DSRF-Export Subfolder (API-Based Approach)

**Location:** `DSRF-to-UPS_WS/DSRF-Export/`

This subfolder contains scripts that use the Intra executeQuery API to fetch shipping data directly, bypassing screen scraping entirely. This is the "New Approach" implementation.

### Files

| File | Purpose |
|------|---------|
| `DSRF-Export.bat` | Standalone batch script (no AHK) - prompts for PK#, exports CSV |
| `DSRF-to-UPS_WS-Paste.ahk` | AHK script - auto-detects DSRF windows, fetches via API, pastes to WorldShip |
| `cookies.txt` | Session cookies for API authentication (user creates) |
| `cookies.txt.template` | Template showing cookie format |
| `README.txt` | Instructions for batch version |
| `README(AHK ver).md` | Instructions for AHK version |

### API Details

**Endpoint:** `https://amazonmailservices.us.spsprod.net/IntraWeb/api/automation/then/executeQuery`

**Method:** POST with JSON body containing SQL query

**Auth:** Session cookies copied from browser DevTools (F12 → Network → Cookie header)

### Field Mapping (itemVar → WorldShip)

| itemVar | Field |
|---------|-------|
| 148 | Company |
| 149 | Recipient Name |
| 150 | Address Line 1 |
| 151 | Address Line 2 |
| 152 | City |
| 153 | State |
| 154 | Postal Code |
| 155 | Service Type |

### Advantages Over Screen Scraping

1. **No coordinate dependencies** - Data fetched directly from database
2. **Multi-browser support** - Firefox, Chrome, Edge all work
3. **More reliable** - No timing/zoom/window size issues for data extraction
4. **Distributable** - Batch version requires no AHK installation

### Hotkeys (AHK Version)

| Hotkey | Action |
|--------|--------|
| `Ctrl+Alt+D` | Fetch DSRF via API and paste to WorldShip |

### Cookie Validation

The script validates API responses before pasting:
- Detects HTML login pages (expired session)
- Validates JSON format
- Checks for error indicators
- Verifies shipping data is present

If cookies are invalid, shows clear error message instead of pasting garbage data.

### Multi-Window Behavior

1. Script scans for "Intra: Shipping Request Form" windows in all browsers
2. If multiple found, prompts user to select which PK#
3. Extracts PK# from URL, calls API, pastes to WorldShip

### Cookie Setup

1. Open browser DevTools (F12) on any Intra page
2. Go to Network tab, refresh page
3. Click any request, find "Cookie:" header
4. Copy entire cookie string to `cookies.txt` (single line)
