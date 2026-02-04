#Requires AutoHotkey v1
#NoEnv
#Warn
#SingleInstance, Force
SendMode Input
SetWorkingDir %A_ScriptDir%
SetKeyDelay, 25

SetTitleMatchMode, 2
SetDefaultMouseSpeed, 0
CoordMode, Mouse, Window

; =============================================================================
; DSRF-to-UPS_WS-Paste.ahk
; Fetches shipping data from Intra API and pastes directly into UPS WorldShip
; =============================================================================

; UPS WorldShip coordinates (window-relative pixels)
worldShipTitle := "UPS WorldShip"
worldShipTabs := {}
worldShipTabs.Service     := {x: 323, y: 162}
worldShipTabs.ShipFrom    := {x: 99,  y: 162}
worldShipTabs.ShipTo      := {x: 47,  y: 162}

worldShipFields := {}
worldShipFields.Company    := {x: 78,  y: 241}
worldShipFields.SFName     := {x: 85,  y: 280}  ; Ship From Name (same coords, different tab)
worldShipFields.STName     := {x: 85,  y: 280}  ; Ship To Name
worldShipFields.Address1   := {x: 85,  y: 323}
worldShipFields.Address2   := {x: 85,  y: 364}
worldShipFields.PostalCode := {x: 215, y: 403}
worldShipFields.SFPhone    := {x: 85,  y: 485}  ; Ship From Phone (same coords, different tab)
worldShipFields.STPhone    := {x: 85,  y: 485}  ; Ship To Phone
worldShipFields.STEmail    := {x: 210, y: 485}
worldShipFields.Ref2       := {x: 721, y: 345}
worldShipFields.DeclValue  := {x: 721, y: 273}

; Service type dropdown menu coordinates
worldShipServiceMenu := {}
worldShipServiceMenu.Selection := {x: 400, y: 231}       ; Dropdown button
worldShipServiceMenu.NextDayAir := {x: 400, y: 262}
worldShipServiceMenu.NextDayAirSaver := {x: 400, y: 278}
worldShipServiceMenu.SecondDayAir := {x: 400, y: 308}
worldShipServiceMenu.ThreeDaySelect := {x: 400, y: 323}
worldShipServiceMenu.Ground := {x: 400, y: 338}

scaleOffClick := {x: 316, y: 559}
scaleClickDone := false

return ; end of auto-execute

; Ctrl+Alt+D: Fetch DSRF data from API and paste to WorldShip
#If (WinActive("Intra: Shipping Request Form") || WinActive("UPS WorldShip"))
^!d::
    scaleClickDone := false
    startTick := A_TickCount

    ; Find all DSRF windows across browsers
    ToolTip, Scanning for DSRF windows...
    dsrfWindows := FindDSRFWindows()

    if (dsrfWindows.Length() = 0)
    {
        ToolTip
        MsgBox, 48, Error, No Shipping Request Form windows found.`n`nOpen a DSRF page in Firefox, Chrome, or Edge.
        return
    }

    ; Select window (prompt if multiple)
    selected := SelectDSRFWindow(dsrfWindows)
    if (!selected)
    {
        ToolTip
        return
    }

    assetId := selected.pk

    ; Get cookies
    cookies := GetCookies()
    if (!cookies)
    {
        ToolTip
        MsgBox, 48, Error, Could not get session cookies.`n`nCreate cookies.txt in the script folder or set INTRA_COOKIES env var.
        return
    }

    ToolTip, Fetching data for %assetId%...

    ; Call API and get data
    data := FetchDSRFData(assetId, cookies)
    if (!data)
    {
        ToolTip
        MsgBox, 48, Cookie Error, API call failed - cookies are likely invalid or expired.`n`nTo fix:`n1. Open Intra in browser and log in`n2. Press F12 for DevTools`n3. Go to Network tab, refresh page`n4. Click any request, copy Cookie header`n5. Paste into cookies.txt (replace all)
        return
    }

    ToolTip, Pasting to WorldShip...

    ; Paste to WorldShip
    PasteToWorldShip(data)

    ; Show completion
    elapsedMs := A_TickCount - startTick
    elapsedSec := Round(elapsedMs / 1000.0, 2)
    doneMsg := "Done! Runtime: " . elapsedSec . "s`nVerify Service, Dimensions, Weight"
    ToolTip, %doneMsg%
    SetTimer, HideTooltip, -5000
return
#If

HideTooltip:
    ToolTip
return

; =============================================================================
; Functions
; =============================================================================

; Find all DSRF windows across Firefox, Chrome, and Edge
FindDSRFWindows()
{
    local dsrfWindows, windowList, hwnd, procName, pk, i

    dsrfWindows := []
    dsrfTitle := "Intra: Shipping Request Form"

    ; Get all windows matching DSRF title
    WinGet, windowList, List, %dsrfTitle%

    Loop, %windowList%
    {
        hwnd := windowList%A_Index%
        WinGet, procName, ProcessName, ahk_id %hwnd%

        ; Only include browser windows (Firefox, Chrome, Edge)
        if (procName = "firefox.exe" || procName = "chrome.exe" || procName = "msedge.exe")
        {
            ; Get PK# from this window's URL
            pk := GetPKFromWindow(hwnd, procName)
            if (pk != "")
            {
                dsrfWindows.Push({hwnd: hwnd, pk: pk, proc: procName})
            }
        }
    }

    return dsrfWindows
}

; Get PK# from a specific browser window by activating it and copying URL
GetPKFromWindow(hwnd, procName)
{
    local ClipSaved, url, pk

    ClipSaved := ClipboardAll
    Clipboard :=

    ; Activate the specific window
    WinActivate, ahk_id %hwnd%
    WinWaitActive, ahk_id %hwnd%,, 2
    if (ErrorLevel)
    {
        Clipboard := ClipSaved
        return ""
    }

    Sleep 100
    SendInput, ^l  ; Focus address bar (works for all browsers)
    Sleep 100
    SendInput, ^c  ; Copy URL
    ClipWait, 1
    SendInput, {Escape}  ; Close address bar

    url := Clipboard
    Clipboard := ClipSaved

    ; Extract assetId=PK###### from URL
    if (RegExMatch(url, "assetId=(PK\d+)", match))
        return match1
    return ""
}

; Prompt user to select a DSRF window if multiple are open
SelectDSRFWindow(windows)
{
    local prompt, choice, idx, w, i

    if (windows.Length() = 0)
        return ""

    ; If only one window, return it directly
    if (windows.Length() = 1)
        return windows[1]

    ; Build selection prompt for multiple windows
    prompt := "Multiple DSRF windows found:`n`n"
    Loop, % windows.Length()
    {
        w := windows[A_Index]
        browserName := GetBrowserName(w.proc)
        prompt .= A_Index . ": " . w.pk . " (" . browserName . ")`n"
    }
    prompt .= "`nEnter number (1-" . windows.Length() . "):"

    InputBox, choice, Select DSRF Window, %prompt%,, 320, 220
    if (ErrorLevel || choice = "")
        return ""

    idx := Floor(choice)
    if (idx >= 1 && idx <= windows.Length())
        return windows[idx]

    MsgBox, 48, Error, Invalid selection.
    return ""
}

; Get friendly browser name from process name
GetBrowserName(procName)
{
    if (procName = "firefox.exe")
        return "Firefox"
    if (procName = "chrome.exe")
        return "Chrome"
    if (procName = "msedge.exe")
        return "Edge"
    return procName
}

GetCookies()
{
    local cookies, configFile

    ; Try environment variable first
    EnvGet, cookies, INTRA_COOKIES
    if (cookies)
        return cookies

    ; Try cookies.txt in script folder
    configFile := A_ScriptDir . "\cookies.txt"
    if (FileExist(configFile))
    {
        FileRead, cookies, %configFile%
        return Trim(cookies, " `t`r`n")
    }

    return ""
}

FetchDSRFData(assetId, cookies)
{
    local sql, bodyFile, outputFile, jsonBody, psScript, jsonResponse, data, trimmedResponse

    ; Build SQL query
    sql := "declare @profileid int = 1; "
    sql .= "declare @assetid nvarchar(50) = '" . assetId . "'; "
    sql .= "SELECT TOP 1 "
    sql .= "ISNULL(iv148.ItemVarValue,'') as company, "
    sql .= "ISNULL(iv149.ItemVarValue,'') as name, "
    sql .= "ISNULL(iv150.ItemVarValue,'') as address1, "
    sql .= "ISNULL(iv151.ItemVarValue,'') as address2, "
    sql .= "ISNULL(iv152.ItemVarValue,'') as city, "
    sql .= "ISNULL(iv153.ItemVarValue,'') as state, "
    sql .= "ISNULL(iv154.ItemVarValue,'') as postal, "
    sql .= "ISNULL(iv155.ItemVarValue,'') as serviceType, "
    sql .= "ISNULL(iv162.ItemVarValue,'') as declaredValue, "
    sql .= "ISNULL(iv202.ItemVarValue,'') as sfName, "
    sql .= "ISNULL(iv203.ItemVarValue,'') as email, "
    sql .= "ISNULL(iv38.ItemVarValue,'') as sfPhone, "
    sql .= "ISNULL(iv41.ItemVarValue,'') as stPhone "
    sql .= "FROM Asset a "
    sql .= "LEFT JOIN assetitemvars iv148 ON iv148.assetid=a.assetid AND iv148.profileid=1 AND iv148.itemvarid=148 "
    sql .= "LEFT JOIN assetitemvars iv149 ON iv149.assetid=a.assetid AND iv149.profileid=1 AND iv149.itemvarid=149 "
    sql .= "LEFT JOIN assetitemvars iv150 ON iv150.assetid=a.assetid AND iv150.profileid=1 AND iv150.itemvarid=150 "
    sql .= "LEFT JOIN assetitemvars iv151 ON iv151.assetid=a.assetid AND iv151.profileid=1 AND iv151.itemvarid=151 "
    sql .= "LEFT JOIN assetitemvars iv152 ON iv152.assetid=a.assetid AND iv152.profileid=1 AND iv152.itemvarid=152 "
    sql .= "LEFT JOIN assetitemvars iv153 ON iv153.assetid=a.assetid AND iv153.profileid=1 AND iv153.itemvarid=153 "
    sql .= "LEFT JOIN assetitemvars iv154 ON iv154.assetid=a.assetid AND iv154.profileid=1 AND iv154.itemvarid=154 "
    sql .= "LEFT JOIN assetitemvars iv155 ON iv155.assetid=a.assetid AND iv155.profileid=1 AND iv155.itemvarid=155 "
    sql .= "LEFT JOIN assetitemvars iv162 ON iv162.assetid=a.assetid AND iv162.profileid=1 AND iv162.itemvarid=162 "
    sql .= "LEFT JOIN assetitemvars iv202 ON iv202.assetid=a.assetid AND iv202.profileid=1 AND iv202.itemvarid=202 "
    sql .= "LEFT JOIN assetitemvars iv203 ON iv203.assetid=a.assetid AND iv203.profileid=1 AND iv203.itemvarid=203 "
    sql .= "LEFT JOIN assetitemvars iv38 ON iv38.assetid=a.assetid AND iv38.profileid=1 AND iv38.itemvarid=38 "
    sql .= "LEFT JOIN assetitemvars iv41 ON iv41.assetid=a.assetid AND iv41.profileid=1 AND iv41.itemvarid=41 "
    sql .= "WHERE a.AssetID=@assetid AND a.ProfileID=@profileid"

    ; Create temp files
    bodyFile := A_Temp . "\dsrf_body.json"
    outputFile := A_Temp . "\dsrf_response.json"
    FileDelete, %bodyFile%
    FileDelete, %outputFile%

    ; Write JSON body
    jsonBody := "{""Sql"":""" . sql . """}"
    FileAppend, %jsonBody%, %bodyFile%

    ; Build PowerShell command with better error handling
    ; Extract fields directly in PowerShell and output as simple KEY=VALUE format
    psScript := "$headers = @{"
    psScript .= "'Content-Type'='application/json'; "
    psScript .= "'Cookie'='" . cookies . "'"
    psScript .= "}; "
    psScript .= "try { "
    psScript .= "$response = Invoke-WebRequest -Uri 'https://amazonmailservices.us.spsprod.net/IntraWeb/api/automation/then/executeQuery' "
    psScript .= "-Method POST -Headers $headers -Body (Get-Content '" . bodyFile . "' -Raw) -UseBasicParsing; "
    psScript .= "if ($response.StatusCode -ne 200) { exit 2 }; "
    psScript .= "$content = $response.Content; "
    psScript .= "if ($content -match '<!DOCTYPE' -or $content -match '<html' -or $content -match '<form.*login' -or $content -match 'SAMLRequest') { exit 3 }; "
    psScript .= "$json = $content | ConvertFrom-Json; "
    psScript .= "if ($json -eq $null -or $json.Count -eq 0) { exit 4 }; "
    psScript .= "$row = $json[0]; "
    psScript .= "$out = @(); "
    psScript .= "$out += 'name=' + $row.name; "
    psScript .= "$out += 'company=' + $row.company; "
    psScript .= "$out += 'address1=' + $row.address1; "
    psScript .= "$out += 'address2=' + $row.address2; "
    psScript .= "$out += 'city=' + $row.city; "
    psScript .= "$out += 'state=' + $row.state; "
    psScript .= "$out += 'postal=' + $row.postal; "
    psScript .= "$out += 'serviceType=' + $row.serviceType; "
    psScript .= "$out += 'declaredValue=' + $row.declaredValue; "
    psScript .= "$out += 'sfName=' + $row.sfName; "
    psScript .= "$out += 'email=' + $row.email; "
    psScript .= "$out += 'sfPhone=' + $row.sfPhone; "
    psScript .= "$out += 'stPhone=' + $row.stPhone; "
    psScript .= "$out | Out-File -Encoding ASCII '" . outputFile . "'"
    psScript .= "} catch { exit 1 }"

    ; Run PowerShell
    RunWait, powershell -Command "%psScript%",, Hide
    psExitCode := ErrorLevel

    if (psExitCode != 0)
        return ""

    ; Read and parse response (KEY=VALUE format)
    if (!FileExist(outputFile))
        return ""

    FileRead, fileContent, %outputFile%
    if (fileContent = "")
        return ""

    ; Parse KEY=VALUE lines
    data := {}
    Loop, Parse, fileContent, `n, `r
    {
        line := A_LoopField
        if (line = "")
            continue
        eqPos := InStr(line, "=")
        if (eqPos > 0)
        {
            key := SubStr(line, 1, eqPos - 1)
            val := SubStr(line, eqPos + 1)
            data[key] := val
        }
    }

    ; Validate we got actual shipping data (at least one address field must be present)
    if (data.name = "" && data.company = "" && data.address1 = "")
        return ""

    return data
}

ExtractJsonField(json, fieldName)
{
    local pattern, match

    ; Simple regex to extract JSON field value
    pattern := """" . fieldName . """\s*:\s*""([^""]*?)"""
    if (RegExMatch(json, pattern, match))
        return match1
    return ""
}

PasteToWorldShip(data)
{
    global worldShipTitle, worldShipTabs, worldShipFields
    ; companyName and delay are implicitly local when global is declared

    ; Activate WorldShip
    WinActivate, %worldShipTitle%
    WinWaitActive, %worldShipTitle%,, 2
    if (ErrorLevel)
    {
        MsgBox, 48, Error, Could not activate UPS WorldShip window.
        return
    }

    DisableWorldShipScale()
    EnsureWorldShipTop()
    Sleep 120

    ; =========================================================================
    ; SHIP FROM TAB
    ; =========================================================================
    if (!EnsureWorldShipActive())
    {
        MsgBox, 48, Focus Lost, WorldShip lost focus. Aborting.
        return
    }

    ; Click Ship From tab
    MouseClick, left, % worldShipTabs.ShipFrom.x, worldShipTabs.ShipFrom.y
    Sleep 120

    ; Ship From Company (use sfName)
    if (data.sfName != "")
    {
        PasteFieldAt(worldShipFields.Company.x, worldShipFields.Company.y, data.sfName)
        Sleep 120

        ; Wait for company autocomplete
        EnsureWorldShipActive()
        MouseClick, left, % worldShipFields.Ref2.x, worldShipFields.Ref2.y
        Sleep 5000

        ; Ship From Name (Attention)
        EnsureWorldShipTop()
        PasteFieldAt(worldShipFields.SFName.x, worldShipFields.SFName.y, data.sfName)
        Sleep 120

        ; Ship From Phone
        if (data.sfPhone != "")
        {
            PasteFieldAt(worldShipFields.SFPhone.x, worldShipFields.SFPhone.y, data.sfPhone)
            Sleep 120
        }
    }

    ; =========================================================================
    ; SHIP TO TAB
    ; =========================================================================
    if (!EnsureWorldShipActive())
    {
        MsgBox, 48, Focus Lost, WorldShip lost focus. Aborting before Ship To tab.
        return
    }

    ; Click Ship To tab
    MouseClick, left, % worldShipTabs.ShipTo.x, worldShipTabs.ShipTo.y
    Sleep 120

    ; Company (use company if available, otherwise use name)
    companyName := data.company != "" ? data.company : data.name
    PasteFieldAt(worldShipFields.Company.x, worldShipFields.Company.y, companyName)
    Sleep 120

    ; Verify focus before clicking away for autocomplete
    EnsureWorldShipActive()

    ; Wait for company autocomplete
    MouseClick, left, % worldShipFields.Ref2.x, worldShipFields.Ref2.y
    delay := (data.company != "") ? 2000 : 5000
    Sleep, %delay%

    ; Name (Attention)
    PasteFieldAt(worldShipFields.STName.x, worldShipFields.STName.y, data.name)
    Sleep 120

    ; Address 1
    PasteFieldAt(worldShipFields.Address1.x, worldShipFields.Address1.y, data.address1)
    Sleep 120

    ; Address 2
    PasteFieldAt(worldShipFields.Address2.x, worldShipFields.Address2.y, data.address2)
    Sleep 120

    ; Postal Code (triggers city/state autofill)
    PastePostalCode(data.postal, delay)

    ; Ship To Phone
    if (data.stPhone != "")
    {
        PasteFieldAt(worldShipFields.STPhone.x, worldShipFields.STPhone.y, data.stPhone)
        Sleep 120
    }

    ; Email (Ship To tab)
    if (data.email != "")
    {
        PasteFieldAt(worldShipFields.STEmail.x, worldShipFields.STEmail.y, data.email)
        Sleep 120
    }

    ; =========================================================================
    ; SERVICE TAB
    ; =========================================================================
    if (!EnsureWorldShipActive())
    {
        MsgBox, 48, Focus Lost, WorldShip lost focus. Aborting before Service tab.
        return
    }

    ; Click Service tab for Declared Value and Reference fields
    MouseClick, left, % worldShipTabs.Service.x, worldShipTabs.Service.y
    Sleep 150

    ; Declared Value
    if (data.declaredValue != "")
    {
        PasteFieldAt(worldShipFields.DeclValue.x, worldShipFields.DeclValue.y, data.declaredValue)
        Sleep 120
    }

    ; Reference Number 2 (Ship From Name for tracking reference)
    if (data.sfName != "")
    {
        PasteFieldAt(worldShipFields.Ref2.x, worldShipFields.Ref2.y, data.sfName)
        Sleep 120
    }

    ; Select Service Type from dropdown
    SelectServiceType(data.serviceType)

    ; Re-enable electronic scale after all paste operations complete
    EnableWorldShipScale()
}

PasteFieldAt(x, y, text)
{
    local ClipSaved

    if (text = "")
        return

    ; Verify WorldShip is active before pasting
    if (!EnsureWorldShipActive())
    {
        MsgBox, 48, Focus Lost, WorldShip lost focus and could not be re-activated.`nPaste operation aborted.
        return
    }

    ClipSaved := ClipboardAll
    Clipboard := text

    ; Click field, clear existing content, then paste
    ; Note: Ctrl+A doesn't work in UPS WorldShip fields
    ; Must use End then Ctrl+Shift+Home to select all
    MouseClick, left, %x%, %y%
    Sleep 150
    SendInput, {End}
    Sleep 80
    SendInput, ^+{Home}  ; Ctrl+Shift+Home to select all
    Sleep 80
    SendInput, {Delete}
    Sleep 120
    SendInput, ^v
    Sleep 100

    Clipboard := ClipSaved
    ClipSaved := ""
}

PastePostalCode(postalCode, ref2Delay := 5000)
{
    global worldShipFields
    ; ClipSaved is implicitly local when global is declared

    if (postalCode = "")
        return

    ; Verify WorldShip is active before pasting
    if (!EnsureWorldShipActive())
    {
        MsgBox, 48, Focus Lost, WorldShip lost focus and could not be re-activated.`nPaste operation aborted.
        return
    }

    ClipSaved := ClipboardAll
    Clipboard := postalCode

    ; Click postal code field and clear it completely
    ; Note: Ctrl+A doesn't work in UPS WorldShip fields
    ; Must use End then Ctrl+Shift+Home to select all
    MouseClick, left, % worldShipFields.PostalCode.x, worldShipFields.PostalCode.y
    Sleep 150
    SendInput, {End}
    Sleep 80
    SendInput, ^+{Home}  ; Ctrl+Shift+Home to select all
    Sleep 80
    SendInput, {Delete}
    Sleep 250

    ; Type the postal code
    SendInput, %postalCode%
    Sleep 250

    ; Verify focus again before clicking away
    EnsureWorldShipActive()

    ; Click away to trigger city/state autofill
    MouseClick, left, % worldShipFields.Ref2.x, worldShipFields.Ref2.y
    Sleep 500

    ; Handle the "State/Province/County field was automatically changed" dialog
    ; This dialog appears when postal code changes location - just press Enter to dismiss
    Loop, 3
    {
        if WinExist("UPS WorldShip ahk_class #32770")  ; Standard dialog class
        {
            SendInput, {Enter}
            Sleep 200
        }
        else
            break
        Sleep 200
    }

    Sleep, %ref2Delay%

    Clipboard := ClipSaved
    ClipSaved := ""
}

EnsureWorldShipTop()
{
    MouseClick, left, 430, 335
    Sleep 150
    SendInput, {WheelUp 5}
    Sleep 200
}

; Verify WorldShip is active, re-activate if not
; Returns true if WorldShip is active, false if we couldn't activate it
EnsureWorldShipActive()
{
    global worldShipTitle

    ; Check if WorldShip is already active
    if WinActive(worldShipTitle)
        return true

    ; Try to re-activate WorldShip
    WinActivate, %worldShipTitle%
    WinWaitActive, %worldShipTitle%,, 1

    if (ErrorLevel)
    {
        ; Second attempt
        Sleep 100
        WinActivate, %worldShipTitle%
        WinWaitActive, %worldShipTitle%,, 1

        if (ErrorLevel)
            return false
    }

    Sleep 100
    return true
}

DisableWorldShipScale()
{
    global scaleOffClick, scaleClickDone
    if (scaleClickDone)
        return
    Sleep 150
    MouseClick, left, % scaleOffClick.x, scaleOffClick.y
    Sleep 250
    scaleClickDone := true
}

EnableWorldShipScale()
{
    global scaleOffClick, scaleClickDone
    if (!scaleClickDone)
        return
    Sleep 150
    MouseClick, left, % scaleOffClick.x, scaleOffClick.y
    Sleep 250
    scaleClickDone := false
}

; Select service type from dropdown menu using direct clicks
; Uses coordinates from UPS_WS_Shortcuts.ahk - no hotkey conflicts
SelectServiceType(serviceType)
{
    global worldShipTabs, worldShipServiceMenu

    if (serviceType = "")
        return

    if (!EnsureWorldShipActive())
        return

    ; Click Service tab
    MouseClick, left, % worldShipTabs.Service.x, % worldShipTabs.Service.y
    Sleep 50

    ; Click dropdown to open menu
    MouseClick, left, % worldShipServiceMenu.Selection.x, % worldShipServiceMenu.Selection.y
    Sleep 250

    ; Determine which option to click based on service type
    ; Check more specific matches first (e.g., "Next Day Air Saver" before "Next Day Air")
    targetY := 0

    if (InStr(serviceType, "Ground"))
        targetY := worldShipServiceMenu.Ground.y
    else if (InStr(serviceType, "Next Day Air Saver"))
        targetY := worldShipServiceMenu.NextDayAirSaver.y
    else if (InStr(serviceType, "Next Day Air"))
        targetY := worldShipServiceMenu.NextDayAir.y
    else if (InStr(serviceType, "2nd Day") || InStr(serviceType, "Second Day"))
        targetY := worldShipServiceMenu.SecondDayAir.y
    else if (InStr(serviceType, "3 Day") || InStr(serviceType, "3-Day"))
        targetY := worldShipServiceMenu.ThreeDaySelect.y

    if (targetY > 0)
    {
        MouseMove, % worldShipServiceMenu.Selection.x, % targetY
        Sleep 250
        MouseClick, left
        Sleep 150
    }
}
