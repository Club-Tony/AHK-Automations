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
worldShipFields.STName     := {x: 85,  y: 280}
worldShipFields.Address1   := {x: 85,  y: 323}
worldShipFields.Address2   := {x: 85,  y: 364}
worldShipFields.PostalCode := {x: 215, y: 403}
worldShipFields.STPhone    := {x: 85,  y: 485}
worldShipFields.Ref2       := {x: 721, y: 345}

scaleOffClick := {x: 316, y: 559}
scaleClickDone := false

return ; end of auto-execute

; Ctrl+Alt+D: Fetch DSRF data from API and paste to WorldShip
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
    sql .= "ISNULL(iv155.ItemVarValue,'') as serviceType "
    sql .= "FROM Asset a "
    sql .= "LEFT JOIN assetitemvars iv148 ON iv148.assetid=a.assetid AND iv148.profileid=1 AND iv148.itemvarid=148 "
    sql .= "LEFT JOIN assetitemvars iv149 ON iv149.assetid=a.assetid AND iv149.profileid=1 AND iv149.itemvarid=149 "
    sql .= "LEFT JOIN assetitemvars iv150 ON iv150.assetid=a.assetid AND iv150.profileid=1 AND iv150.itemvarid=150 "
    sql .= "LEFT JOIN assetitemvars iv151 ON iv151.assetid=a.assetid AND iv151.profileid=1 AND iv151.itemvarid=151 "
    sql .= "LEFT JOIN assetitemvars iv152 ON iv152.assetid=a.assetid AND iv152.profileid=1 AND iv152.itemvarid=152 "
    sql .= "LEFT JOIN assetitemvars iv153 ON iv153.assetid=a.assetid AND iv153.profileid=1 AND iv153.itemvarid=153 "
    sql .= "LEFT JOIN assetitemvars iv154 ON iv154.assetid=a.assetid AND iv154.profileid=1 AND iv154.itemvarid=154 "
    sql .= "LEFT JOIN assetitemvars iv155 ON iv155.assetid=a.assetid AND iv155.profileid=1 AND iv155.itemvarid=155 "
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

    ; Click Ship To tab
    MouseClick, left, % worldShipTabs.ShipTo.x, worldShipTabs.ShipTo.y
    Sleep 120

    ; Company (use company if available, otherwise use name)
    companyName := data.company != "" ? data.company : data.name
    PasteFieldAt(worldShipFields.Company.x, worldShipFields.Company.y, companyName)
    Sleep 120

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

    ; Phone (if available - not in current API mapping but placeholder)
    ; PasteFieldAt(worldShipFields.STPhone.x, worldShipFields.STPhone.y, data.phone)

    ; Click Service tab to finish
    MouseClick, left, % worldShipTabs.Service.x, worldShipTabs.Service.y
    Sleep 150
}

PasteFieldAt(x, y, text)
{
    local ClipSaved

    if (text = "")
        return

    ClipSaved := ClipboardAll
    Clipboard := text

    ; Click field, select all, delete, then paste
    MouseClick, left, %x%, %y%
    Sleep 150
    SendInput, ^a  ; Select all text in field
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

    ClipSaved := ClipboardAll
    Clipboard := postalCode

    ; Click postal code field and clear it completely
    MouseClick, left, % worldShipFields.PostalCode.x, worldShipFields.PostalCode.y
    Sleep 150
    SendInput, ^a  ; Select all
    Sleep 80
    SendInput, {Delete}
    Sleep 250

    ; Type the postal code
    SendInput, %postalCode%
    Sleep 250

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
