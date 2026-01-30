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

Esc::ExitApp

; Ctrl+Alt+D: Fetch DSRF data from API and paste to WorldShip
^!d::
    scaleClickDone := false
    startTick := A_TickCount

    ; Get asset ID from Firefox URL
    assetId := GetAssetIdFromFirefox()
    if (!assetId)
    {
        MsgBox, 48, Error, Could not extract PK# from Firefox URL.`n`nMake sure Firefox has a Shipping Request Form page open.
        return
    }

    ; Get cookies
    cookies := GetCookies()
    if (!cookies)
    {
        MsgBox, 48, Error, Could not get session cookies.`n`nCreate cookies.txt in the script folder or set INTRA_COOKIES env var.
        return
    }

    ToolTip, Fetching data for %assetId%...

    ; Call API and get data
    data := FetchDSRFData(assetId, cookies)
    if (!data)
    {
        ToolTip
        MsgBox, 48, Error, API call failed or no data returned.`n`nCheck that cookies are valid and not expired.
        return
    }

    ToolTip, Pasting to WorldShip...

    ; Paste to WorldShip
    PasteToWorldShip(data)

    ; Show completion
    elapsedMs := A_TickCount - startTick
    elapsedSec := Round(elapsedMs / 1000.0, 2)
    ToolTip, Done! Runtime: %elapsedSec%s`nVerify Service, Dimensions, Weight
    SetTimer, HideTooltip, -5000
return

HideTooltip:
    ToolTip
return

; =============================================================================
; Functions
; =============================================================================

GetAssetIdFromFirefox()
{
    ; Save current clipboard
    ClipSaved := ClipboardAll
    Clipboard :=

    ; Activate Firefox and get URL
    WinActivate, ahk_class MozillaWindowClass
    WinWaitActive, ahk_class MozillaWindowClass,, 2
    if (ErrorLevel)
    {
        Clipboard := ClipSaved
        return ""
    }

    Sleep 100
    SendInput, ^l  ; Focus address bar
    Sleep 100
    SendInput, ^c  ; Copy URL
    ClipWait, 1
    SendInput, {Escape}  ; Close address bar

    url := Clipboard
    Clipboard := ClipSaved
    ClipSaved := ""

    ; Extract assetId=PK###### from URL
    if (RegExMatch(url, "assetId=(PK\d+)", match))
        return match1
    return ""
}

GetCookies()
{
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

    ; Build PowerShell command
    psScript := "$headers = @{"
    psScript .= "'Content-Type'='application/json'; "
    psScript .= "'Cookie'='" . cookies . "'"
    psScript .= "}; "
    psScript .= "try { "
    psScript .= "$response = Invoke-RestMethod -Uri 'https://amazonmailservices.us.spsprod.net/IntraWeb/api/automation/then/executeQuery' "
    psScript .= "-Method POST -Headers $headers -Body (Get-Content '" . bodyFile . "' -Raw); "
    psScript .= "$response | ConvertTo-Json -Depth 10 | Out-File -Encoding UTF8 '" . outputFile . "'"
    psScript .= "} catch { exit 1 }"

    ; Run PowerShell
    RunWait, powershell -Command "%psScript%",, Hide

    if (ErrorLevel != 0)
        return ""

    ; Read and parse response
    if (!FileExist(outputFile))
        return ""

    FileRead, jsonResponse, %outputFile%
    if (jsonResponse = "")
        return ""

    ; Parse JSON using regex (simple extraction)
    data := {}
    data.name := ExtractJsonField(jsonResponse, "name")
    data.company := ExtractJsonField(jsonResponse, "company")
    data.address1 := ExtractJsonField(jsonResponse, "address1")
    data.address2 := ExtractJsonField(jsonResponse, "address2")
    data.city := ExtractJsonField(jsonResponse, "city")
    data.state := ExtractJsonField(jsonResponse, "state")
    data.postal := ExtractJsonField(jsonResponse, "postal")
    data.serviceType := ExtractJsonField(jsonResponse, "serviceType")

    ; Validate we got some data
    if (data.name = "" && data.company = "" && data.address1 = "")
        return ""

    return data
}

ExtractJsonField(json, fieldName)
{
    ; Simple regex to extract JSON field value
    pattern := """" . fieldName . """\s*:\s*""([^""]*?)"""
    if (RegExMatch(json, pattern, match))
        return match1
    return ""
}

PasteToWorldShip(data)
{
    global worldShipTitle, worldShipTabs, worldShipFields

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
    if (text = "")
        return

    ClipSaved := ClipboardAll
    Clipboard := text

    MouseClick, left, %x%, %y%
    Sleep 150
    SendInput, {Home}
    Sleep 80
    SendInput, +{End}
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

    if (postalCode = "")
        return

    ClipSaved := ClipboardAll
    Clipboard := postalCode

    MouseClick, left, % worldShipFields.PostalCode.x, worldShipFields.PostalCode.y
    Sleep 150
    SendInput, {Home}
    Sleep 80
    SendInput, +{End}
    Sleep 80
    SendInput, {Delete}
    Sleep 250
    SendInput, %postalCode%
    Sleep 250

    ; Click away to trigger city/state autofill
    MouseClick, left, % worldShipFields.Ref2.x, worldShipFields.Ref2.y
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
