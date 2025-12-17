#Requires AutoHotkey v1
#NoEnv
#Warn
#SingleInstance, Force
SendMode Input
SetWorkingDir %A_ScriptDir%
; For use with Chrome/Edge when Firefox is unavailable (no zoom dependency).
SetKeyDelay, 25

SetTitleMatchMode, 2
SetDefaultMouseSpeed, 0
CoordMode, Mouse, Window
CoordMode, Caret, Window

; Window targets
intraWinTitle := "Intra: Shipping Request Form"
intraWinExes := ["firefox.exe", "chrome.exe", "msedge.exe"]
worldShipTitle := "UPS WorldShip"

; Normalized window size used for ratio/absolute clicks (matches Intra_Buttons)
intraWinW := 1530
intraWinH := 1399
neutralClickR := {x: 0.895425, y: 0.50683} ; same neutral spot used in Intra_Buttons

; Intra: Shipping Request Form field coordinates (window-relative pixels at normalized size)
; Coordinates are the post-scroll positions for Chrome/Edge.
intraFields := {}
intraFields.CostCenter    := {x: 411, y: 587}
intraFields.Alias         := {x: 410, y: 885}
intraFields.SFName        := {x: 800, y: 880}
intraFields.SFPhone       := {x: 1040, y: 788}
intraFields.STName        := {x: 476, y: 1365}
intraFields.Company       := {x: 980, y: 1368}
intraFields.Address1      := {x: 473, y: 545}
intraFields.Address2      := {x: 400, y: 635}
intraFields.STPhone       := {x: 976, y: 542}
intraFields.PostalCode    := {x: 1120, y: 633}
intraFields.DeclaredValue := {x: 410, y: 830} ; note y is 741 if ship from name field closes [Theory]

; UPS WorldShip coordinates (window-relative pixels)
worldShipTabs := {}
worldShipTabs.Service     := {x: 323, y: 162}
worldShipTabs.ShipFrom    := {x: 99,  y: 162}
worldShipTabs.ShipTo      := {x: 47,  y: 162}
worldShipTabs.Options     := {x: 372, y: 162}
worldShipTabs.QVN         := {x: 381, y: 282}
worldShipTabs.Recipients  := {x: 560, y: 253}
worldShipTabs.QVNEmail    := {x: 414, y: 103}

worldShipFields := {}
worldShipFields.SFName     := {x: 85,  y: 280}
worldShipFields.STName     := {x: 85,  y: 280}
worldShipFields.SFPhone    := {x: 85,  y: 485}
worldShipFields.STPhone    := {x: 85,  y: 485}
worldShipFields.STEmail    := {x: 210, y: 485}
worldShipFields.SFAttn     := {x: 85,  y: 280}
worldShipFields.Company    := {x: 78,  y: 241}
worldShipFields.Address1   := {x: 85,  y: 323}
worldShipFields.Address2   := {x: 85,  y: 364}
worldShipFields.PostalCode := {x: 215, y: 403}
worldShipFields.Ref1       := {x: 721, y: 309}
worldShipFields.Ref2       := {x: 721, y: 345}
worldShipFields.DeclVal    := {x: 721, y: 273}
scaleOffClick := {x: 316, y: 559} ; click to disable electronic scale lag

; Button targets (window-relative pixels)
PersonalButtonX := 274
PersonalButtonY := 486
BusinessButtonX := 365
BusinessButtonY := 486

return  ; end of auto-execute section

Esc::ExitApp

^!b:: ; Business Form: collect first, then paste
    scaleClickDone := false
    startTick := A_TickCount
    data := CollectIntraDataLegacy("business")
    PasteBusinessToWorldShip(data)
    FocusIntraWindow()
    EnsureIntraWindow()
    ShowHotkeyRuntime(startTick)
return

^!p:: ; Personal Form: collect first, then paste
    scaleClickDone := false
    startTick := A_TickCount
    data := CollectIntraDataLegacy("personal")
    PastePersonalToWorldShip(data)
    FocusIntraWindow()
    EnsureIntraWindow()
    ShowHotkeyRuntime(startTick)
return

CollectIntraDataLegacy(mode)
{
    global intraFields, BusinessButtonX, BusinessButtonY, PersonalButtonX, PersonalButtonY
    offsetY := (mode = "personal") ? -90 : 0
    includeCostCenter := (mode = "business")

    FocusIntraWindow()
    EnsureIntraWindow()
    Sleep 250
    NeutralAndHome()
    Sleep 80

    if (mode = "business")
    {
        SendInput, {WheelUp 15}
        Sleep 120
        MouseClick, left, %BusinessButtonX%, %BusinessButtonY%, 2
    }
    else
    {
        MouseClick, left, %PersonalButtonX%, %PersonalButtonY%, 2
    }
    Sleep 150

    payload := {}
    if (includeCostCenter)
        payload.CostCenter := CopyFieldAt(intraFields.CostCenter.x, intraFields.CostCenter.y + offsetY)
    payload.Alias      := CopyFieldAt(intraFields.Alias.x, intraFields.Alias.y + offsetY)
    payload.SFName     := CopyFieldAt(intraFields.SFName.x, intraFields.SFName.y + offsetY)
    payload.SFPhone    := CopyFieldAt(intraFields.SFPhone.x, intraFields.SFPhone.y + offsetY)
    payload.Company    := CopyFieldAt(intraFields.Company.x, intraFields.Company.y + offsetY)
    payload.STName     := CopyFieldAt(intraFields.STName.x, intraFields.STName.y + offsetY)

    ; Scroll to Ship To block before grabbing the lower fields
    ScrollDownToShipToLegacy()

    payload.Address1   := CopyFieldAt(intraFields.Address1.x, intraFields.Address1.y + offsetY)
    payload.Address2   := CopyFieldAt(intraFields.Address2.x, intraFields.Address2.y + offsetY)
    payload.STPhone    := CopyFieldAt(intraFields.STPhone.x, intraFields.STPhone.y + offsetY)
    payload.PostalCode := CopyFieldAt(intraFields.PostalCode.x, intraFields.PostalCode.y + offsetY)
    payload.DeclaredValue := CopyDeclaredValueLegacy(intraFields, offsetY)

    FocusWorldShipWindow()
    return payload
}

PasteBusinessToWorldShip(data)
{
    global worldShipTabs, worldShipFields
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    Sleep 120

    MouseClick, left, % worldShipTabs.ShipTo.x, worldShipTabs.ShipTo.y
    Sleep 120
    MouseClick, left, % worldShipTabs.Service.x, worldShipTabs.Service.y
    Sleep 120
    PasteFieldAt(worldShipFields.Ref1.x, worldShipFields.Ref1.y, data.CostCenter)
    Sleep 120

    MouseClick, left, % worldShipTabs.Service.x, worldShipTabs.Service.y
    Sleep 120
    MouseClick, left, % worldShipTabs.ShipFrom.x, worldShipTabs.ShipFrom.y
    Sleep 120
    PasteFieldAt(worldShipFields.Company.x, worldShipFields.Company.y, data.SFName)
    Sleep 120
    MouseClick, left, % worldShipFields.Ref2.x, worldShipFields.Ref2.y
    Sleep 5000
    PasteFieldAt(worldShipFields.SFName.x, worldShipFields.SFName.y, data.SFName)
    Sleep 120
    EnsureWorldShipTop()
    PasteFieldAt(worldShipFields.Ref2.x, worldShipFields.Ref2.y, data.SFName)
    Sleep 120
    PasteFieldAt(worldShipFields.SFPhone.x, worldShipFields.SFPhone.y, data.SFPhone)
    Sleep 120

    MouseClick, left, % worldShipTabs.ShipTo.x, worldShipTabs.ShipTo.y
    Sleep 120
    MouseClick, left, % worldShipTabs.Service.x, worldShipTabs.Service.y
    Sleep 120

    companyName := data.Company != "" ? data.Company : data.STName
    PasteFieldAt(worldShipFields.Company.x, worldShipFields.Company.y, companyName)
    Sleep 120
    MouseClick, left, % worldShipFields.Ref2.x, worldShipFields.Ref2.y
    delay := (companyName = data.Company && companyName != "") ? 2000 : 5000
    Sleep, %delay%
    PasteFieldAt(worldShipFields.STName.x, worldShipFields.STName.y, data.STName)
    Sleep 120

    PasteFieldAt(worldShipFields.Address1.x, worldShipFields.Address1.y, data.Address1)
    Sleep 120
    PasteFieldAt(worldShipFields.Address2.x, worldShipFields.Address2.y, data.Address2)
    Sleep 120
    PasteFieldAt(worldShipFields.STPhone.x, worldShipFields.STPhone.y, data.STPhone)
    Sleep 120

    PastePostalCode(data.PostalCode)
    PasteDeclaredValue(data.DeclaredValue)
    PasteAliasAndQVn(data.Alias)
}

PastePersonalToWorldShip(data)
{
    global worldShipTabs, worldShipFields
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    Sleep 120

    MouseClick, left, % worldShipTabs.Service.x, worldShipTabs.Service.y
    Sleep 120
    MouseClick, left, % worldShipTabs.ShipFrom.x, worldShipTabs.ShipFrom.y
    Sleep 120
    PasteFieldAt(worldShipFields.Company.x, worldShipFields.Company.y, data.SFName)
    Sleep 120
    MouseClick, left, % worldShipFields.Ref2.x, worldShipFields.Ref2.y
    Sleep 5000
    PasteFieldAt(worldShipFields.SFName.x, worldShipFields.SFName.y, data.SFName)
    Sleep 120
    EnsureWorldShipTop()
    PasteFieldAt(worldShipFields.Ref2.x, worldShipFields.Ref2.y, data.SFName)
    Sleep 120
    PasteFieldAt(worldShipFields.SFPhone.x, worldShipFields.SFPhone.y, data.SFPhone)
    Sleep 120

    MouseClick, left, % worldShipTabs.ShipTo.x, worldShipTabs.ShipTo.y
    Sleep 120

    companyName := data.Company != "" ? data.Company : data.STName
    PasteFieldAt(worldShipFields.Company.x, worldShipFields.Company.y, companyName)
    Sleep 120
    MouseClick, left, % worldShipFields.Ref2.x, worldShipFields.Ref2.y
    delay := (companyName = data.Company && companyName != "") ? 2000 : 5000
    Sleep, %delay%
    PasteFieldAt(worldShipFields.STName.x, worldShipFields.STName.y, data.STName)
    Sleep 120

    PasteFieldAt(worldShipFields.Address1.x, worldShipFields.Address1.y, data.Address1)
    Sleep 120
    PasteFieldAt(worldShipFields.Address2.x, worldShipFields.Address2.y, data.Address2)
    Sleep 120
    PasteFieldAt(worldShipFields.STPhone.x, worldShipFields.STPhone.y, data.STPhone)
    Sleep 120

    PastePostalCode(data.PostalCode)
    PasteDeclaredValue(data.DeclaredValue)
    PasteAliasAndQVn(data.Alias)
}

ScrollDownToShipToLegacy()
{
    FocusIntraWindow()
    EnsureIntraWindow()
    Sleep 150
    NeutralClick()
    Sleep 250
    Loop 11
    {
        Sleep 25
        SendInput, {WheelDown}
        Sleep 25
    }
    Sleep 150
    ; Settling time + neutral click before resuming copies.
    NeutralClick()
    Sleep 200
}

CopyDeclaredValueLegacy(fields, offsetY := 0)
{
    ClipSaved := ClipboardAll
    Clipboard :=
    MouseClick, left, % fields.PostalCode.x, fields.PostalCode.y + offsetY
    Sleep 100
    Loop 6
    {
        Sleep 50
        SendInput, {Tab}
        Sleep 50
    }
    Sleep 120
    value := CopyCaretValue()
    if (!RegExMatch(value, "^\d{0,5}$"))
    {
        SendInput, {Tab}
        Sleep 120
        value := CopyCaretValue()
    }
    Clipboard := ClipSaved
    ClipSaved := ""
    return value
}

PastePostalCode(postalCode)
{
    global worldShipFields, worldShipTabs
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
    MouseClick, left, % worldShipFields.Ref2.x, worldShipFields.Ref2.y
    Sleep 2000  ; allow any address book/city-state prompts to settle
    Clipboard := ClipSaved
    ClipSaved := ""
}

PasteDeclaredValue(declaredValue)
{
    global worldShipTabs, worldShipFields
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    Sleep 120
    MouseClick, left, % worldShipTabs.ShipTo.x, worldShipTabs.ShipTo.y
    Sleep 120
    MouseClick, left, % worldShipTabs.Service.x, worldShipTabs.Service.y
    Sleep 120
    PasteFieldAt(worldShipFields.DeclVal.x, worldShipFields.DeclVal.y, declaredValue)
    Sleep 120
}

PasteAliasAndQVn(alias)
{
    global worldShipTabs, worldShipFields
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    Sleep 120
    PasteFieldAt(worldShipFields.STEmail.x, worldShipFields.STEmail.y, alias)
    Sleep 120
    Send {End}
    Sleep 120
    MouseClick, left, % worldShipTabs.Options.x, worldShipTabs.Options.y
    Sleep 120
    MouseClick, left, % worldShipTabs.QVN.x, worldShipTabs.QVN.y
    Sleep 120
    MouseClick, left, % worldShipTabs.Recipients.x, worldShipTabs.Recipients.y
    Sleep 250
    WinWaitActive, Quantum View Notify Recipients (Shipment),, 1
    Sleep 250
    Loop 2
    {
        Sleep 25
        Send {Tab}
        Sleep 25
    }
    Sleep 120
    ClipSaved := ClipboardAll
    Clipboard := alias
    SendInput, ^v
    Sleep 120
    Clipboard := ClipSaved
    ClipSaved := ""
    Sleep 120
    Send {End}
    Sleep 120
    Send {Enter}
}

FocusIntraWindow()
{
    title := GetIntraWindowTitle()
    WinActivate, %title%
    WinWaitActive, %title%,, 1
}

EnsureIntraWindow()
{
    title := GetIntraWindowTitle()
    ; Match the working dimensions used in Intra_Buttons. Adjust here if the target size changes.
    WinMove, %title%,, 1917, 0, 1530, 1399
    Sleep 150
}

FocusWorldShipWindow()
{
    global worldShipTitle
    WinActivate, %worldShipTitle%
    WinWaitActive, %worldShipTitle%,, 1
    DisableWorldShipScale()
}

EnsureWorldShipTop()
{
    ; Neutral click + slight scroll up to keep WorldShip fields in view
    MouseClick, left, 430, 335
    Sleep 150
    SendInput, {WheelUp 5}
    Sleep 200
}

NeutralClick()
{
    global neutralClickR
    WinGetPos,,, winW, winH, A
    targetX := Floor(winW * neutralClickR.x)
    targetY := Floor(winH * neutralClickR.y)
    MouseClick, left, %targetX%, %targetY%
    Sleep 200
}

NeutralAndHome()
{
    global neutralClickR
    WinGetPos,,, winW, winH, A
    targetX := Floor(winW * neutralClickR.x)
    targetY := Floor(winH * neutralClickR.y)
    MouseClick, left, %targetX%, %targetY%
    Sleep 200
    SendInput, ^{Home}
    Sleep 200
}

CopyFieldAt(x, y)
{
    local ClipSaved, text
    ClipSaved := ClipboardAll
    Clipboard :=
    MouseClick, left, %x%, %y%
    Sleep 150
    SendInput, ^a
    Sleep 80
    SendInput, ^c
    ClipWait, 0.5
    text := Clipboard
    Clipboard := ClipSaved
    ClipSaved := ""
    return text
}

CopyCaretValue()
{
    local ClipSaved, text
    ClipSaved := ClipboardAll
    Clipboard :=
    SendInput, ^a
    Sleep 80
    SendInput, ^c
    ClipWait, 0.5
    text := Clipboard
    Clipboard := ClipSaved
    ClipSaved := ""
    return text
}

PasteFieldAt(x, y, text)
{
    local ClipSaved
    if (text = "")
        return
    FocusWorldShipWindow()
    ClipSaved := ClipboardAll
    Clipboard := text
    MouseClick, left, %x%, %y%
    Sleep 150
    ; WorldShip ignores Ctrl+A, so clear via Home -> Shift+End -> Delete
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

EnsureDeclaredValueFocus(targetX := "", targetY := "", showTooltip := true)
{
    global intraFields
    if (targetX = "")
        targetX := intraFields.DeclaredValue.x
    if (targetY = "")
        targetY := intraFields.DeclaredValue.y
    if (TryDeclaredValueCaretHit(targetX, targetY))
        return
    SendInput, {Tab}
    Sleep 120
    if (TryDeclaredValueCaretHit(targetX, targetY))
        return
    if (showTooltip)
    {
        ToolTip, Declared Value field focus failed; please click it manually.
        SetTimer, HideDVTooltip, -5000
    }
}

TryDeclaredValueCaretHit(targetX, targetY)
{
    caretX := A_CaretX
    caretY := A_CaretY
    if (caretX = "" || caretY = "")
        return false
    dx := Abs(caretX - targetX)
    dy := Abs(caretY - targetY)
    return (dx <= 60 && dy <= 60)
}

HideDVTooltip:
    ToolTip
return

ShowHotkeyRuntime(startTick)
{
    elapsedMs := A_TickCount - startTick
    elapsedSec := Round(elapsedMs / 1000.0, 2)
    ToolTip, Hotkey Runtime: %elapsedSec% seconds
    SetTimer, HideRuntimeTooltip, -4000
}

HideRuntimeTooltip:
    ToolTip
return

GetIntraWindowTitle()
{
    global intraWinTitle, intraWinExes
    Loop % intraWinExes.Length()
    {
        candidate := intraWinTitle " ahk_exe " intraWinExes[A_Index]
        if (WinExist(candidate))
            return candidate
    }
    return intraWinTitle
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
