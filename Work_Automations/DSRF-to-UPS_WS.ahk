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
CoordMode, Caret, Window

; Window targets
intraWinTitle := "Intra: Shipping Request Form"
intraWinExes := ["firefox.exe", "chrome.exe", "msedge.exe"]
worldShipTitle := "UPS WorldShip"

; Normalized window size used for ratio/absolute clicks (matches Intra_Buttons)
intraWinW := 1530
intraWinH := 1399
neutralClickR := {x: 0.895425, y: 0.50683}

; Intra: Shipping Request Form field coordinates (window-relative pixels at normalized size)
intraFields := {}
intraFields.CostCenter    := {x: 553, y: 394}
intraFields.Alias         := {x: 551, y: 578}
intraFields.SFName        := {x: 754, y: 572}
intraFields.SFPhone       := {x: 934, y: 517}
intraFields.STName        := {x: 587, y: 863}
intraFields.Company       := {x: 887, y: 863}
intraFields.Address1      := {x: 587, y: 1031}
intraFields.Address2      := {x: 543, y: 1084}
intraFields.STPhone       := {x: 878, y: 1031}
intraFields.PostalCode    := {x: 969, y: 1085}

personalFields := {}
personalFields.Alias         := {x: 552, y: 524}
personalFields.SFName        := {x: 749, y: 518}
personalFields.SFPhone       := {x: 920, y: 464}
personalFields.STName        := {x: 598, y: 809}
personalFields.Company       := {x: 917, y: 809}
personalFields.Address1      := {x: 589, y: 976}
personalFields.Address2      := {x: 548, y: 1030}
personalFields.STPhone       := {x: 891, y: 976}
personalFields.PostalCode    := {x: 1000, y: 1031}

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
PersonalButtonX := 469
PersonalButtonY := 338
BusinessButtonX := 523
BusinessButtonY := 338

return  ; end of auto-execute section

Esc::ExitApp

^!b:: ; Business Form: capture everything in Intra, then paste in WorldShip
    scaleClickDone := false
    startTick := A_TickCount
    data := CollectIntraData("business")
    PasteBusinessToWorldShip(data)
    FocusIntraWindow()
    EnsureIntraWindow()
    EnsureIntraZoom100()
    ShowHotkeyRuntime(startTick)
return

^!p:: ; Personal Form: capture everything in Intra, then paste in WorldShip
    scaleClickDone := false
    startTick := A_TickCount
    data := CollectIntraData("personal")
    PastePersonalToWorldShip(data)
    FocusIntraWindow()
    EnsureIntraWindow()
    EnsureIntraZoom100()
    ShowHotkeyRuntime(startTick)
return

CollectIntraData(mode)
{
    global intraFields, personalFields, BusinessButtonX, BusinessButtonY, PersonalButtonX, PersonalButtonY
    FocusIntraWindow()
    EnsureIntraWindow()
    EnsureIntraZoom60()
    Sleep 250
    NeutralAndHome()
    Sleep 80

    fields := (mode = "business") ? intraFields : personalFields
    includeCostCenter := (mode = "business")

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
        payload.CostCenter := CopyFieldAt(fields.CostCenter.x, fields.CostCenter.y)
    payload.Alias      := CopyFieldAt(fields.Alias.x, fields.Alias.y)
    payload.SFName     := CopyFieldAt(fields.SFName.x, fields.SFName.y)
    payload.SFPhone    := CopyFieldAt(fields.SFPhone.x, fields.SFPhone.y)
    payload.Company    := CopyFieldAt(fields.Company.x, fields.Company.y)
    payload.STName     := CopyFieldAt(fields.STName.x, fields.STName.y)
    payload.Address1   := CopyFieldAt(fields.Address1.x, fields.Address1.y)
    payload.Address2   := CopyFieldAt(fields.Address2.x, fields.Address2.y)
    payload.STPhone    := CopyFieldAt(fields.STPhone.x, fields.STPhone.y)
    payload.PostalCode := CopyFieldAt(fields.PostalCode.x, fields.PostalCode.y)
    payload.DeclaredValue := CopyDeclaredValue(fields)
    ; Jump to WorldShip immediately after collecting all fields to avoid any extra scrolling in Intra
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

CopyDeclaredValue(fields)
{
    ClipSaved := ClipboardAll
    Clipboard :=
    ; Start from Postal Code then tab into Declared Value (coordinate-light like the other scripts)
    MouseClick, left, % fields.PostalCode.x, fields.PostalCode.y
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
    Sleep 2000
    Clipboard := ClipSaved
    ClipSaved := ""
}

PasteDeclaredValue(declaredValue)
{
    global worldShipTabs, worldShipFields
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
    PasteFieldAt(worldShipFields.STEmail.x, worldShipFields.STEmail.y, alias)
    Sleep 150
    Send {End}
    Sleep 150
    MouseClick, left, % worldShipTabs.Options.x, worldShipTabs.Options.y
    Sleep 150
    MouseClick, left, % worldShipTabs.QVN.x, worldShipTabs.QVN.y
    Sleep 150
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
    Sleep 150
    ClipSaved := ClipboardAll
    Clipboard := alias
    SendInput, ^v
    Sleep 150
    Clipboard := ClipSaved
    ClipSaved := ""
    Sleep 150
    Send {End}
    Sleep 150
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
    WinMove, %title%,, 1917, 0, 1530, 1399
    Sleep 150
}

EnsureIntraZoom60()
{
    SendInput, ^0
    Sleep 150
    Loop 4
    {
        SendInput, ^{WheelDown}
        Sleep 120
    }
}

EnsureIntraZoom100()
{
    SendInput, ^0
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
    MouseClick, left, 430, 335
    Sleep 150
    SendInput, {WheelUp 5}
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
    if (text = "")
        return
    FocusWorldShipWindow()
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
