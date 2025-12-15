#Requires AutoHotkey v1
#NoEnv ; Prevents Unnecessary Environment Variable lookup
#Warn ; Warn All (All Warnings Enabled)
#SingleInstance, Force  ; Reload without prompt when Esc is pressed.
SendMode Input ; Send works as SendInput
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
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
intraFields.DeclaredValue := {x: 540, y: 1273}

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
personalFields.DeclaredValue := {x: 544, y: 1246}

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

^!b:: ; Business Form (mirrors ^!p without offsets/cost center)
    scaleClickDone := false
    startTick := A_TickCount
    FocusIntraWindow()
    EnsureIntraWindow()
    EnsureIntraZoom60()
    Sleep 250
    NeutralAndHome()
    Sleep 50
    SendInput, {WheelUp 15}
    Sleep 150
    MouseClick, left, %BusinessButtonX%, %BusinessButtonY%, 2
    Sleep 150
    FocusIntraWindow()
    costCenter := CopyFieldAt(intraFields.CostCenter.x, intraFields.CostCenter.y)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    Sleep 150
    MouseClick, left, % worldShipTabs.ShipTo.x, worldShipTabs.ShipTo.y
    Sleep 150
    MouseClick, left, % worldShipTabs.Service.x, worldShipTabs.Service.y
    Sleep 150
    PasteFieldAt(worldShipFields.Ref1.x, worldShipFields.Ref1.y, costCenter)
    Sleep 150
    FocusIntraWindow()
    sfName := CopyFieldAt(intraFields.SFName.x, intraFields.SFName.y)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    Sleep 150
    MouseClick, left, % worldShipTabs.Service.x, worldShipTabs.Service.y  ; ensure Service tab active before first paste
    Sleep 150
    MouseClick, left, % worldShipTabs.ShipFrom.x, worldShipTabs.ShipFrom.y
    Sleep 150
    FocusWorldShipWindow()
    PasteFieldAt(worldShipFields.Company.x, worldShipFields.Company.y, sfName)
    Sleep 150
    MouseClick, left, % worldShipFields.Ref2.x, worldShipFields.Ref2.y
    Sleep 3000
    FocusWorldShipWindow()
    PasteFieldAt(worldShipFields.SFName.x, worldShipFields.SFName.y, sfName)
    Sleep 150
    EnsureWorldShipTop()
    Sleep 150
    FocusWorldShipWindow()
    PasteFieldAt(worldShipFields.Ref2.x, worldShipFields.Ref2.y, sfName)

    FocusIntraWindow()
    EnsureIntraWindow()
    Sleep 50
    NeutralAndHome()
    FocusIntraWindow()
    sfPhone := CopyFieldAt(intraFields.SFPhone.x, intraFields.SFPhone.y)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    Sleep 150
    FocusWorldShipWindow()
    PasteFieldAt(worldShipFields.SFPhone.x, worldShipFields.SFPhone.y, sfPhone)
    Sleep 150
    FocusWorldShipWindow()
    MouseClick, left, % worldShipTabs.ShipTo.x, worldShipTabs.ShipTo.y
    Sleep 150
    MouseClick, left, % worldShipTabs.Service.x, worldShipTabs.Service.y

    FocusIntraWindow()
    EnsureIntraWindow()
    Sleep 50
    NeutralAndHome()
    FocusIntraWindow()
    company := CopyFieldAt(intraFields.Company.x, intraFields.Company.y)

    stName := CopyFieldAt(intraFields.STName.x, intraFields.STName.y)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    Sleep 150
    if (company != "")
    {
        FocusWorldShipWindow()
        PasteFieldAt(worldShipFields.Company.x, worldShipFields.Company.y, company)
        Sleep 150
        MouseClick, left, % worldShipFields.Ref2.x, worldShipFields.Ref2.y
        Sleep 3000
    }
    else
    {
        FocusWorldShipWindow()
        PasteFieldAt(worldShipFields.Company.x, worldShipFields.Company.y, stName)
        Sleep 150
        MouseClick, left, % worldShipFields.Ref2.x, worldShipFields.Ref2.y
        Sleep 3000  ; allow WorldShip address book fill to process
    }
    FocusWorldShipWindow()
    PasteFieldAt(worldShipFields.STName.x, worldShipFields.STName.y, stName)

    FocusIntraWindow()
    Address1 := CopyFieldAt(intraFields.Address1.x, intraFields.Address1.y)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    MouseClick, left, % worldShipFields.Address1.x, worldShipFields.Address1.y
    Sleep 150
    FocusWorldShipWindow()
    PasteFieldAt(worldShipFields.Address1.x, worldShipFields.Address1.y, Address1)

    FocusIntraWindow()
    EnsureIntraWindow()
    Sleep 50
    FocusIntraWindow()
    Address2 := CopyFieldAt(intraFields.Address2.x, intraFields.Address2.y)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    MouseClick, left, % worldShipFields.Address2.x, worldShipFields.Address2.y
    Sleep 150
    FocusWorldShipWindow()
    PasteFieldAt(worldShipFields.Address2.x, worldShipFields.Address2.y, Address2)

    FocusIntraWindow()
    EnsureIntraWindow()
    Sleep 50
    FocusIntraWindow()
    STPhone := CopyFieldAt(intraFields.STPhone.x, intraFields.STPhone.y)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    MouseClick, left, % worldShipFields.STPhone.x, worldShipFields.STPhone.y
    Sleep 150
    FocusWorldShipWindow()
    PasteFieldAt(worldShipFields.STPhone.x, worldShipFields.STPhone.y, STPhone)

    ; Clear PostalCode first in WorldShip (handles saved address autofill), then copy/paste
    FocusIntraWindow()
    EnsureIntraWindow()
    Sleep 50
    FocusIntraWindow()
    PostalCode := CopyFieldAt(intraFields.PostalCode.x, intraFields.PostalCode.y)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    ClipSaved := ClipboardAll
    Clipboard := PostalCode
    MouseClick, left, % worldShipFields.PostalCode.x, worldShipFields.PostalCode.y
    Sleep 150
    SendInput, {Home}
    Sleep 80
    SendInput, +{End}
    Sleep 80
    SendInput, {Delete}
    Sleep 250
    SendInput, %PostalCode%
    Sleep 250
    MouseClick, left, % worldShipFields.Ref2.x, worldShipFields.Ref2.y
    Sleep 2000  ; allow any address book/city-state prompts to settle

    FocusIntraWindow()
    EnsureIntraWindow()
    Sleep 50
    FocusIntraWindow()
    MouseClick, left, % intraFields.PostalCode.x, intraFields.PostalCode.y
    Sleep 150
    Loop 6
    {
        Sleep 50
        SendInput, {Tab}
        Sleep 50
    }
    Sleep 150
    DeclaredValue := CopyCaretValue()
    if (!RegExMatch(DeclaredValue, "^\d{0,5}$"))
    {
        SendInput, {Tab}
        Sleep 120
        DeclaredValue := CopyCaretValue()
    }
    Clipboard := ClipSaved
    ClipSaved := ""
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    Sleep 150
    MouseClick, left, % worldShipTabs.ShipTo.x, worldShipTabs.ShipTo.y  ; keep Ship To tab active
    Sleep 150
    MouseClick, left, % worldShipTabs.Service.x, worldShipTabs.Service.y
    Sleep 150
    PasteFieldAt(worldShipFields.DeclVal.x, worldShipFields.DeclVal.y, DeclaredValue)
    Sleep 150

    ; alias paste in email field then select options-qvn-recipients-
    ; paste into qvnemail-then done, optionally implement a send enter
    FocusIntraWindow()
    EnsureIntraWindow()
    Sleep 50
    NeutralAndHome()
    Sleep 150
    Alias := CopyFieldAt(intraFields.Alias.x, intraFields.Alias.y)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    Sleep 150
    PasteFieldAt(worldShipFields.STEmail.x, worldShipFields.STEmail.y, Alias)
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
    Sleep 150
    ClipSaved := ClipboardAll
    Clipboard := Alias
    SendInput, ^v
    Sleep 150
    Clipboard := ClipSaved
    ClipSaved := ""
    Sleep 150
    Send {End}
    Sleep 150
    Send {Enter}
    ; Return to Intra and restore default zoom
    FocusIntraWindow()
    EnsureIntraWindow()
    EnsureIntraZoom100()
    ShowHotkeyRuntime(startTick)
return

^!p:: ; Personal Form
    scaleClickDone := false
    startTick := A_TickCount
    FocusIntraWindow()
    EnsureIntraWindow()
    EnsureIntraZoom60()
    Sleep 250
    NeutralAndHome()
    Sleep 50
    MouseClick, left, %PersonalButtonX%, %PersonalButtonY%, 2
    Sleep 150
    FocusIntraWindow()
    sfName := CopyFieldAt(personalFields.SFName.x, personalFields.SFName.y)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    Sleep 150
    MouseClick, left, % worldShipTabs.Service.x, worldShipTabs.Service.y  ; ensure Service tab active before first paste
    Sleep 150
    MouseClick, left, % worldShipTabs.ShipFrom.x, worldShipTabs.ShipFrom.y
    Sleep 150
    FocusWorldShipWindow()
    PasteFieldAt(worldShipFields.Company.x, worldShipFields.Company.y, sfName)
    Sleep 150
    MouseClick, left, % worldShipFields.Ref2.x, worldShipFields.Ref2.y
    Sleep 3000
    FocusWorldShipWindow()
    PasteFieldAt(worldShipFields.SFName.x, worldShipFields.SFName.y, sfName)
    Sleep 150
    EnsureWorldShipTop()
    Sleep 150
    FocusWorldShipWindow()
    PasteFieldAt(worldShipFields.Ref2.x, worldShipFields.Ref2.y, sfName)

    FocusIntraWindow()
    EnsureIntraWindow()
    Sleep 50
    NeutralAndHome()
    FocusIntraWindow()
    sfPhone := CopyFieldAt(personalFields.SFPhone.x, personalFields.SFPhone.y)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    Sleep 150
    FocusWorldShipWindow()
    PasteFieldAt(worldShipFields.SFPhone.x, worldShipFields.SFPhone.y, sfPhone)
    Sleep 150
    FocusWorldShipWindow()
    MouseClick, left, % worldShipTabs.ShipTo.x, worldShipTabs.ShipTo.y

    FocusIntraWindow()
    EnsureIntraWindow()
    Sleep 50
    NeutralAndHome()
    FocusIntraWindow()
    company := CopyFieldAt(personalFields.Company.x, personalFields.Company.y)

    FocusIntraWindow()
    EnsureIntraWindow()
    Sleep 50
    NeutralAndHome()
    FocusIntraWindow()
    stName := CopyFieldAt(personalFields.STName.x, personalFields.STName.y)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    Sleep 150
    if (company != "")
    {
        FocusWorldShipWindow()
        PasteFieldAt(worldShipFields.Company.x, worldShipFields.Company.y, company)
        Sleep 150
        MouseClick, left, % worldShipFields.Ref2.x, worldShipFields.Ref2.y
        Sleep 2000
    }
    else
    {
        FocusWorldShipWindow()
        PasteFieldAt(worldShipFields.Company.x, worldShipFields.Company.y, stName)
        Sleep 150
        MouseClick, left, % worldShipFields.Ref2.x, worldShipFields.Ref2.y
        Sleep 3000  ; allow WorldShip address book fill to process
    }
    FocusWorldShipWindow()
    PasteFieldAt(worldShipFields.STName.x, worldShipFields.STName.y, stName)

    FocusIntraWindow()
    Address1 := CopyFieldAt(personalFields.Address1.x, personalFields.Address1.y)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    MouseClick, left, % worldShipFields.Address1.x, worldShipFields.Address1.y
    Sleep 150
    FocusWorldShipWindow()
    PasteFieldAt(worldShipFields.Address1.x, worldShipFields.Address1.y, Address1)

    FocusIntraWindow()
    EnsureIntraWindow()
    Sleep 50
    FocusIntraWindow()
    Address2 := CopyFieldAt(personalFields.Address2.x, personalFields.Address2.y)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    MouseClick, left, % worldShipFields.Address2.x, worldShipFields.Address2.y
    Sleep 150
    FocusWorldShipWindow()
    PasteFieldAt(worldShipFields.Address2.x, worldShipFields.Address2.y, Address2)

    FocusIntraWindow()
    EnsureIntraWindow()
    Sleep 50
    FocusIntraWindow()
    STPhone := CopyFieldAt(personalFields.STPhone.x, personalFields.STPhone.y)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    MouseClick, left, % worldShipFields.STPhone.x, worldShipFields.STPhone.y
    Sleep 150
    FocusWorldShipWindow()
    PasteFieldAt(worldShipFields.STPhone.x, worldShipFields.STPhone.y, STPhone)

    ; Clear PostalCode first in WorldShip (handles saved address autofill), then copy/paste
    FocusIntraWindow()
    EnsureIntraWindow()
    Sleep 50
    FocusIntraWindow()
    PostalCode := CopyFieldAt(personalFields.PostalCode.x, personalFields.PostalCode.y)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    ClipSaved := ClipboardAll
    Clipboard := PostalCode
    MouseClick, left, % worldShipFields.PostalCode.x, worldShipFields.PostalCode.y
    Sleep 150
    SendInput, {Home}
    Sleep 80
    SendInput, +{End}
    Sleep 80
    SendInput, {Delete}
    Sleep 250
    SendInput, %PostalCode%
    Sleep 250
    MouseClick, left, % worldShipFields.Ref2.x, worldShipFields.Ref2.y
    Sleep 2000  ; allow any address book/city-state prompts to settle

    FocusIntraWindow()
    EnsureIntraWindow()
    Sleep 50
    FocusIntraWindow()
    MouseClick, left, % personalFields.PostalCode.x, personalFields.PostalCode.y, 2
    Sleep 150
    SendInput, {Right}
    Sleep 100
    Loop 6
    {
        Sleep 50
        SendInput, {Tab}
        Sleep 50
    }
    Sleep 150
    DeclaredValue := CopyCaretValue()
    if (!RegExMatch(DeclaredValue, "^\d{0,5}$"))
    {
        SendInput, {Tab}
        Sleep 120
        DeclaredValue := CopyCaretValue()
    }
    Clipboard := ClipSaved
    ClipSaved := ""
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    Sleep 150
    MouseClick, left, % worldShipTabs.ShipTo.x, worldShipTabs.ShipTo.y  ; keep Ship To tab active
    Sleep 150
    MouseClick, left, % worldShipTabs.Service.x, worldShipTabs.Service.y
    Sleep 150
    PasteFieldAt(worldShipFields.DeclVal.x, worldShipFields.DeclVal.y, DeclaredValue)
    Sleep 150

    ; alias paste in email field then select options-qvn-recipients-
    ; paste into qvnemail-then done, optionally implement a send enter
    FocusIntraWindow()
    EnsureIntraWindow()
    Sleep 50
    NeutralAndHome()
    Sleep 150
    Alias := CopyFieldAt(personalFields.Alias.x, personalFields.Alias.y)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    Sleep 150
    PasteFieldAt(worldShipFields.STEmail.x, worldShipFields.STEmail.y, Alias)
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
    Sleep 150
    ClipSaved := ClipboardAll
    Clipboard := Alias
    SendInput, ^v
    Sleep 150
    Clipboard := ClipSaved
    ClipSaved := ""
    Sleep 150
    Send {End}
    Sleep 150
    Send {Enter}
    ; Return to Intra and restore default zoom
    FocusIntraWindow()
    EnsureIntraWindow()
    EnsureIntraZoom100()
    ShowHotkeyRuntime(startTick)
return

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

EnsureIntraZoom60()
{
    ; Reset zoom, then use Ctrl+WheelDown steps (10% each in Firefox) to land on ~60%.
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
    ; second extra Tab to cover cases where focus lands on the checkbox just before Declared Value
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
