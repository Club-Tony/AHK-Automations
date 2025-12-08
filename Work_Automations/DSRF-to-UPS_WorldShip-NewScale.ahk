#Requires AutoHotkey v1
#NoEnv
#Warn
#SingleInstance, Force
SendMode Input
SetWorkingDir %A_ScriptDir%

SetTitleMatchMode, 2
SetDefaultMouseSpeed, 0
CoordMode, Mouse, Window

; Window targets
intraWinTitle := "Intra: Shipping Request Form ahk_exe firefox.exe"
worldShipTitle := "UPS WorldShip"

; NOTE: Re-record only the Intra coordinates below at 50% zoom. Current values are
; placeholders copied from the 100% version. WorldShip coordinates remain the same.

intraFields := {}
intraFields.CostCenter    := {x: 410, y: 581}
intraFields.Alias         := {x: 410, y: 788}
intraFields.SFName        := {x: 800, y: 880}
intraFields.SFPhone       := {x: 1040, y: 788}
intraFields.STName        := {x: 400, y: 1361}
intraFields.Company       := {x: 800, y: 1360}
intraFields.Address1      := {x: 410, y: 380}
intraFields.Address2      := {x: 410, y: 470}
intraFields.STPhone       := {x: 800, y: 381}
intraFields.PostalCode    := {x: 1040, y: 470}
intraFields.DeclaredValue := {x: 410, y: 830}

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

return  ; end auto-execute

Esc::ExitApp

^!b:: ; Business Form (50% zoom scaffold)
    FocusIntraWindow()
    EnsureIntraWindowZoom50()
    Sleep 50
    NeutralAndHome()
    Sleep 50
    costCenter := CopyFieldAt(intraFields.CostCenter.x, intraFields.CostCenter.y)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    MouseClick, left, % worldShipTabs.ShipTo.x, worldShipTabs.ShipTo.y
    Sleep 150
    PasteFieldAt(worldShipFields.Ref1.x, worldShipFields.Ref1.y, costCenter)

    FocusIntraWindow()
    EnsureIntraWindowZoom50()
    Sleep 50
    NeutralAndHome()
    sfName := CopyFieldAt(intraFields.SFName.x, intraFields.SFName.y)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    MouseClick, left, % worldShipTabs.Service.x, worldShipTabs.Service.y
    Sleep 150
    MouseClick, left, % worldShipTabs.ShipFrom.x, worldShipTabs.ShipFrom.y
    Sleep 150
    PasteFieldAt(worldShipFields.Company.x, worldShipFields.Company.y, sfName)
    Sleep 150
    MouseClick, left, % worldShipFields.Ref2.x, worldShipFields.Ref2.y
    Sleep 2000
    PasteFieldAt(worldShipFields.SFName.x, worldShipFields.SFName.y, sfName)
    Sleep 150
    EnsureWorldShipTop()
    Sleep 150
    PasteFieldAt(worldShipFields.Ref2.x, worldShipFields.Ref2.y, sfName)

    FocusIntraWindow()
    EnsureIntraWindowZoom50()
    Sleep 50
    NeutralAndHome()
    sfPhone := CopyFieldAt(intraFields.SFPhone.x, intraFields.SFPhone.y)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    Sleep 150
    PasteFieldAt(worldShipFields.SFPhone.x, worldShipFields.SFPhone.y, sfPhone)
    Sleep 150
    FocusWorldShipWindow()
    MouseClick, left, % worldShipTabs.ShipTo.x, worldShipTabs.ShipTo.y

    FocusIntraWindow()
    EnsureIntraWindowZoom50()
    Sleep 50
    NeutralAndHome()
    company := CopyFieldAt(intraFields.Company.x, intraFields.Company.y)

    FocusIntraWindow()
    EnsureIntraWindowZoom50()
    Sleep 50
    NeutralAndHome()
    stName := CopyFieldAt(intraFields.STName.x, intraFields.STName.y)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    Sleep 150
    if (company != "")
    {
        PasteFieldAt(worldShipFields.Company.x, worldShipFields.Company.y, company)
        Sleep 150
        MouseClick, left, % worldShipFields.Ref2.x, worldShipFields.Ref2.y
        Sleep 2000
    }
    else
    {
        PasteFieldAt(worldShipFields.Company.x, worldShipFields.Company.y, stName)
        Sleep 150
        MouseClick, left, % worldShipFields.Ref2.x, worldShipFields.Ref2.y
        Sleep 2000
    }
    PasteFieldAt(worldShipFields.STName.x, worldShipFields.STName.y, stName)

    ; Address block (no scrolling needed at 50% zoom)
    FocusIntraWindow()
    EnsureIntraWindowZoom50()
    Sleep 50
    Address1 := CopyFieldAt(intraFields.Address1.x, intraFields.Address1.y)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    MouseClick, left, % worldShipTabs.Address1.x, worldShipTabs.Address1.y
    Sleep 150
    PasteFieldAt(worldShipFields.Address1.x, worldShipFields.Address1.y, Address1)

    FocusIntraWindow()
    EnsureIntraWindowZoom50()
    Sleep 50
    Address2 := CopyFieldAt(intraFields.Address2.x, intraFields.Address2.y)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    MouseClick, left, % worldShipTabs.Address2.x, worldShipTabs.Address2.y
    Sleep 150
    PasteFieldAt(worldShipFields.Address2.x, worldShipFields.Address2.y, Address2)

    FocusIntraWindow()
    EnsureIntraWindowZoom50()
    Sleep 50
    STPhone := CopyFieldAt(intraFields.STPhone.x, intraFields.STPhone.y)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    MouseClick, left, % worldShipTabs.STPhone.x, worldShipTabs.STPhone.y
    Sleep 150
    PasteFieldAt(worldShipFields.STPhone.x, worldShipFields.STPhone.y, STPhone)

    ; Postal code: clear field first to handle autofill, then paste
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    MouseClick, left, % worldShipFields.PostalCode.x, worldShipFields.PostalCode.y
    Sleep 150
    SendInput, {Home}
    Sleep 80
    SendInput, +{End}
    Sleep 80
    SendInput, {Delete}
    MouseClick, left, % worldShipFields.Ref2.x, worldShipFields.Ref2.y
    Sleep 2000
    SendInput, {Enter}

    FocusIntraWindow()
    EnsureIntraWindowZoom50()
    Sleep 50
    PostalCode := CopyFieldAt(intraFields.PostalCode.x, intraFields.PostalCode.y)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    Sleep 250
    PasteFieldAt(worldShipFields.PostalCode.x, worldShipFields.PostalCode.y, PostalCode)
    Sleep 250
    MouseClick, left, % worldShipFields.Ref2.x, worldShipFields.Ref2.y
    Sleep 2000

    ; Declared value
    FocusIntraWindow()
    EnsureIntraWindowZoom50()
    Sleep 50
    ClipSaved := ClipboardAll
    Clipboard :=
    MouseClick, left, % intraFields.DeclaredValue.x, intraFields.DeclaredValue.y
    Sleep 150
    SendInput, {End}
    Sleep 100
    SendInput, ^+{Left}
    Sleep 100
    SendInput, ^c
    ClipWait, 0.5
    DeclaredValue := Clipboard
    Clipboard := ClipSaved
    ClipSaved := ""
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    Sleep 150
    PasteFieldAt(worldShipFields.DeclVal.x, worldShipFields.DeclVal.y, DeclaredValue)

    ; Alias -> email/QVN
    FocusIntraWindow()
    EnsureIntraWindowZoom50()
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
    Sleep 100
    Send @amazon.com
    Sleep 150
    MouseClick, left, % worldShipTabs.Options.x, worldShipTabs.Options.y
    Sleep 150
    MouseClick, left, % worldShipTabs.QVN.x, worldShipTabs.QVN.y
    Sleep 150
    MouseClick, left, % worldShipTabs.Recipients.x, worldShipTabs.Recipients.y
    Sleep 150
    EnsureWorldShipTop()
    Sleep 150
    PasteFieldAt(worldShipTabs.QVNEmail.x, worldShipTabs.QVNEmail.y, Alias)
    Sleep 150
    Send {End}
    Sleep 150
    Send @amazon.com
    Sleep 150
    Send {Enter}
return

^!p:: ; Personal Form (50% zoom scaffold)
    offsetY := -85

    FocusIntraWindow()
    EnsureIntraWindowZoom50()
    Sleep 50
    NeutralAndHome()
    Sleep 50
    sfName := CopyFieldAt(intraFields.SFName.x, intraFields.SFName.y + offsetY)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    Sleep 50
    MouseClick, left, % worldShipTabs.Service.x, worldShipTabs.Service.y
    Sleep 150
    MouseClick, left, % worldShipTabs.ShipFrom.x, worldShipTabs.ShipFrom.y
    Sleep 150
    PasteFieldAt(worldShipFields.Company.x, worldShipFields.Company.y, sfName)
    Sleep 150
    MouseClick, left, % worldShipFields.Ref2.x, worldShipFields.Ref2.y
    Sleep 2000
    PasteFieldAt(worldShipFields.SFName.x, worldShipFields.SFName.y, sfName)
    Sleep 150
    EnsureWorldShipTop()
    Sleep 150
    PasteFieldAt(worldShipFields.Ref2.x, worldShipFields.Ref2.y, sfName)

    FocusIntraWindow()
    EnsureIntraWindowZoom50()
    Sleep 50
    NeutralAndHome()
    sfPhone := CopyFieldAt(intraFields.SFPhone.x, intraFields.SFPhone.y + offsetY)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    Sleep 150
    PasteFieldAt(worldShipFields.SFPhone.x, worldShipFields.SFPhone.y, sfPhone)
    Sleep 150
    FocusWorldShipWindow()
    MouseClick, left, % worldShipTabs.ShipTo.x, worldShipTabs.ShipTo.y

    FocusIntraWindow()
    EnsureIntraWindowZoom50()
    Sleep 50
    NeutralAndHome()
    company := CopyFieldAt(intraFields.Company.x, intraFields.Company.y + offsetY)

    FocusIntraWindow()
    EnsureIntraWindowZoom50()
    Sleep 50
    NeutralAndHome()
    stName := CopyFieldAt(intraFields.STName.x, intraFields.STName.y + offsetY)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    Sleep 150
    if (company != "")
    {
        PasteFieldAt(worldShipFields.Company.x, worldShipFields.Company.y, company)
        Sleep 150
        MouseClick, left, % worldShipFields.Ref2.x, worldShipFields.Ref2.y
        Sleep 2000
    }
    else
    {
        PasteFieldAt(worldShipFields.Company.x, worldShipFields.Company.y, stName)
        Sleep 150
        MouseClick, left, % worldShipFields.Ref2.x, worldShipFields.Ref2.y
        Sleep 2000
    }
    PasteFieldAt(worldShipFields.STName.x, worldShipFields.STName.y, stName)

    ; Address block (no scrolling needed at 50% zoom)
    FocusIntraWindow()
    EnsureIntraWindowZoom50()
    Sleep 50
    Address1 := CopyFieldAt(intraFields.Address1.x, intraFields.Address1.y + offsetY)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    MouseClick, left, % worldShipTabs.Address1.x, worldShipTabs.Address1.y
    Sleep 150
    PasteFieldAt(worldShipFields.Address1.x, worldShipFields.Address1.y, Address1)

    FocusIntraWindow()
    EnsureIntraWindowZoom50()
    Sleep 50
    Address2 := CopyFieldAt(intraFields.Address2.x, intraFields.Address2.y + offsetY)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    MouseClick, left, % worldShipTabs.Address2.x, worldShipTabs.Address2.y
    Sleep 150
    PasteFieldAt(worldShipFields.Address2.x, worldShipFields.Address2.y, Address2)

    FocusIntraWindow()
    EnsureIntraWindowZoom50()
    Sleep 50
    STPhone := CopyFieldAt(intraFields.STPhone.x, intraFields.STPhone.y + offsetY)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    MouseClick, left, % worldShipTabs.STPhone.x, worldShipTabs.STPhone.y
    Sleep 150
    PasteFieldAt(worldShipFields.STPhone.x, worldShipFields.STPhone.y, STPhone)

    ; Postal code: clear field first to handle autofill, then paste
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    MouseClick, left, % worldShipFields.PostalCode.x, worldShipFields.PostalCode.y
    Sleep 150
    SendInput, {Home}
    Sleep 80
    SendInput, +{End}
    Sleep 80
    SendInput, {Delete}
    MouseClick, left, % worldShipFields.Ref2.x, worldShipFields.Ref2.y
    Sleep 2000
    SendInput, {Enter}

    FocusIntraWindow()
    EnsureIntraWindowZoom50()
    Sleep 50
    PostalCode := CopyFieldAt(intraFields.PostalCode.x, intraFields.PostalCode.y + offsetY)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    Sleep 250
    PasteFieldAt(worldShipFields.PostalCode.x, worldShipFields.PostalCode.y, PostalCode)
    Sleep 250
    MouseClick, left, % worldShipFields.Ref2.x, worldShipFields.Ref2.y
    Sleep 2000

    ; Declared value
    FocusIntraWindow()
    EnsureIntraWindowZoom50()
    Sleep 50
    ClipSaved := ClipboardAll
    Clipboard :=
    MouseClick, left, % intraFields.DeclaredValue.x, intraFields.DeclaredValue.y + offsetY
    Sleep 150
    SendInput, {End}
    Sleep 100
    SendInput, ^+{Left}
    Sleep 100
    SendInput, ^c
    ClipWait, 0.5
    DeclaredValue := Clipboard
    Clipboard := ClipSaved
    ClipSaved := ""
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    Sleep 150
    PasteFieldAt(worldShipFields.DeclVal.x, worldShipFields.DeclVal.y, DeclaredValue)

    ; Alias -> email/QVN
    FocusIntraWindow()
    EnsureIntraWindowZoom50()
    Sleep 50
    NeutralAndHome()
    Sleep 150
    Alias := CopyFieldAt(intraFields.Alias.x, intraFields.Alias.y + offsetY)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    Sleep 150
    PasteFieldAt(worldShipFields.STEmail.x, worldShipFields.STEmail.y, Alias)
    Sleep 150
    Send {End}
    Sleep 100
    Send @amazon.com
    Sleep 150
    MouseClick, left, % worldShipTabs.Options.x, worldShipTabs.Options.y
    Sleep 150
    MouseClick, left, % worldShipTabs.QVN.x, worldShipTabs.QVN.y
    Sleep 150
    MouseClick, left, % worldShipTabs.Recipients.x, worldShipTabs.Recipients.y
    Sleep 150
    EnsureWorldShipTop()
    Sleep 150
    PasteFieldAt(worldShipTabs.QVNEmail.x, worldShipTabs.QVNEmail.y, Alias)
    Sleep 150
    Send {End}
    Sleep 150
    Send @amazon.com
    Sleep 150
    Send {Enter}
return

FocusIntraWindow()
{
    global intraWinTitle
    WinActivate, %intraWinTitle%
    WinWaitActive, %intraWinTitle%,, 1
}

EnsureIntraWindowZoom50()
{
    global intraWinTitle
    WinMove, %intraWinTitle%,, 1917, 0, 1530, 1399
    ; TODO: set browser zoom to 50% (implement per your browser/shortcut).
    Sleep 150
}

FocusWorldShipWindow()
{
    global worldShipTitle
    WinActivate, %worldShipTitle%
    WinWaitActive, %worldShipTitle%,, 1
}

EnsureWorldShipTop()
{
    ; Keep WorldShip fields in view; adjust if needed for 50% scaling.
    MouseClick, left, 430, 335
    Sleep 150
    SendInput, {WheelUp 5}
    Sleep 200
}

NeutralAndHome()
{
    ; Click neutral area and Home to ensure top of Intra form (adjust target if needed).
    MouseClick, left, 1718, 708
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

PasteFieldAt(x, y, text)
{
    local ClipSaved
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
