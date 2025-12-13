#Requires AutoHotkey v1
#NoEnv ; Prevents Unnecessary Environment Variable lookup
#Warn ; Warn All (All Warnings Enabled)
#SingleInstance, Force  ; Reload without prompt when Esc is pressed.
SendMode Input ; Send works as SendInput
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetKeyDelay, 25

/*
TODO (capture fresh absolute coords at 50% zoom via AHK Window Spy):
- Intra fields: CostCenter, Alias, SFName, SFPhone, STName, Company, Address1, Address2, STPhone, PostalCode, DeclaredValue
- WorldShip tabs: (reuse current coords)
- WorldShip fields: (reuse current coords)
- Buttons: PersonalButton, BusinessButton
*/

SetTitleMatchMode, 2
SetDefaultMouseSpeed, 0
CoordMode, Mouse, Window

; Window targets
intraWinTitle := "Intra: Shipping Request Form ahk_exe firefox.exe"
worldShipTitle := "UPS WorldShip"

; Normalized window size used for ratio/absolute clicks (matches Intra_Buttons)
intraWinW := 1530
intraWinH := 1399
neutralClickR := {x: 0.895425, y: 0.50683} ; same neutral spot used in Intra_Buttons

; Intra: Shipping Request Form field coordinates (window-relative pixels at normalized size)
intraFields := {}
intraFields.CostCenter    := {x: 0, y: 0}
intraFields.Alias         := {x: 0, y: 0}
intraFields.SFName        := {x: 0, y: 0}
intraFields.SFPhone       := {x: 0, y: 0}
intraFields.STName        := {x: 0, y: 0}
intraFields.Company       := {x: 0, y: 0}
intraFields.Address1      := {x: 0, y: 0}
intraFields.Address2      := {x: 0, y: 0}
intraFields.STPhone       := {x: 0, y: 0}
intraFields.PostalCode    := {x: 0, y: 0}
intraFields.DeclaredValue := {x: 0, y: 0}

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

; Button targets (window-relative pixels)
PersonalButtonX := 0
PersonalButtonY := 0
BusinessButtonX := 0
BusinessButtonY := 0

return  ; end of auto-execute section

Esc::ExitApp

^!b:: ; Business Form (mirrors ^!p without offsets/cost center)
    FocusIntraWindow()
    EnsureIntraWindow()
    Sleep 50
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
    Sleep 2000
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

    FocusIntraWindow()
    EnsureIntraWindow()
    Sleep 50
    NeutralAndHome()
    FocusIntraWindow()
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
        Sleep 2000
    }
    else
    {
        FocusWorldShipWindow()
        PasteFieldAt(worldShipFields.Company.x, worldShipFields.Company.y, stName)
        Sleep 150
        MouseClick, left, % worldShipFields.Ref2.x, worldShipFields.Ref2.y
        Sleep 250  ; allow WorldShip address book fill to process
    }
    FocusWorldShipWindow()
    PasteFieldAt(worldShipFields.STName.x, worldShipFields.STName.y, stName)

    ; Scroll Down
    FocusIntraWindow()
    EnsureIntraWindow()
    Sleep 150
    NeutralClick()
    Sleep 250
    Loop 10
        {
            Sleep 25
            SendInput, {WheelDown}
            Sleep 25
        }
    Sleep 150
    FocusIntraWindow()
    Address1 := CopyFieldAt(intraFields.Address1.x, intraFields.Address1.y)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    MouseClick, left, % worldShipTabs.Address1.x, worldShipTabs.Address1.y
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
    MouseClick, left, % worldShipTabs.Address2.x, worldShipTabs.Address2.y
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
    MouseClick, left, % worldShipTabs.STPhone.x, worldShipTabs.STPhone.y
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
    MouseClick, left, % worldShipFields.PostalCode.x, worldShipFields.PostalCode.y
    Sleep 150
    FocusIntraWindow()
    Sleep 50
    ; Copy last token in Declared Value via End then Ctrl+Shift+Left to avoid whole-field select
    ClipSaved := ClipboardAll
    Clipboard :=
    MouseClick, left, % intraFields.DeclaredValue.x, intraFields.DeclaredValue.y, 2
    Sleep 150
    SendInput, {End}
    Sleep 120
    SendInput, ^+{Left}
    Sleep 120
    Clipboard :=  ; clear before copy to avoid stale values
    SendInput, ^c
    ClipWait, 0.5
    DeclaredValue := Clipboard
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
    Sleep 100
    Send @amazon.com
    Sleep 150
    MouseClick, left, % worldShipTabs.Options.x, worldShipTabs.Options.y
    Sleep 150
    MouseClick, left, % worldShipTabs.QVN.x, worldShipTabs.QVN.y
    Sleep 150
    MouseClick, left, % worldShipTabs.Recipients.x, worldShipTabs.Recipients.y
    Sleep 250
    Send {Tab 2}  ; move focus to QVN email field
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
    Send @amazon.com
    Sleep 150
    Send {Enter}
return

^!p:: ; Personal Form
    offsetY := -90              ; Y offset (General) 
    ScrollOffsetY := -180      ; Y offset (Scrolled-Down Section)

    FocusIntraWindow()
    EnsureIntraWindow()
    Sleep 50
    NeutralAndHome()
    Sleep 50
    MouseClick, left, %PersonalButtonX%, %PersonalButtonY%, 2
    Sleep 150
    FocusIntraWindow()
    sfName := CopyFieldAt(intraFields.SFName.x, intraFields.SFName.y + offsetY)
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
    Sleep 2000
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
    sfPhone := CopyFieldAt(intraFields.SFPhone.x, intraFields.SFPhone.y + offsetY)
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
    company := CopyFieldAt(intraFields.Company.x, intraFields.Company.y + offsetY)

    FocusIntraWindow()
    EnsureIntraWindow()
    Sleep 50
    NeutralAndHome()
    FocusIntraWindow()
    stName := CopyFieldAt(intraFields.STName.x, intraFields.STName.y + offsetY)
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
        Sleep 250  ; allow WorldShip address book fill to process
    }
    FocusWorldShipWindow()
    PasteFieldAt(worldShipFields.STName.x, worldShipFields.STName.y, stName)

    ; Scroll Down
    FocusIntraWindow()
    EnsureIntraWindow()
    Sleep 150
    NeutralClick()
    Sleep 250
    Loop 10
        {
            Sleep 25
            SendInput, {WheelDown}
            Sleep 25
        }
    Sleep 150
    FocusIntraWindow()
    Address1 := CopyFieldAt(intraFields.Address1.x, intraFields.Address1.y + ScrollOffsetY)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    MouseClick, left, % worldShipTabs.Address1.x, worldShipTabs.Address1.y
    Sleep 150
    FocusWorldShipWindow()
    PasteFieldAt(worldShipFields.Address1.x, worldShipFields.Address1.y, Address1)

    FocusIntraWindow()
    EnsureIntraWindow()
    Sleep 50
    FocusIntraWindow()
    Address2 := CopyFieldAt(intraFields.Address2.x, intraFields.Address2.y + ScrollOffsetY)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    MouseClick, left, % worldShipTabs.Address2.x, worldShipTabs.Address2.y
    Sleep 150
    FocusWorldShipWindow()
    PasteFieldAt(worldShipFields.Address2.x, worldShipFields.Address2.y, Address2)

    FocusIntraWindow()
    EnsureIntraWindow()
    Sleep 50
    FocusIntraWindow()
    STPhone := CopyFieldAt(intraFields.STPhone.x, intraFields.STPhone.y + ScrollOffsetY)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    MouseClick, left, % worldShipTabs.STPhone.x, worldShipTabs.STPhone.y
    Sleep 150
    FocusWorldShipWindow()
    PasteFieldAt(worldShipFields.STPhone.x, worldShipFields.STPhone.y, STPhone)

    ; Clear PostalCode first in WorldShip (handles saved address autofill), then copy/paste
    FocusIntraWindow()
    EnsureIntraWindow()
    Sleep 50
    FocusIntraWindow()
    PostalCode := CopyFieldAt(intraFields.PostalCode.x, intraFields.PostalCode.y + ScrollOffsetY)
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
    Sleep 2000


    FocusIntraWindow()
    Sleep 50
    ; Copy last token in Declared Value via End then Ctrl+Shift+Left to avoid whole-field select
    ClipSaved := ClipboardAll
    Clipboard :=
    MouseClick, left, 410, 651, 2  ; absolute click for declared value (personal form)
    Sleep 150
    SendInput, {End}
    Sleep 120
    SendInput, ^+{Left}
    Sleep 120
    Clipboard :=  ; clear before copy to avoid stale values
    SendInput, ^c
    ClipWait, 0.5
    DeclaredValue := Clipboard
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

    ; alias paste in email field then select options-qvn-recipients-
    ; paste into qvnemail-then done, optionally implement a send enter
    FocusIntraWindow()
    EnsureIntraWindow()
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
    Sleep 250
    Send {Tab 2}  ; move focus to QVN email field
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

EnsureIntraWindow()
{
    global intraWinTitle
    ; Match the working dimensions used in Intra_Buttons. Adjust here if the target size changes.
    WinMove, %intraWinTitle%,, 1917, 0, 1530, 1399
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
