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

; Window targets
intraWinTitle := "Intra: Shipping Request Form ahk_exe firefox.exe"
worldShipTitle := "UPS WorldShip"

; Normalized window size used for ratio/absolute clicks (matches Intra_Buttons)
intraWinW := 1530
intraWinH := 1399
neutralClickR := {x: 0.895425, y: 0.50683} ; same neutral spot used in Intra_Buttons

; Intra: Shipping Request Form field coordinates (window-relative pixels at normalized size)
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

; Button targets (window-relative pixels)
PersonalButtonX := 274
PersonalButtonY := 486
BusinessButtonX := 365
BusinessButtonY := 486

return  ; end of auto-execute section

Esc::ExitApp

^!b:: ; Business Form (mirrors ^!p without offsets/cost center)
    FocusIntraWindow()
    EnsureIntraWindow()
    Sleep 25
    NeutralAndHome()
    Sleep 25
    SendInput, {WheelUp 15}
    Sleep 75
    MouseClick, left, %BusinessButtonX%, %BusinessButtonY%, 2
    Sleep 75
    FocusIntraWindow()
    costCenter := CopyFieldAt(intraFields.CostCenter.x, intraFields.CostCenter.y)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    Sleep 75
    MouseClick, left, % worldShipTabs.ShipTo.x, worldShipTabs.ShipTo.y
    Sleep 75
    MouseClick, left, % worldShipTabs.Service.x, worldShipTabs.Service.y
    Sleep 75
    PasteFieldAt(worldShipFields.Ref1.x, worldShipFields.Ref1.y, costCenter)
    Sleep 75
    FocusIntraWindow()
    sfName := CopyFieldAt(intraFields.SFName.x, intraFields.SFName.y)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    Sleep 75
    MouseClick, left, % worldShipTabs.Service.x, worldShipTabs.Service.y  ; ensure Service tab active before first paste
    Sleep 75
    MouseClick, left, % worldShipTabs.ShipFrom.x, worldShipTabs.ShipFrom.y
    Sleep 75
    FocusWorldShipWindow()
    PasteFieldAt(worldShipFields.Company.x, worldShipFields.Company.y, sfName)
    Sleep 75
    MouseClick, left, % worldShipFields.Ref2.x, worldShipFields.Ref2.y
    Sleep 1000
    FocusWorldShipWindow()
    PasteFieldAt(worldShipFields.SFName.x, worldShipFields.SFName.y, sfName)
    Sleep 75
    EnsureWorldShipTop()
    Sleep 75
    FocusWorldShipWindow()
    PasteFieldAt(worldShipFields.Ref2.x, worldShipFields.Ref2.y, sfName)

    FocusIntraWindow()
    EnsureIntraWindow()
    Sleep 25
    NeutralAndHome()
    FocusIntraWindow()
    sfPhone := CopyFieldAt(intraFields.SFPhone.x, intraFields.SFPhone.y)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    Sleep 75
    FocusWorldShipWindow()
    PasteFieldAt(worldShipFields.SFPhone.x, worldShipFields.SFPhone.y, sfPhone)
    Sleep 75
    FocusWorldShipWindow()
    MouseClick, left, % worldShipTabs.ShipTo.x, worldShipTabs.ShipTo.y
    Sleep 75
    MouseClick, left, % worldShipTabs.Service.x, worldShipTabs.Service.y

    FocusIntraWindow()
    EnsureIntraWindow()
    Sleep 25
    NeutralAndHome()
    FocusIntraWindow()
    company := CopyFieldAt(intraFields.Company.x, intraFields.Company.y)

    FocusIntraWindow()
    EnsureIntraWindow()
    Sleep 25
    NeutralAndHome()
    FocusIntraWindow()
    stName := CopyFieldAt(intraFields.STName.x, intraFields.STName.y)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    Sleep 75
    if (company != "")
    {
        FocusWorldShipWindow()
        PasteFieldAt(worldShipFields.Company.x, worldShipFields.Company.y, company)
        Sleep 75
        MouseClick, left, % worldShipFields.Ref2.x, worldShipFields.Ref2.y
        Sleep 1000
    }
    else
    {
        FocusWorldShipWindow()
        PasteFieldAt(worldShipFields.Company.x, worldShipFields.Company.y, stName)
        Sleep 75
        MouseClick, left, % worldShipFields.Ref2.x, worldShipFields.Ref2.y
        Sleep 125  ; allow WorldShip address book fill to process
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
    Sleep 75
    FocusWorldShipWindow()
    PasteFieldAt(worldShipFields.Address1.x, worldShipFields.Address1.y, Address1)

    FocusIntraWindow()
    EnsureIntraWindow()
    Sleep 25
    FocusIntraWindow()
    Address2 := CopyFieldAt(intraFields.Address2.x, intraFields.Address2.y)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    MouseClick, left, % worldShipTabs.Address2.x, worldShipTabs.Address2.y
    Sleep 75
    FocusWorldShipWindow()
    PasteFieldAt(worldShipFields.Address2.x, worldShipFields.Address2.y, Address2)

    FocusIntraWindow()
    EnsureIntraWindow()
    Sleep 25
    FocusIntraWindow()
    STPhone := CopyFieldAt(intraFields.STPhone.x, intraFields.STPhone.y)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    MouseClick, left, % worldShipTabs.STPhone.x, worldShipTabs.STPhone.y
    Sleep 75
    FocusWorldShipWindow()
    PasteFieldAt(worldShipFields.STPhone.x, worldShipFields.STPhone.y, STPhone)

    ; Clear PostalCode first in WorldShip (handles saved address autofill), then copy/paste
    FocusIntraWindow()
    EnsureIntraWindow()
    Sleep 25
    FocusIntraWindow()
    PostalCode := CopyFieldAt(intraFields.PostalCode.x, intraFields.PostalCode.y)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    ClipSaved := ClipboardAll
    Clipboard := PostalCode
    MouseClick, left, % worldShipFields.PostalCode.x, worldShipFields.PostalCode.y
    Sleep 75
    SendInput, {Home}
    Sleep 40
    SendInput, +{End}
    Sleep 40
    SendInput, {Delete}
    Sleep 125
    SendInput, %PostalCode%
    Sleep 125
    MouseClick, left, % worldShipFields.Ref2.x, worldShipFields.Ref2.y
    Sleep 1000  ; allow any address book/city-state prompts to settle
    MouseClick, left, % worldShipFields.PostalCode.x, worldShipFields.PostalCode.y
    Sleep 75
    FocusIntraWindow()
    Sleep 25
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
    Sleep 25
    NeutralAndHome()
    Sleep 75
    Alias := CopyFieldAt(intraFields.Alias.x, intraFields.Alias.y)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    Sleep 75
    PasteFieldAt(worldShipFields.STEmail.x, worldShipFields.STEmail.y, Alias)
    Sleep 75
    Send {End}
    Sleep 50
    Send @amazon.com
    Sleep 75
    MouseClick, left, % worldShipTabs.Options.x, worldShipTabs.Options.y
    Sleep 75
    MouseClick, left, % worldShipTabs.QVN.x, worldShipTabs.QVN.y
    Sleep 75
    MouseClick, left, % worldShipTabs.Recipients.x, worldShipTabs.Recipients.y
    Sleep 125
    Send {Tab 2}  ; move focus to QVN email field
    Sleep 75
    ClipSaved := ClipboardAll
    Clipboard := Alias
    SendInput, ^v
    Sleep 75
    Clipboard := ClipSaved
    ClipSaved := ""
    Sleep 75
    Send {End}
    Sleep 75
    Send @amazon.com
    Sleep 75
    Send {Enter}
return

^!p:: ; Personal Form
    offsetY := -90              ; Y offset (General) 
    ScrollOffsetY := -180      ; Y offset (Scrolled-Down Section)

    FocusIntraWindow()
    EnsureIntraWindow()
    Sleep 25
    NeutralAndHome()
    Sleep 25
    MouseClick, left, %PersonalButtonX%, %PersonalButtonY%, 2
    Sleep 75
    FocusIntraWindow()
    sfName := CopyFieldAt(intraFields.SFName.x, intraFields.SFName.y + offsetY)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    Sleep 75
    MouseClick, left, % worldShipTabs.Service.x, worldShipTabs.Service.y  ; ensure Service tab active before first paste
    Sleep 75
    MouseClick, left, % worldShipTabs.ShipFrom.x, worldShipTabs.ShipFrom.y
    Sleep 75
    FocusWorldShipWindow()
    PasteFieldAt(worldShipFields.Company.x, worldShipFields.Company.y, sfName)
    Sleep 75
    MouseClick, left, % worldShipFields.Ref2.x, worldShipFields.Ref2.y
    Sleep 1000
    FocusWorldShipWindow()
    PasteFieldAt(worldShipFields.SFName.x, worldShipFields.SFName.y, sfName)
    Sleep 75
    EnsureWorldShipTop()
    Sleep 75
    FocusWorldShipWindow()
    PasteFieldAt(worldShipFields.Ref2.x, worldShipFields.Ref2.y, sfName)

    FocusIntraWindow()
    EnsureIntraWindow()
    Sleep 25
    NeutralAndHome()
    FocusIntraWindow()
    sfPhone := CopyFieldAt(intraFields.SFPhone.x, intraFields.SFPhone.y + offsetY)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    Sleep 75
    FocusWorldShipWindow()
    PasteFieldAt(worldShipFields.SFPhone.x, worldShipFields.SFPhone.y, sfPhone)
    Sleep 75
    FocusWorldShipWindow()
    MouseClick, left, % worldShipTabs.ShipTo.x, worldShipTabs.ShipTo.y

    FocusIntraWindow()
    EnsureIntraWindow()
    Sleep 25
    NeutralAndHome()
    FocusIntraWindow()
    company := CopyFieldAt(intraFields.Company.x, intraFields.Company.y + offsetY)

    FocusIntraWindow()
    EnsureIntraWindow()
    Sleep 25
    NeutralAndHome()
    FocusIntraWindow()
    stName := CopyFieldAt(intraFields.STName.x, intraFields.STName.y + offsetY)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    Sleep 75
    if (company != "")
    {
        FocusWorldShipWindow()
        PasteFieldAt(worldShipFields.Company.x, worldShipFields.Company.y, company)
        Sleep 75
        MouseClick, left, % worldShipFields.Ref2.x, worldShipFields.Ref2.y
        Sleep 1000
    }
    else
    {
        FocusWorldShipWindow()
        PasteFieldAt(worldShipFields.Company.x, worldShipFields.Company.y, stName)
        Sleep 75
        MouseClick, left, % worldShipFields.Ref2.x, worldShipFields.Ref2.y
        Sleep 125  ; allow WorldShip address book fill to process
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
    Sleep 75
    FocusWorldShipWindow()
    PasteFieldAt(worldShipFields.Address1.x, worldShipFields.Address1.y, Address1)

    FocusIntraWindow()
    EnsureIntraWindow()
    Sleep 25
    FocusIntraWindow()
    Address2 := CopyFieldAt(intraFields.Address2.x, intraFields.Address2.y + ScrollOffsetY)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    MouseClick, left, % worldShipTabs.Address2.x, worldShipTabs.Address2.y
    Sleep 75
    FocusWorldShipWindow()
    PasteFieldAt(worldShipFields.Address2.x, worldShipFields.Address2.y, Address2)

    FocusIntraWindow()
    EnsureIntraWindow()
    Sleep 25
    FocusIntraWindow()
    STPhone := CopyFieldAt(intraFields.STPhone.x, intraFields.STPhone.y + ScrollOffsetY)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    MouseClick, left, % worldShipTabs.STPhone.x, worldShipTabs.STPhone.y
    Sleep 75
    FocusWorldShipWindow()
    PasteFieldAt(worldShipFields.STPhone.x, worldShipFields.STPhone.y, STPhone)

    ; Clear PostalCode first in WorldShip (handles saved address autofill), then copy/paste
    FocusIntraWindow()
    EnsureIntraWindow()
    Sleep 25
    FocusIntraWindow()
    PostalCode := CopyFieldAt(intraFields.PostalCode.x, intraFields.PostalCode.y + ScrollOffsetY)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    ClipSaved := ClipboardAll
    Clipboard := PostalCode
    MouseClick, left, % worldShipFields.PostalCode.x, worldShipFields.PostalCode.y
    Sleep 75
    SendInput, {Home}
    Sleep 40
    SendInput, +{End}
    Sleep 40
    SendInput, {Delete}
    Sleep 125
    SendInput, %PostalCode%  
    Sleep 125
    MouseClick, left, % worldShipFields.Ref2.x, worldShipFields.Ref2.y
    Sleep 1000


    FocusIntraWindow()
    Sleep 25
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
    Sleep 25
    NeutralAndHome()
    Sleep 75
    Alias := CopyFieldAt(intraFields.Alias.x, intraFields.Alias.y + offsetY)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    Sleep 75
    PasteFieldAt(worldShipFields.STEmail.x, worldShipFields.STEmail.y, Alias)
    Sleep 75
    Send {End}
    Sleep 50
    Send @amazon.com
    Sleep 75
    MouseClick, left, % worldShipTabs.Options.x, worldShipTabs.Options.y
    Sleep 75
    MouseClick, left, % worldShipTabs.QVN.x, worldShipTabs.QVN.y
    Sleep 75
    MouseClick, left, % worldShipTabs.Recipients.x, worldShipTabs.Recipients.y
    Sleep 125
    Send {Tab 2}  ; move focus to QVN email field
    Sleep 75
    ClipSaved := ClipboardAll
    Clipboard := Alias
    SendInput, ^v
    Sleep 75
    Clipboard := ClipSaved
    ClipSaved := ""
    Sleep 75
    Send {End}
    Sleep 75
    Send @amazon.com
    Sleep 75
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
    Sleep 75
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
    Sleep 75
    SendInput, {WheelUp 5}
    Sleep 100
}

NeutralClick()
{
    global neutralClickR
    WinGetPos,,, winW, winH, A
    targetX := Floor(winW * neutralClickR.x)
    targetY := Floor(winH * neutralClickR.y)
    MouseClick, left, %targetX%, %targetY%
    Sleep 100
}

NeutralAndHome()
{
    global neutralClickR
    WinGetPos,,, winW, winH, A
    targetX := Floor(winW * neutralClickR.x)
    targetY := Floor(winH * neutralClickR.y)
    MouseClick, left, %targetX%, %targetY%
    Sleep 100
    SendInput, ^{Home}
    Sleep 100
}

CopyFieldAt(x, y)
{
    local ClipSaved, text
    ClipSaved := ClipboardAll
    Clipboard :=
    MouseClick, left, %x%, %y%
    Sleep 75
    SendInput, ^a
    Sleep 40
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
    Sleep 75
    ; WorldShip ignores Ctrl+A, so clear via Home -> Shift+End -> Delete
    SendInput, {Home}
    Sleep 40
    SendInput, +{End}
    Sleep 40
    SendInput, {Delete}
    Sleep 60
    SendInput, ^v
    Sleep 50
    Clipboard := ClipSaved
    ClipSaved := ""
}

