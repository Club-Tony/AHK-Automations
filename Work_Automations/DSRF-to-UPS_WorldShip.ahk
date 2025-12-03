#Requires AutoHotkey v1
#NoEnv ; Prevents Unnecessary Environment Variable lookup
#Warn ; Warn All (All Warnings Enabled)
#SingleInstance, Force  ; Reload without prompt when Esc is pressed.
SendMode Input ; Send works as SendInput
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

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
intraFields.DeclaredValue := {x: 410, y: 830}

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

return  ; end of auto-execute section

Esc::ExitApp

^!b:: ; Business Form
    FocusIntraWindow()
    EnsureIntraWindow()
    Sleep 50
    NeutralAndHome()
    Sleep 50
    SendInput, {WheelUp 15} 
    Sleep 400                
    costCenter := CopyFieldAt(intraFields.CostCenter.x, intraFields.CostCenter.y)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    MouseClick, left, % worldShipTabs.ShipTo.x, worldShipTabs.ShipTo.y
    Sleep 150
    PasteFieldAt(worldShipFields.Ref1.x, worldShipFields.Ref1.y, costCenter)

    FocusIntraWindow()
    EnsureIntraWindow()
    Sleep 50
    NeutralAndHome()
    sfName := CopyFieldAt(intraFields.SFName.x, intraFields.SFName.y)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    MouseClick, left, % worldShipTabs.Service.x, worldShipTabs.Service.y
    Sleep 150
    MouseClick, left, % worldShipTabs.ShipFrom.x, worldShipTabs.ShipFrom.y
    Sleep 150
    ; Also drop SF Name into Company for Ship From before Name/Attn
    PasteFieldAt(worldShipFields.Company.x, worldShipFields.Company.y, sfName)
    Sleep 150
    MouseClick, left, % worldShipFields.SFName.x, worldShipFields.SFName.y
    Sleep 2500
    PasteFieldAt(worldShipFields.SFName.x, worldShipFields.SFName.y, sfName)
    Sleep 150
    EnsureWorldShipTop()
    Sleep 150
    PasteFieldAt(worldShipFields.Ref2.x, worldShipFields.Ref2.y, sfName)

    FocusIntraWindow()
    EnsureIntraWindow()
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
    EnsureIntraWindow()
    Sleep 50
    NeutralAndHome()
    company := CopyFieldAt(intraFields.Company.x, intraFields.Company.y)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    Sleep 150
    PasteFieldAt(worldShipFields.Company.x, worldShipFields.Company.y, company)

    FocusIntraWindow()
    EnsureIntraWindow()
    Sleep 50
    NeutralAndHome()
    stName := CopyFieldAt(intraFields.STName.x, intraFields.STName.y)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    Sleep 150
    PasteFieldAt(worldShipFields.STName.x, worldShipFields.STName.y, stName)

    ; Scroll Down
    FocusIntraWindow()
    EnsureIntraWindow()
    Sleep 150
    NeutralClick()
    Sleep 250
    Loop 10
        {
            Sleep 50
            SendInput, {WheelDown}
            Sleep 50
        }
    Sleep 150   
    Address1 := CopyFieldAt(intraFields.Address1.x, intraFields.Address1.y)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    MouseClick, left, % worldShipTabs.Address1.x, worldShipTabs.Address1.y
    Sleep 150
    PasteFieldAt(worldShipFields.Address1.x, worldShipFields.Address1.y, Address1)

    FocusIntraWindow()
    EnsureIntraWindow()
    Sleep 50
    Address2 := CopyFieldAt(intraFields.Address2.x, intraFields.Address2.y)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    MouseClick, left, % worldShipTabs.Address2.x, worldShipTabs.Address2.y
    Sleep 150
    PasteFieldAt(worldShipFields.Address2.x, worldShipFields.Address2.y, Address2)

    FocusIntraWindow()
    EnsureIntraWindow()
    Sleep 50
    STPhone := CopyFieldAt(intraFields.STPhone.x, intraFields.STPhone.y)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    MouseClick, left, % worldShipTabs.STPhone.x, worldShipTabs.STPhone.y
    Sleep 150
    PasteFieldAt(worldShipFields.STPhone.x, worldShipFields.STPhone.y, STPhone)

    FocusIntraWindow()
    EnsureIntraWindow()
    Sleep 50
    PostalCode := CopyFieldAt(intraFields.PostalCode.x, intraFields.PostalCode.y)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    Sleep 250
    PasteFieldAt(worldShipFields.PostalCode.x, worldShipFields.PostalCode.y, PostalCode)    
    Sleep 250
    NeutralClick()
    FocusIntraWindow()
    EnsureIntraWindow()
    Sleep 2500
    ; Copy last token in Declared Value via End then Ctrl+Shift+Left to avoid whole-field select
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

^!p:: ; Personal Form
    offsetY := -85  ; personal form intra Y offset

    FocusIntraWindow()
    EnsureIntraWindow()
    Sleep 50
    NeutralAndHome()
    Sleep 50
    SendInput, {WheelUp 15}
    Sleep 400
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
    NeutralClick()
    Sleep 2500
    PasteFieldAt(worldShipFields.SFName.x, worldShipFields.SFName.y, sfName)
    Sleep 150
    EnsureWorldShipTop()
    Sleep 150
    PasteFieldAt(worldShipFields.Ref2.x, worldShipFields.Ref2.y, sfName)

    FocusIntraWindow()
    EnsureIntraWindow()
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
    EnsureIntraWindow()
    Sleep 50
    NeutralAndHome()
    company := CopyFieldAt(intraFields.Company.x, intraFields.Company.y + offsetY)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    Sleep 150
    PasteFieldAt(worldShipFields.Company.x, worldShipFields.Company.y, company)

    FocusIntraWindow()
    EnsureIntraWindow()
    Sleep 50
    NeutralAndHome()
    stName := CopyFieldAt(intraFields.STName.x, intraFields.STName.y + offsetY)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    Sleep 150
    PasteFieldAt(worldShipFields.STName.x, worldShipFields.STName.y, stName)

    ; Scroll Down
    FocusIntraWindow()
    EnsureIntraWindow()
    Sleep 150
    NeutralClick()
    Sleep 250
    Loop 10
        {
            Sleep 50
            SendInput, {WheelDown}
            Sleep 50
        }
    Sleep 150
    Address1 := CopyFieldAt(intraFields.Address1.x, intraFields.Address1.y + offsetY)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    MouseClick, left, % worldShipTabs.Address1.x, worldShipTabs.Address1.y
    Sleep 150
    PasteFieldAt(worldShipFields.Address1.x, worldShipFields.Address1.y, Address1)

    FocusIntraWindow()
    EnsureIntraWindow()
    Sleep 50
    Address2 := CopyFieldAt(intraFields.Address2.x, intraFields.Address2.y + offsetY)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    MouseClick, left, % worldShipTabs.Address2.x, worldShipTabs.Address2.y
    Sleep 150
    PasteFieldAt(worldShipFields.Address2.x, worldShipFields.Address2.y, Address2)

    FocusIntraWindow()
    EnsureIntraWindow()
    Sleep 50
    STPhone := CopyFieldAt(intraFields.STPhone.x, intraFields.STPhone.y + offsetY)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    MouseClick, left, % worldShipTabs.STPhone.x, worldShipTabs.STPhone.y
    Sleep 150
    PasteFieldAt(worldShipFields.STPhone.x, worldShipFields.STPhone.y, STPhone)

    FocusIntraWindow()
    EnsureIntraWindow()
    Sleep 50
    PostalCode := CopyFieldAt(intraFields.PostalCode.x, intraFields.PostalCode.y + offsetY)
    FocusWorldShipWindow()
    EnsureWorldShipTop()
    Sleep 250
    PasteFieldAt(worldShipFields.PostalCode.x, worldShipFields.PostalCode.y, PostalCode)    
    Sleep 250
    NeutralClick()
    FocusIntraWindow()
    EnsureIntraWindow()
    Sleep 2500
    ; Copy last token in Declared Value via End then Ctrl+Shift+Left to avoid whole-field select
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
