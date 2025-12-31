#Requires AutoHotkey v1
#NoEnv
#Warn
#SingleInstance, Force
; Use Event so Ctrl+C/V in WorldShip behave like physical keypresses.
SendMode Input
; Moderate key delay for stability.
SetKeyDelay, 25, 25
SetWorkingDir %A_ScriptDir%
SetTitleMatchMode, 2
CoordMode, Mouse, Window

; WorldShip coordinates pulled from DSRF-to-UPS_WS script
worldShipTitle := "UPS WorldShip"
qvnTitle := "Quantum View Notify Recipients"
worldShipTabs := {}
worldShipTabs.ShipTo  := {x: 47,  y: 162}
worldShipTabs.ShipFrom := {x: 99,  y: 162}
worldShipTabs.Service := {x: 323, y: 162}
worldShipFields := {}
worldShipFields.STPhone := {x: 85,  y: 485}  ; Ship To Phone
worldShipFields.SFPhone := {x: 85,  y: 485}  ; Ship From Phone (same coords)
worldShipFields.STName  := {x: 85,  y: 280}  ; Attention / Name
worldShipFields.Company := {x: 78,  y: 241}  ; Company (Ship From)
worldShipFields.Ref2    := {x: 721, y: 345}  ; Reference 2

^Esc::Reload

#If (WinActive(worldShipTitle) || WinActive(qvnTitle))
~Alt::
    Sleep 250
return

!a::
    ; If Alt was held, make sure it's released before typing.
    SendInput, {Alt up}
    KeyWait, Alt
    Sleep 50
    SendInput, @amazon.com
return

$!Tab::
    Send, {Alt up}
    Loop 6
    {          
        SendInput, {Tab}
        Sleep 25
    }
return

!1::
    Loop 2
    {
        Sleep 5
        SendInput, {Tab}
        Sleep 25
    }
return

!2::
    Loop 8
    {
        Sleep 5
        SendInput, {Tab}
        Sleep 25
    }
return
#If

; Alt+P: copy Ship To Phone, switch to Ship From, paste into SF Phone.
!p::
    if (!FocusWorldShip())
        return
    Clipboard := ""
    EnsureWorldShipTop()
    ; Ship To tab
    MouseClick, left, % worldShipTabs.ShipTo.x, % worldShipTabs.ShipTo.y
    Sleep 30
    copiedText := CopyFieldAt(worldShipFields.STPhone.x, worldShipFields.STPhone.y)
    Sleep 30
    ; Ship From tab
    MouseClick, left, % worldShipTabs.ShipFrom.x, % worldShipTabs.ShipFrom.y
    Sleep 30
    PasteFieldAt(worldShipFields.SFPhone.x, worldShipFields.SFPhone.y, copiedText)
return

; Alt+N: Copy attention/name from Ship To; if empty, copy Ship From company; paste into Ref2.
!n::
    if (!FocusWorldShip())
        return
    Clipboard := ""
    EnsureWorldShipTop()
    ; Ship From tab -> Service -> copy Company -> paste Ref2
    MouseClick, left, % worldShipTabs.ShipFrom.x, % worldShipTabs.ShipFrom.y
    Sleep 30
    MouseClick, left, % worldShipTabs.Service.x, % worldShipTabs.Service.y
    Sleep 30
    copiedText := CopyFieldAt(worldShipFields.Company.x, worldShipFields.Company.y)
    MouseClick, left, % worldShipFields.Ref2.x, % worldShipFields.Ref2.y
    Sleep 30
    PasteFieldAt(worldShipFields.Ref2.x, worldShipFields.Ref2.y, copiedText)
return

FocusWorldShip()
{
    if WinExist("UPS WorldShip")
    {
        WinActivate
        WinWaitActive, UPS WorldShip,, 1
        return !ErrorLevel
    }
    return false
}

EnsureWorldShipTop()
{
    MouseClick, left, 430, 335
    Sleep 60
    SendInput, {WheelUp 5}
    Sleep 80
}

CopyFieldAt(x, y)
{
    ClipSaved := ClipboardAll
    Clipboard :=
    MouseClick, left, %x%, %y%
    Sleep 120
    SendEvent, {End}
    Sleep 100
    SendEvent, +{Home}
    Sleep 100
    text := ""
    ; Try up to 5 copies; stop on first non-empty capture.
    Loop, 5
    {
        Clipboard :=
        SendEvent, {Ctrl down}c{Ctrl up}
        Sleep 100
        ClipWait, 0.5
        if (!ErrorLevel && Clipboard != "")
        {
            text := Clipboard
            break
        }
        Sleep 100
    }
    Clipboard := ClipSaved
    return text
}

PasteFieldAt(x, y, text)
{
    ClipSaved := ClipboardAll
    Clipboard := text
    MouseClick, left, %x%, %y%
    Sleep 120
    if (text = "")
    {
        Clipboard := ClipSaved
        return
    }
    SendEvent, {Home}
    Sleep 100
    SendEvent, +{End}
    Sleep 100
    SendEvent, {Delete}
    Sleep 100
    SendEvent, {Ctrl down}v{Ctrl up}
    Sleep 100
    Clipboard := ClipSaved
}
