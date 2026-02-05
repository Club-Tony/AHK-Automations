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
worldShipFields.STPhone := {x: 85,  y: 485}   ; Ship To Phone
worldShipFields.STEmail := {x: 210, y: 485}   ; Ship To Email
worldShipFields.SFPhone := {x: 85,  y: 485}   ; Ship From Phone (same coords)
worldShipFields.STName  := {x: 85,  y: 280}   ; Attention / Name
worldShipFields.Company := {x: 78,  y: 241}   ; Company (Ship From)
worldShipFields.Ref1    := {x: 728, y: 308}   ; Reference 1
worldShipFields.Ref2    := {x: 731, y: 347}   ; Reference 2
worldShipFields.Weight  := {x: 346, y: 535}   ; Package weight field
worldShipScaleToggle := {x: 316, y: 559}      ; Electronic scale toggle button (from DSRF-to-UPS_WS-Paste)
worldShipServiceMenu := {}
worldShipServiceMenu.Selection := {x: 400, y: 231}       ; UPS Service Selection dropdown
worldShipServiceMenu.NextDayAir := {x: 400, y: 262}
worldShipServiceMenu.NextDayAirSaver := {x: 400, y: 278}
worldShipServiceMenu.SecondDayAir := {x: 400, y: 308}
worldShipServiceMenu.ThreeDaySelect := {x: 400, y: 323}
worldShipServiceMenu.Ground := {x: 400, y: 338}
worldShipTabs.Options := {x: 372, y: 162}
worldShipTabs.QVN := {x: 381, y: 282}
worldShipTabs.Recipients := {x: 560, y: 253}

#If WinActive(worldShipTitle)
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

!e::
    SendInput, {Alt up}
    EnsureWorldShipTop()
    MouseClick, left, % worldShipTabs.ShipTo.x, % worldShipTabs.ShipTo.y
    Sleep 50
    copiedText := CopyFieldAt(worldShipFields.STEmail.x, worldShipFields.STEmail.y)
    Clipboard := copiedText
    Sleep 100
    MouseClick, left, % worldShipTabs.Options.x, % worldShipTabs.Options.y
    Sleep 150
    MouseClick, left, % worldShipTabs.QVN.x, % worldShipTabs.QVN.y, 2
    Sleep 150
    MouseClick, left, % worldShipTabs.Recipients.x, % worldShipTabs.Recipients.y
    Sleep 200
    WinWaitActive, %qvnTitle%,, 2
    Sleep 100
    Loop 2
    {
        SendInput, {Tab}
        Sleep 40
    }
return

!1::
    SendInput, {Alt up}
    MouseClick, left, % worldShipTabs.Service.x, % worldShipTabs.Service.y
    Sleep 50
    MouseClick, left, % worldShipFields.Ref1.x, % worldShipFields.Ref1.y
return

!2::
    SendInput, {Alt up}
    MouseClick, left, % worldShipTabs.Service.x, % worldShipTabs.Service.y
    Sleep 50
    MouseClick, left, % worldShipFields.Ref2.x, % worldShipFields.Ref2.y
return

; Alt+Tab: highlight current field (End, then Ctrl+Shift+Home)
$!Tab::
    Send, {Alt up}
    SendInput, {End}
    Sleep 50
    SendInput, ^+{Home}
return

!s::
    SendInput, {Alt up}
    MouseClick, left, % worldShipTabs.Service.x, % worldShipTabs.Service.y
    Sleep 50
    MouseClick, left, % worldShipServiceMenu.Selection.x, % worldShipServiceMenu.Selection.y
return

^!q::
    SendInput, {Alt up}
    MouseClick, left, % worldShipTabs.Options.x, % worldShipTabs.Options.y
    Sleep 150
    MouseClick, left, % worldShipTabs.QVN.x, % worldShipTabs.QVN.y
    Sleep 150
    MouseClick, left, % worldShipTabs.Recipients.x, % worldShipTabs.Recipients.y
return

!d::
    SendInput, {Alt up}
    MouseClick, left, % worldShipTabs.Service.x, % worldShipTabs.Service.y
    Sleep 50
    MouseClick, left, % worldShipFields.Weight.x, % worldShipFields.Weight.y, 2
return

!0::  ; Toggle electronic scale button (helps avoid scale lag)
    SendInput, {Alt up}
    EnsureWorldShipTop()
    MouseClick, left, % worldShipTabs.Service.x, % worldShipTabs.Service.y
    Sleep 50
    MouseClick, left, % worldShipScaleToggle.x, % worldShipScaleToggle.y
return

#If (WinActive(worldShipTitle) && !CoordHelperActive())
!g::
    SendInput, {Alt up}
    MouseClick, left, % worldShipTabs.Service.x, % worldShipTabs.Service.y
    Sleep 50
    MouseClick, left, % worldShipServiceMenu.Selection.x, % worldShipServiceMenu.Selection.y
    Sleep 250
    MouseMove, % worldShipServiceMenu.Ground.x, % worldShipServiceMenu.Ground.y
    Sleep 250
    MouseClick, left
return
#If WinActive(worldShipTitle)

!3::
    SendInput, {Alt up}
    MouseClick, left, % worldShipTabs.Service.x, % worldShipTabs.Service.y
    Sleep 50
    MouseClick, left, % worldShipServiceMenu.Selection.x, % worldShipServiceMenu.Selection.y
    Sleep 250
    MouseMove, % worldShipServiceMenu.NextDayAir.x, % worldShipServiceMenu.NextDayAir.y
    Sleep 250
    MouseClick, left
return

!4::
    SendInput, {Alt up}
    MouseClick, left, % worldShipTabs.Service.x, % worldShipTabs.Service.y
    Sleep 50
    MouseClick, left, % worldShipServiceMenu.Selection.x, % worldShipServiceMenu.Selection.y
    Sleep 250
    MouseMove, % worldShipServiceMenu.NextDayAirSaver.x, % worldShipServiceMenu.NextDayAirSaver.y
    Sleep 250
    MouseClick, left
return

!5::
    SendInput, {Alt up}
    MouseClick, left, % worldShipTabs.Service.x, % worldShipTabs.Service.y
    Sleep 50
    MouseClick, left, % worldShipServiceMenu.Selection.x, % worldShipServiceMenu.Selection.y
    Sleep 250
    MouseMove, % worldShipServiceMenu.SecondDayAir.x, % worldShipServiceMenu.SecondDayAir.y
    Sleep 250
    MouseClick, left
return

!6::
    SendInput, {Alt up}
    MouseClick, left, % worldShipTabs.Service.x, % worldShipTabs.Service.y
    Sleep 50
    MouseClick, left, % worldShipServiceMenu.Selection.x, % worldShipServiceMenu.Selection.y
    Sleep 250
    MouseMove, % worldShipServiceMenu.ThreeDaySelect.x, % worldShipServiceMenu.ThreeDaySelect.y
    Sleep 250
    MouseClick, left
return
#If

#If WinActive(qvnTitle)
~Alt::
    Sleep 250
return

$!Tab::
    Send, {Alt up}
    Loop 6
    {          
        SendInput, {Tab}
        Sleep 25
    }
return

!a::
    ; Allow @amazon.com paste in QVN window too.
    SendInput, {Alt up}
    KeyWait, Alt
    Sleep 50
    SendInput, @amazon.com
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

#If WinActive(worldShipTitle)
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
#If

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

CoordHelperActive()
{
    DetectHiddenWindows, On
    running := WinExist("Coord_Capture.ahk ahk_class AutoHotkey")
    DetectHiddenWindows, Off
    return running
}
