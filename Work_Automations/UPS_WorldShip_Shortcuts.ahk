#Requires AutoHotkey v1
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Warn  ; Enable warnings to assist with detecting common errors.
#SingleInstance, Force  ; Reload without prompt when Esc is pressed.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetTitleMatchMode, 2  ; Allow partial title matches for WorldShip.

SetKeyDelay, 25
worldShipTitle := "UPS WorldShip"
qvnTitle := "Quantum View Notify Recipients"

^Esc::Reload

#If (WinActive(worldShipTitle) || WinActive(qvnTitle))
!a::
    ; Use SendEvent so the SetKeyDelay applies even though SendMode is Input.
    SendEvent, @amazon.com
return

$!Tab::
    ; Release Alt first so Windows doesn't trigger task switch.
    Send, {Alt up}
    Loop 6
    {
        Send, {Tab}
        Sleep, 50
    }
    KeyWait, Alt  ; ensure Alt is up before continuing.
return

!1::
    Loop 2
    {
        Send, {Tab}
        Sleep, 50
    }
return

!2::
    Loop 8
    {
        Send, {Tab}
        Sleep, 50
    }
return
#If