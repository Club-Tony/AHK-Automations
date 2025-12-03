#Requires AutoHotkey v1
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
#SingleInstance, Force  ; Reload without prompt when Esc is pressed.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

; Window Coordinates (Intra Desktop Client - Assign Recip):
; Window Position: x: -7 y: 0 w: 1322 h: 1339
; 200, 245 = Scan field
; 130, 850 = Name field
; 1035, 185 = enable item var lookup button
; 1060, 185 = applying to all item vars button
; 1100, 365 = Package type
; 1100, 535 = Pieces field
; 1100, 650 = BSC Location (Destination)
; 500, 1025 = Select Preset
; 350, 900 = Docksided items preset select 
; 725, 190 = Enable Parent Item #
; 500, 120 = Print Label button
; 40, 1300 = All Button

^Esc::Reload

#IfWinActive, Intra Desktop Client - Assign Recip
SetKeyDelay 150

!i::
    Mouseclick, left, 40, 1300, 2
    Sleep 150
    Mouseclick, left, 200, 245, 2
    Sleep 150
    Send ^n
    Sleep 450
    MouseClick, left, 130, 850, 2
    Sleep 150
    Send sea22, invictus
    Sleep 150
    Send {enter}
    Sleep 750
    loop, 2
    {
        Send {down}
        Sleep 150
    }
    Sleep 150
    Mouseclick, left, 1100, 365, 2
    Sleep 150
    Send p
    Sleep 150
    Send {enter}
    Sleep 150
    Mouseclick, left, 1100, 650, 2
    Sleep 150
    Send SEA22
    Sleep 150
    Send, {F5}
    Sleep 450
return

^i:: ; Ends at pieces field
    Mouseclick, left, 40, 1300, 2
    Sleep 150
    Mouseclick, left, 200, 245, 2
    Sleep 150
    Send ^n
    Sleep 450
    MouseClick, left, 130, 850, 2
    Sleep 150
    Send sea22, invictus
    Sleep 150
    Send {enter}
    Sleep 750
    loop, 2
    {
        Send {down}
        Sleep 150
    }
    Sleep 150
    Mouseclick, left, 1100, 365, 2
    Sleep 150
    Send p
    Sleep 150
    Send {enter}
    Sleep 150
    Mouseclick, left, 1100, 535, 2
    Sleep 150
    Send 2
    Sleep 150
    Mouseclick, left, 1100, 650, 2
    Sleep 150
    Send SEA22
    Sleep 150
    MouseClick, left, 1100, 535, 2
return

#IfWinActive