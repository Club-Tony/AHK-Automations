#Requires AutoHotkey v1
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
#SingleInstance, Force  ; Reload without prompt when Esc is pressed.
SendMode Input ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

^Esc::Reload

; Window Coordinates (Intra Desktop Client - Assign Recip):
; Window Position: x: -7 y: 0 w: 1322 h: 1339
; 200, 245 = Scan field
; 130, 850 = Name field
; 1035, 185 = enable item var lookup button
; 1060, 185 = applying to all item vars button
; 1100, 365 = Package type
; 1100, 650 = BSC Location (Destination)
; 300, 1025 = Select Preset
; 350, 900 = Docksided items preset select 
; 725, 190 = Enable Parent Item #
; 300, 120 = Print Label button

#IfWinActive, Intra Desktop Client - Assign Recip
SetKeyDelay 150

!r::
    MouseClick, left, 1035, 185
    Sleep 300
    MouseClick, left, 1060, 185
    Sleep 300
    MouseClick, left, 300, 120
    Sleep 300
    MouseClick, left, 70, 1345, 2
    Sleep 250
    SendInput, {enter}
    Sleep 100

    MouseMove, 1100, 365
    Sleep 100
    Loop 50
    {
        MouseClick, WheelUp
    }
    MouseClick, left, 40, 1300, 2
    Sleep 100
    MouseClick, left, 200, 245, 2
Return

!3:: ; click print label button
    MouseClick, left, 300, 120
    Sleep 150
    MouseClick, left, 200, 245, 2
Return

#IfWinActive
