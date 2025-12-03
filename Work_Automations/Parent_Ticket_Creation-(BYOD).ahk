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
; 1100, 650 = BSC Location (Destination)
; 500, 1025 = Select Preset
; 350, 900 = Docksided items preset select 
; 725, 190 = Enable Parent Item #
; 500, 120 = Print Label button

Esc::ExitApp
#IfWinActive, Intra Desktop Client - Assign Recip
SetKeyDelay 150

!p::
    ; First close conflicting scripts
    SetTitleMatchMode, 2
    DetectHiddenWindows, On
    WinGet, myPID, PID, Parent_Ticket_Creation-(GENERAL)
    if myPID ; Only close if a PID was found
        Process, Close, %myPID%
    Sleep 50
    WinGet, myPID, PID, IT_Requested_IOs-Faster_Assigning
    if myPID ; Only close if a PID was found
        Process, Close, %myPID%
    Sleep 250
    SetTitleMatchMode, 1
    Sleep 100
    DetectHiddenWindows, Off
    Sleep 250

    ; Rest of the script
    MouseClick, left, 130, 850, 2
    Sleep 250
    Send byod
    Sleep 250
    SendInput, {Enter}
    Sleep 2000
    SendInput, {Down}
    Sleep 250
    MouseClick, left, 1035, 185
    Sleep 250
    MouseClick, left, 1060, 185
    Sleep 250
    MouseClick, left, 1100, 365, 2
    Sleep 250
    Send mid
    Sleep 250
    SendInput, {Enter}
    Sleep 250
    MouseClick, left, 1100, 650, 2
    Sleep 250
    Send SEA124
    Sleep 250
    MouseClick, left, 200, 245, 2
    Sleep 250
    SendInput, ^n
    Sleep 1000
    Send, {F5}
    Sleep 5000
    MouseClick, left, 130, 850, 2
    Sleep 250
    Send byod
    Sleep 250
    SendInput, {Enter}
    Sleep 2000
    SendInput, {Down}
    Sleep 250
    MouseClick, left, 1100, 365, 2
    Sleep 250
    Send mid
    Sleep 250
    SendInput, {Enter}
    Sleep 250
    MouseClick, left, 1100, 650, 2
    Sleep 250
    Send SEA124
    Sleep 250
    MouseClick, left, 300, 120
    Sleep 250
    MouseClick, left, 725, 190, 2
    Sleep 250
    ExitApp
Return

#IfWinActive
