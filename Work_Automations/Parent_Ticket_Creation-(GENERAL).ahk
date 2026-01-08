#Requires AutoHotkey v1
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
#SingleInstance, Force  ; Reload without prompt when Esc is pressed.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
bscLocationWait := ""
endWaitActive := false

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

^!p::
    SetTitleMatchMode, 2
    DetectHiddenWindows, On
    WinGet, myPID, PID, Parent_Ticket_Creation-(BYOD)
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

    ; Start with toggle selection/adjusting variables
    MouseClick, left, 1035, 185
    Sleep 250
    MouseClick, left, 1060, 185
    Sleep 250

    ; Click alias/starts with delimiter, wait for enter input
    MouseClick, left, 930, 830
    Sleep 250
    MouseClick, left, 925, 850, 2
    Tooltip, Type alias and press Enter to continue script

Return

; wait for enter input after typing alias
Enter::
    if (endWaitActive)
        Return
    if (bscLocationWait)
        Return
    Tooltip
    SendInput, {enter}
    WinWaitActive, Intra Desktop Client - Assign Recip,, 3
    Sleep 150
    Send, {Down}
    Sleep 250
    MouseClick, left, 1100, 365, 2
    Sleep 250
    Send mid
    Sleep 250
    SendInput, {enter}
    MouseClick, left, 1100, 650, 2
    Sleep 250
    bscLocationWait := "enter"
    Tooltip, Type BSC Location & press Ctrl+Enter to continue script
Return

; continue after typing BSC location from Enter hotkey
ContinueAfterBscEnter:
    bscLocationWait := ""
    Tooltip
    Sleep 250
    MouseClick, left, 200, 245, 2
    Sleep 250
    SendInput, ^n
    Sleep 1000
    Send, {F5}
    Sleep 5000

    ; Click alias/starts with delimiter, wait for enter input
    MouseMove, 930, 830
    Sleep 250
    MouseClick, left
    Sleep 250
    MouseMove, 925, 850
    Sleep 250
    MouseClick, left, , , 2
    Tooltip, Type alias & press Ctrl+Enter to continue script
Return

; continue after typing BSC location from Ctrl+Enter hotkey
ContinueAfterBscCtrl:
    bscLocationWait := ""
    Tooltip
    Sleep 250
    MouseClick, left, 300, 120
    Sleep 250
    MouseClick, left, 725, 190, 2
    Sleep 250
    endWaitActive := true
    Tooltip, Scan into Parent field - Script Done
    MouseGetPos, startX, startY
    lastIdle := A_TimeIdlePhysical
    Loop 300  ; ~60 seconds total at 200 ms intervals
    {
        Sleep 200
        if (GetKeyState("Esc", "P"))
            break
        currentIdle := A_TimeIdlePhysical
        if (currentIdle < lastIdle)
            break
        lastIdle := currentIdle
        MouseGetPos, curX, curY
        if (curX != startX || curY != startY)
            break
    }
    endWaitActive := false
    Tooltip
    ExitApp
Return

; wait for enter input after typing alias
^Enter::
    if (endWaitActive)
        Return
    if (bscLocationWait = "enter")
    {
        Gosub, ContinueAfterBscEnter
        Return
    }
    if (bscLocationWait = "ctrl")
    {
        Gosub, ContinueAfterBscCtrl
        Return
    }
    Tooltip
    SendInput, {Enter}
    Sleep 250
    SendInput, {Down}
    Sleep 250
    MouseClick, left, 1100, 365, 2
    Sleep 250
    Send mid
    Sleep 250
    SendInput, {enter}
    MouseClick, left, 1100, 650, 2
    Sleep 250
    bscLocationWait := "ctrl"
    Tooltip, Type BSC Location & press Ctrl+Enter to continue script
Return

#IfWinActive



