#Requires AutoHotkey v1
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
#SingleInstance, Force  ; Reload without prompt when Esc is pressed.
SendMode Input ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

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
; 120, 720 = Load Saved Search button (Search - General)
; 140, 575 = BSC IT Asset Move preset select (Search - General)
; 725, 190 = Enable Parent Item #
; 300, 120 = Print Label button
; 190, 205 = Status field

SetTitleMatchMode, 2 ; allows for partial window title matches
Esc::ExitApp
SetKeyDelay 150
isRunning := false
#IfWinActive, Intra Desktop Client - Assign Recip

CloseScriptIfRunning(name) {
    local pid := ""
    Loop 5
    {
        WinGet, pid, PID, %name%.ahk
        if (!pid)
            WinGet, pid, PID, %name%
        if (!pid)
        {
            Process, Exist, %name%.ahk
            pid := ErrorLevel
        }
        if (pid)
            break
        Sleep 150
    }
    if (pid)
    {
        Process, Close, %pid%
        Process, WaitClose, %pid%, 1
        if (ErrorLevel)
        {
            Process, Close, %pid%
            Process, WaitClose, %pid%, 1
        }
    }
}

!t::
    if (isRunning)
        return
    isRunning := true
    ; Stop the Alt+S assignment helper to avoid hotkey conflicts.
    SetTitleMatchMode, 2
    WinActivate, Intra Desktop Client - Assign Recip
    WinWaitActive, Intra Desktop Client - Assign Recip,, 2
    DetectHiddenWindows, On
    ; Close conflicting scripts launched by Ctrl+Alt+I (Alt+S helper, BYOD, General).
    CloseScriptIfRunning("IT_Requested_IOs-Faster_Assigning")
    CloseScriptIfRunning("Parent_Ticket_Creation-(BYOD)")
    CloseScriptIfRunning("Parent_Ticket_Creation-(GENERAL)")
    DetectHiddenWindows, Off

    ; Attempt parent ticket creation first
    MouseClick, left, 130, 850, 2
    Sleep 250
    Send {space}it-
    Sleep 250
    SendInput, {Enter}
    Sleep 2000
    SendInput, {Down}
    Sleep 250
    MouseClick, left, 1035, 185
    Sleep 250
    MouseClick, left, 1060, 185
    Sleep 250
    MouseMove, 1100, 365
    Sleep 100
    Loop 50
    {
        MouseClick, WheelUp
    }
    MouseClick, left, 1100, 365, 2
    Sleep 200
    SendInput, ^a
    Sleep 80
    SendInput, {Delete}
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
    Sleep 750
    Send, {F5}
    Sleep 3500
    MouseClick, left, 130, 850, 2
    Sleep 250
    Send {space}it-
    Sleep 250
    SendInput, {Enter}
    Sleep 2000
    SendInput, {Down}
    Sleep 250
    MouseMove, 1100, 365
    Sleep 100
    Loop 50
    {
        MouseClick, WheelUp
    }
    MouseClick, left, 1100, 365, 2
    Sleep 200
    SendInput, ^a
    Sleep 80
    SendInput, {Delete}
    Sleep 250
    Send mid
    Sleep 250
    SendInput, {Enter}
    Sleep 250
    MouseClick, left, 300, 120  ; Click print label button
    Sleep 250
    MouseClick, left, 725, 190, 2  ; Enable Parent Item #
    Sleep 250
    ToolTip, Scan parent label to continue script
    changed := false
    initialCaptured := false
    initialText := ""
    focusedCtrl := ""
    Loop 150  ; ~30 seconds total at 200 ms intervals
    {
        Sleep 200
        ControlGetFocus, loopFocus, A
        if (loopFocus = "")
            continue
        if (!initialCaptured)
        {
            ControlGetText, initialText, %loopFocus%, A
            focusedCtrl := loopFocus
            initialCaptured := true
            continue
        }
        ControlGetText, newText, %focusedCtrl%, A
        if (newText != initialText && newText != "")
        {
            changed := true
            break
        }
    }
    ToolTip
    if (!changed)
    {
        isRunning := false
        return
    }
    DetectHiddenWindows, Off
    
    ; Resume rest of script
    Sleep 2000  ; allow the scan input to finish before continuing
    SetTitleMatchMode, 2
    SendInput, ^f
    WinWait, Search - General,, 3
    if (ErrorLevel)
    {
        isRunning := false
        return
    }
    WinActivate, Search - General
    WinWaitActive, Search - General,, 2
    Sleep 750
    MouseClick, left, 120, 720  ; Load Saved Search button (BSC IT Asset Move)
    Sleep 400
    MouseClick, left, 140, 575  ; BSC IT Asset Move preset select
    Sleep 400
    MouseClick, left, 875, 770 ; Click Search
    Sleep 1500
    WinWait, Search Results:, , 10
    if ErrorLevel
    {
        isRunning := false
        Return    
    }
    Sleep 200
    SendInput, !{space}
    Sleep 300
    SendInput, s
    Sleep 300
    SendInput, {Right}
    Sleep 300
    SendInput, {Down}
    MouseMove, 2050, 1025
    Sleep 300
    SendInput, {Enter}
    Sleep 300
    MouseMove, 945, 70
    isRunning := false
Return
#If
