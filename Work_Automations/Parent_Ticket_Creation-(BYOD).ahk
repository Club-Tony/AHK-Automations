#Requires AutoHotkey v1
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
#SingleInstance, Force  ; Reload without prompt when Esc is pressed.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
searchShortcutsPath := A_ScriptDir "\Intra_Desktop_Search_Shortcuts.ahk"

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
        if (!initialCaptured || loopFocus != focusedCtrl)
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
    if (!changed)
    {
        ToolTip, BYOD script timed out waiting for scan
        Sleep 4000
        ToolTip
        ExitApp
    }
    ToolTip  ; clear scan prompt on success

    Sleep 2000  ; allow the scan input to finish
    SendInput, ^f
    WinWait, Search - General,, 3
    if (!ErrorLevel)
    {
        WinActivate, Search - General
        WinWaitActive, Search - General,, 2
        Sleep 400
        running := EnsureSearchShortcutsScript()
        if (running)
        {
            SendInput, !d  ; trigger docksided preset via search shortcuts script
            Sleep 800
        }
        else
        {
            ; Fallback: click Load Saved Search and Docksided preset directly.
            MouseClick, left, 120, 720  ; Load Saved Search
            Sleep 400
            MouseClick, left, 200, 600  ; Docksided preset
            Sleep 400
        }
    }
    ExitApp
Return

EnsureSearchShortcutsScript()
{
    global searchShortcutsPath
    DetectHiddenWindows, On
    running := WinExist("Intra_Desktop_Search_Shortcuts.ahk ahk_class AutoHotkey")
    if (!running && FileExist(searchShortcutsPath))
        Run, %searchShortcutsPath%
    DetectHiddenWindows, Off
    return running || WinExist("Intra_Desktop_Search_Shortcuts.ahk ahk_class AutoHotkey")
}

#IfWinActive
