#Requires AutoHotkey v1
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
#SingleInstance, Force  ; Reload without prompt when Esc is pressed.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

; 945, 850 = coordinates for alias field
; 200, 245 = coordinates for Scan field
; 930, 830 = coordinates for alias field filter dropdown
; 925, 850 = coordinates for alias field filter dropdown: 'starts with' selection
; Window Position: x: -7 y: 0 w: 1322 h: 1339

; Script Function: if window active, alt+s selects 'start with' delimiter, then alias field, at which point you type, upon hitting
; -enter, triggers hit down arrowkey, then clicks Scan field. Upon scan, sends F5 for submission and repeats alt+s function again
; alt+s is hotkey for script 
; Esc = End alt+s loop

Esc::ExitApp
SetKeyDelay 100
scanReady := false
lastActivity := 0
timeoutActive := false
timeoutShowing := false
timeoutMouseX := 0
timeoutMouseY := 0

#If AssignRecipHotkeyAllowed()
!s::
    BumpAssignActivity()
    StartAssignTimeout()
    FocusAssignRecipWindow()
    if (!WinActive("Intra Desktop Client - Assign Recip"))
        Return
    CloseParentTicketGeneral()
    CloseParentTicketBYOD()
    Sleep 50
    MouseClick, left, 930, 830, 2
    Sleep 100
    MouseClick, left, 925, 850
Return

#IfWinExist, Intra Desktop Client - Assign Recip
Enter::
    BumpAssignActivity()
    SendInput, {enter}
    Sleep 400
    Send, {Down}
    Sleep 400
    MouseClick, left, 200, 245, 2

    ; insert text detection here, where upon text detected under mouse, wait for it to finish and then send so it mimicks scanner carriage return
    MouseGetPos, , , winID, ctrlUnderMouse
    if (ctrlUnderMouse = "")
        Return
    ControlGetText, oldText, %ctrlUnderMouse%, ahk_id %winID%
    lastText := oldText
    stableCount := 0

    Loop, 80 {  ; up to ~8s
        Sleep 100
        ControlGetText, newText, %ctrlUnderMouse%, ahk_id %winID%
        if (newText != "" && newText != lastText) {
            lastText := newText
            stableCount := 1
        } else if (newText != "" && newText = lastText) {
            stableCount++
        }
        if (stableCount >= 3) {
            Sleep 150
            ControlSend, %ctrlUnderMouse%, {Enter}, ahk_id %winID%
            Sleep 150
            scanReady := true
            if (scanReady) {
                scanReady := false
                Sleep 250
                Send, {F5}
                Sleep 250
                Gosub, PostF5Recovery
            }
            Break
        }
    }
Return

F5::
    BumpAssignActivity()
    if (!AssignRecipHotkeyAllowed())
        return
    if (!scanReady)
        return
    scanReady := false
    Send, {F5}
    Gosub, PostF5Recovery
return

PostF5Recovery:
    ; Wait up to ~6s for Intra to clear the scan field after refresh instead of fixed 2s.
    WinGet, winIDAfter, ID, Intra Desktop Client - Assign Recip
    Loop 60 {
        Sleep 100
        ControlGetFocus, focusedCtrl, ahk_id %winIDAfter%
        if (focusedCtrl = "")
            continue
        ControlGetText, postF5Text, %focusedCtrl%, ahk_id %winIDAfter%
        if (postF5Text = "")
            Break
    }
    ; Wait for Assign Recip to regain active status after the brief loading overlay, then refocus the alias field.
    WinWaitActive, Intra Desktop Client - Assign Recip,, 3
    FocusAssignRecipWindow()
    MouseClick, left, 945, 850
    Sleep 100
    Sleep 150
    Gosub, !s
return

CheckAssignTimeout:
    if (!timeoutActive)
        return
    elapsed := A_TickCount - lastActivity
    if (elapsed >= 30000) {
        timeoutActive := false
        timeoutShowing := true
        MouseGetPos, timeoutMouseX, timeoutMouseY
        Tooltip, Alt+S Faster Assigning Script Timeout
        SetTimer, TimeoutMouseCheck, 100
        SetTimer, TimeoutExitSub, -3000
    }
return

TimeoutExitSub:
    SetTimer, TimeoutMouseCheck, Off
    Tooltip
    ExitApp
return

BumpAssignActivity() {
    global lastActivity
    lastActivity := A_TickCount
}

StartAssignTimeout() {
    global timeoutActive
    if (!timeoutActive) {
        SetTimer, CheckAssignTimeout, 500
        timeoutActive := true
    }
    BumpAssignActivity()
}

#If (timeoutShowing)
~Esc::Gosub TimeoutExitSub
#If

TimeoutMouseCheck:
    MouseGetPos, curX, curY
    if (curX != timeoutMouseX || curY != timeoutMouseY)
        Gosub, TimeoutExitSub
return

#IfWinActive

CloseParentTicketGeneral() {
    ; Ensure the general parent ticket script is not running because it registers its own Enter hotkeys.
    SetTitleMatchMode, 2
    DetectHiddenWindows, On
    WinGet, conflictingPID, PID, Parent_Ticket_Creation-(GENERAL)
    if (conflictingPID)
        Process, Close, %conflictingPID%
    DetectHiddenWindows, Off
    SetTitleMatchMode, 1
}

CloseParentTicketBYOD() {
    ; Ensure the BYOD parent ticket script is not running because it registers its own Enter hotkeys.
    SetTitleMatchMode, 2
    DetectHiddenWindows, On
    WinGet, conflictingPID, PID, Parent_Ticket_Creation-(BYOD)
    if (conflictingPID)
        Process, Close, %conflictingPID%
    DetectHiddenWindows, Off
    SetTitleMatchMode, 1
}

FocusAssignRecipWindow() {
    ; Bring the Intra Assign Recip window forward even if another window is active.
    SetTitleMatchMode, 2
    WinActivate, Intra Desktop Client - Assign Recip
    WinWaitActive, Intra Desktop Client - Assign Recip,, 2
    SetTitleMatchMode, 1
}

AssignRecipHotkeyAllowed() {
    ; Enable Alt+S only when Assign Recip exists and we're not in conflicting Intra windows.
    SetTitleMatchMode, 2
    allowed := WinExist("Intra Desktop Client - Assign Recip")
        && !WinActive("Intra Desktop Client - Update")
        && !WinActive("Search - General")
        && !WinActive("Search Results:")
        && !WinActive("Item Details")
    SetTitleMatchMode, 1
    return allowed
}
