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

#If AssignRecipHotkeyAllowed()
!s::
    FocusAssignRecipWindow()
    if (!WinActive("Intra Desktop Client - Assign Recip"))
        Return
    CloseParentTicketGeneral()
    CloseParentTicketBYOD()
    Sleep 250
    MouseClick, left, 930, 830, 2
    Sleep 500
    MouseClick, left, 925, 850
Return

#IfWinExist, Intra Desktop Client - Assign Recip
Enter::
    SendInput, {enter}
    Sleep 250
    Send, {Down}
    Sleep 250
    MouseClick, left, 200, 245, 2

    ; insert text detection here, where upon text detected under mouse, wait for it to finish and then send so it mimicks scanner carriage return
    MouseGetPos, , , winID, ctrlUnderMouse
    ControlGetText, oldText, %ctrlUnderMouse%, ahk_id %winID%

    Loop {
        Sleep 100
        ControlGetText, newText, %ctrlUnderMouse%, ahk_id %winID%
        if (newText != oldText && newText != "") {
            Sleep 100
            ControlSend, %ctrlUnderMouse%, {Enter}, ahk_id %winID%
            Sleep 100
            Send, {F5}
            ; Wait up to ~6s for Intra to clear the scan field after refresh instead of fixed 2s.
            Loop 60 {
                Sleep 100
                ControlGetFocus, focusedCtrl, ahk_id %winID%
                if (focusedCtrl = "")
                    continue
                ControlGetText, postF5Text, %focusedCtrl%, ahk_id %winID%
                if (postF5Text = "")
                    Break
            }
            Sleep 250
            Gosub, !s
            Break
        }
    }
Return

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
