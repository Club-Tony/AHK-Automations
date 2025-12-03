#Requires AutoHotkey v1
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Warn  ; Enable warnings to assist with detecting common errors.
#SingleInstance, Force  ; Reload without prompt when Esc is pressed.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetTitleMatchMode, 2  ; Allow partial title matches for Intra windows.

assignTitle := "Intra Desktop Client - Assign Recip"
updateTitle := "Intra Desktop Client - Update"
worldShipTitle := "UPS WorldShip"
qvnTitle := "Quantum View Notify Recipients"

^Esc::Reload

#If (WinActive(assignTitle) || WinActive(updateTitle))
$!Tab::
    ; Only intercept Alt+Tab while in Assign/Update; otherwise let other scripts/OS handle it.
    Send, {Alt up}
    if WinActive(assignTitle)
    {
        if WinExist(updateTitle)
        {
            WinActivate, %updateTitle%
            WinWaitActive, %updateTitle%,, 1
        }
        else
        {
            Send, !{Tab}
        }
    }
    else if WinActive(updateTitle)
    {
        if WinExist(assignTitle)
        {
            WinActivate, %assignTitle%
            WinWaitActive, %assignTitle%,, 1
        }
        else
        {
            Send, !{Tab}
        }
    }
    else
    {
        Send, !{Tab}
    }
return
#If