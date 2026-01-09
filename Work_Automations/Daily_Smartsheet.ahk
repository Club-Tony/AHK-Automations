#Requires AutoHotkey v1
#NoEnv ; Prevents Unnecessary Environment Variable lookup
#Warn ; Warn All (All Warnings Enabled)
#SingleInstance, Force  ; Reload without prompt when Esc is pressed.
SendMode Input ; Overrided by SendMode Event below
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

SendMode, Event ; Since target app is ignoring SendMode, Input
SetKeyDelay 50
SetTitleMatchMode, 2
smartsheetWinTitle := "Ouroboros BSC - SEA124"
smartsheetFocusTip := "Smartsheet window inactive, try again"
Esc::ExitApp

^!s:: 
    if (!RequireSmartsheetWindow())
        Return
    Send {Tab 2}
    Send {Space}
    Send {Tab 2}
    Send {Space}
    Send {Tab}
    Send daveyuan
    Send {Tab}
    Send bsc
    Sleep 100 
    if (!RequireSmartsheetWindow())
        Return
    Send {Enter down}
    Sleep 50
    Send {Enter up}
    Sleep 100
    if (!RequireSmartsheetWindow())
        Return
    Send {Tab 16}
    Send {Space}
    Send {Tab}
    Send ouroboros-bsc@amazon.com
    Send +{Tab 13}
Return

RequireSmartsheetWindow()
{
    global smartsheetWinTitle, smartsheetFocusTip
    if WinActive(smartsheetWinTitle)
        return true
    ToolTip, %smartsheetFocusTip%
    SetTimer, ClearSmartsheetTip, -5000
    return false
}

ClearSmartsheetTip:
    ToolTip
Return
