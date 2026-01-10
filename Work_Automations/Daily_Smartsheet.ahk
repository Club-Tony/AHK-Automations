#Requires AutoHotkey v1
#NoEnv ; Prevents Unnecessary Environment Variable lookup
#Warn ; Warn All (All Warnings Enabled)
#SingleInstance, Force  ; Reload without prompt when Esc is pressed.
SendMode Event 
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

SetKeyDelay 50
SetTitleMatchMode, 2
smartsheetWinTitle := "Ouroboros BSC - SEA124"
smartsheetFocusTip := "Smartsheet window inactive, try again"
tabAssistActive := false
tabAssistStep := 0
Esc::ExitApp

^!s:: 
    tabAssistActive := false
    tabAssistStep := 0
    if (!RequireSmartsheetWindow())
        Return
    Mouseclick, left, 1013, 455
    Sleep 100
    Mouseclick, left, 938, 766
    Sleep 100
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
    tabAssistActive := true
    tabAssistStep := 0
Return

#If (tabAssistActive && WinActive(smartsheetWinTitle))
$Tab::
    tabAssistStep += 1
    if (tabAssistStep = 1)
    {
        Send {Tab 3}
    }
    else if (tabAssistStep = 2 || tabAssistStep = 3)
    {
        Send {Tab 2}
    }
    else if (tabAssistStep = 4)
    {
        Send {Tab 7}
        tabAssistActive := false
    }
    else
    {
        Send {Tab}
        tabAssistActive := false
    }
Return

^Tab::
    SendInput, {Ctrl up}{Tab}
Return
#If

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
