#Requires AutoHotkey v1
#NoEnv ; Prevents Unnecessary Environment Variable lookup
#Warn ; Warn All (All Warnings Enabled)
#SingleInstance, Force  ; Reload without prompt when Esc is pressed.
SendMode Event 
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

SetKeyDelay 50
SetTitleMatchMode, 2
CoordMode, Mouse, Window
smartsheetWinTitle := "Ouroboros BSC - SEA124"
smartsheetFocusTip := "Smartsheet window inactive, try again"
; Expected minimum window dimensions for coordinate clicks to be valid
expectedMinWidth := 1100
expectedMinHeight := 800
tabAssistActive := false
tabAssistStep := 0

; Inter-script trigger (used by Launcher_Script full-auto).
dailyMsgId := 0x5558
OnMessage(dailyMsgId, "HandleDailySmartsheetMessage")

Esc::ExitApp

^+!s::
    DoDailySmartsheet()
Return

HandleDailySmartsheetMessage(wParam, lParam, msg, hwnd)
{
    global dailyMsgId
    if (msg != dailyMsgId)
        return
    if (wParam = 1)
        DoDailySmartsheet()
}

DoDailySmartsheet()
{
    global tabAssistActive, tabAssistStep
    tabAssistActive := false
    tabAssistStep := 0
    if (!RequireSmartsheetWindow())
        return
    if (!CheckWindowGeometry())
        return
    Mouseclick, left, 1013, 455
    Sleep 100
    Send {Tab 2}
    Send {Space}
    Sleep 100
    Send {Tab}
    Send daveyuan
    Send {Tab}
    Send bsc
    Sleep 100
    if (!RequireSmartsheetWindow())
        return
    Send {Enter down}
    Sleep 50
    Send {Enter up}
    Sleep 100
    if (!RequireSmartsheetWindow())
        return
    Send {Tab 16}
    Send {Space}
    Send {Tab}
    Send ouroboros-bsc@amazon.com
    Sleep 150  ; Allow email field to process
    if (!RequireSmartsheetWindow())
        return
    Send +{Tab 13}
    tabAssistActive := true
    tabAssistStep := 0
    ShowCompletionTip("Smartsheet form ready - Tab Assist active")
}

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

CheckWindowGeometry()
{
    global smartsheetWinTitle, expectedMinWidth, expectedMinHeight
    WinGetPos, winX, winY, winW, winH, %smartsheetWinTitle%
    if (winW < expectedMinWidth || winH < expectedMinHeight)
    {
        ToolTip, Window too small (%winW%x%winH%). Expected at least %expectedMinWidth%x%expectedMinHeight%
        SetTimer, ClearSmartsheetTip, -5000
        return false
    }
    return true
}

ShowCompletionTip(msg)
{
    ToolTip, %msg%
    SetTimer, ClearSmartsheetTip, -4000
}
