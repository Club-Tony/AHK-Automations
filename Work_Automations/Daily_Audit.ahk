#Requires AutoHotkey v1
#NoEnv ; Prevents Unnecessary Environment Variable lookup
#Warn ; Warn All (All Warnings Enabled)
#SingleInstance, Force  ; Reload without prompt when Esc is pressed.
SendMode Event 
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

SetKeyDelay 50
SetTitleMatchMode, 2
CoordMode, Mouse, Window
auditWinTitle := "Daily BSC Audit"
auditFocusTip := "Audit window inactive, try again"
; Expected minimum window dimensions for coordinate clicks to be valid
expectedMinWidth := 1400
expectedMinHeight := 600
Esc::ExitApp

#If WinActive(auditWinTitle)
^+!d::
    if (!RequireAuditWindow())
        Return
    if (!CheckWindowGeometry())
        Return
    Mouseclick, left, 1392, 230
    Sleep 100
    Send {Tab 2}
    Send {Space}
    Sleep 100
    Send {Tab} 
    Send {Space}
    Send {Tab} 
    Send {Down 3}
    Sleep 100 
    if (!RequireAuditWindow())
        Return

    Send {Enter down}
    Sleep 50
    Send {Enter up}
    Sleep 100
    if (!RequireAuditWindow())
        Return

    Send {Tab 2} 
    Send puget
    Sleep 100
    Send {Enter} 
    Send {Tab 2} 
    Send 124
    Sleep 100
    if (!RequireAuditWindow())
        Return

    Send {Enter down}
    Sleep 50
    Send {Enter up}
    Sleep 100
    if (!RequireAuditWindow())
        Return
    
    Loop 4 {
        if (!RequireAuditWindow())
            Return
        Send {Tab 2} 
        Send {Down} 
        Send {Enter}
    }
    
    if (!RequireAuditWindow())
        Return
    Send {Tab 2} 
    Send {Down 2} 
    Send {Enter}
    
    Loop 3 {
        if (!RequireAuditWindow())
            Return
        Send {Tab 2} 
        Send {Down} 
        Send {Enter}
    }
    
    if (!RequireAuditWindow())
        Return
    Send {Tab 2}
    Send Anthony Davey
    Sleep 100  ; Allow name field to process
    if (!RequireAuditWindow())
        Return
    Send {Tab}
    Send {Space}
    Send {Tab}
    Send ouroboros-bsc@amazon.com
    Sleep 150  ; Allow email field to process
    if (!RequireAuditWindow())
        Return
    Send {Tab}
    ShowCompletionTip("Daily Audit form completed")
Return
#If

RequireAuditWindow()
{
    global auditWinTitle, auditFocusTip
    if WinActive(auditWinTitle)
        return true
    ToolTip, %auditFocusTip%
    SetTimer, ClearAuditTip, -5000
    return false
}

ClearAuditTip:
    ToolTip
Return

CheckWindowGeometry()
{
    global auditWinTitle, expectedMinWidth, expectedMinHeight
    WinGetPos, winX, winY, winW, winH, %auditWinTitle%
    if (winW < expectedMinWidth || winH < expectedMinHeight)
    {
        ToolTip, Window too small (%winW%x%winH%). Expected at least %expectedMinWidth%x%expectedMinHeight%
        SetTimer, ClearAuditTip, -5000
        return false
    }
    return true
}

ShowCompletionTip(msg)
{
    ToolTip, %msg%
    SetTimer, ClearAuditTip, -4000
}
