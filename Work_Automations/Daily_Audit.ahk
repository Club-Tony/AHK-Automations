#Requires AutoHotkey v1
#NoEnv ; Prevents Unnecessary Environment Variable lookup
#Warn ; Warn All (All Warnings Enabled)
#SingleInstance, Force  ; Reload without prompt when Esc is pressed.
SendMode Input ; Overrided by SendMode Event below
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

SendMode Event ; Since target app is ignoring SendMode, Input
SetKeyDelay 50
SetTitleMatchMode, 2
auditWinTitle := "Daily BSC Audit"
auditFocusTip := "Smartsheet window inactive, try again"
Esc::ExitApp
^!d:: 
    if (!RequireAuditWindow())
        Return
    Send {Tab 3}  
    Send {Space}
    Send {Tab 2} 
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
    Send {Tab} 
    Send {Space} 
    Send {Tab} 
    Send ouroboros-bsc@amazon.com
    Send {Tab}
Return

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
