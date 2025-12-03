#Requires AutoHotkey v1
#NoEnv ; Prevents Unnecessary Environment Variable lookup
#Warn ; Warn All (All Warnings Enabled)
#SingleInstance, Force  ; Reload without prompt when Esc is pressed.
SendMode Input ; Overrided by SendMode Event below
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

SendMode Event ; Since target app is ignoring SendMode, Input
SetKeyDelay 50
Esc::ExitApp
^!d:: 
    Send {Tab 3}  
    Send {Space}
    Send {Tab 2} 
    Send {Space}
    Send {Tab} 
    Send {Down 3}
    Sleep 100 

    Send {Enter down}
    Sleep 50
    Send {Enter up}
    Sleep 100

    Send {Tab 2} 
    Send puget
    Sleep 100
    Send {Enter} 
    Send {Tab 2} 
    Send 124
    Sleep 100

    Send {Enter down}
    Sleep 50
    Send {Enter up}
    Sleep 100
    
    Loop 4 {
        Send {Tab 2} 
        Send {Down} 
        Send {Enter}
    }
    
    Send {Tab 2} 
    Send {Down 2} 
    Send {Enter}
    
    Loop 3 {
        Send {Tab 2} 
        Send {Down} 
        Send {Enter}
    }
    
    Send {Tab 2} 
    Send Anthony Davey
    Send {Tab} 
    Send {Space} 
    Send {Tab} 
    Send ouroboros-bsc@amazon.com
    Send {Tab}
Return
