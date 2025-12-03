#Requires AutoHotkey v1
#NoEnv ; Prevents Unnecessary Environment Variable lookup
#Warn ; Warn All (All Warnings Enabled)
#SingleInstance, Force  ; Reload without prompt when Esc is pressed.
SendMode Input ; Overrided by SendMode Event below
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

SendMode, Event ; Since target app is ignoring SendMode, Input
SetKeyDelay 50
Esc::ExitApp

^!s:: 
    Send {Tab 2}
    Send {Space}
    Send {Tab 2}
    Send {Space}
    Send {Tab}
    Send daveyuan
    Send {Tab}
    Send bsc
    Sleep 100 
    Send {Enter down}
    Sleep 50
    Send {Enter up}
    Sleep 100
    Send {Tab 16}
    Send {Space}
    Send {Tab}
    Send ouroboros-bsc@amazon.com
    Send +{Tab 13}
Return
