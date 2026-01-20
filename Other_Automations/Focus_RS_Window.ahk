#Requires AutoHotkey v1
#NoEnv ; Prevents Unnecessary Environment Variable lookup
#Warn ; Warn All (All Warnings Enabled)
#SingleInstance, Force ; Removes script already open warning when reloading scripts
SendMode Input
SetWorkingDir, %A_ScriptDir%

Esc::ExitApp

!z::
    SetTitleMatchMode, 2
    WinActivate, RuneScape
    WinWaitActive, RuneScape, , 1
Return
