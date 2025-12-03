#Requires AutoHotkey v1
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

#r::
    GoSub, OpenRun
Return

OpenRun:
    Run, explorer.exe shell:::{2559a1f3-21d7-11d4-bdaf-00c04f60b9f0}  ; directly open Run dialog
    ToolTip, Win+R Triggered
    Sleep 1500
    ToolTip
Return
