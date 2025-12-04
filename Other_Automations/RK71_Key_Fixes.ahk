#Requires AutoHotkey v1
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

lastAltATick := 0

#r::
    GoSub, OpenRun
Return

!a::
    SendInput, {Alt up}          ; release Alt so Start honors the key
    Sleep 30
    SendStartMenu()              ; try Start via Ctrl+Esc + Win tap
    lastAltATick := A_TickCount
    ToolTip, Press Alt+R within 5s to open Run dialog
    SetTimer, ClearRunHint, -5000
Return

!r::
    if (lastAltATick && (A_TickCount - lastAltATick <= 5000))
    {
        GoSub, OpenRun
        ToolTip
        lastAltATick := 0
    }
Return

OpenRun:
    Run, explorer.exe shell:::{2559a1f3-21d7-11d4-bdaf-00c04f60b9f0}  ; directly open Run dialog
    ToolTip, Win+R Triggered
    Sleep 1500
    ToolTip
Return

ClearRunHint:
    ToolTip
    lastAltATick := 0
Return

SendStartMenu()
{
    ; Use Ctrl+Esc plus a low-level Win tap to handle cases where the Win key is blocked.
    SendInput, {Ctrl down}{Esc}{Ctrl up}
    Sleep 60
    DllCall("keybd_event", "UChar", 0x5B, "UChar", 0, "UInt", 0, "UPtr", 0)      ; Win down
    Sleep 60
    DllCall("keybd_event", "UChar", 0x5B, "UChar", 0, "UInt", 2, "UPtr", 0)      ; Win up
}
