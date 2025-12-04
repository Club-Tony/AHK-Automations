#NoEnv
#Warn
SendMode Input
SetWorkingDir %A_ScriptDir%

lastAltATick := 0

!a::
    SendInput, {Alt up}  ; ensure Alt isn’t held so Windows honors the Win press
    Sleep 30
    SendStartMenu()      ; more direct Start/Search opener (Ctrl+Esc)
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

ClearRunHint:
    ToolTip
    lastAltATick := 0
Return

OpenRun:
    Run, explorer.exe shell:::{2559a1f3-21d7-11d4-bdaf-00c04f60b9f0}  ; open Run dialog same as Win+R helper
Return

SendWinTap()
{
    ; Simulate a real Win key press via keybd_event to avoid Alt interference or blocked Send.
    DllCall("keybd_event", "UChar", 0x5B, "UChar", 0, "UInt", 0, "UPtr", 0)      ; Win down
    Sleep 70
    DllCall("keybd_event", "UChar", 0x5B, "UChar", 0, "UInt", 2, "UPtr", 0)      ; Win up
}

SendStartMenu()
{
    ; Ctrl+Esc is a long-standing Start menu shortcut that often succeeds when Win key is blocked.
    SendInput, {Ctrl down}{Esc}{Ctrl up}
}
