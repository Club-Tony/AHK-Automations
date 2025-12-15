#Requires AutoHotkey v1
#NoEnv
#SingleInstance, Force
SendMode Input
SetWorkingDir %A_ScriptDir%

; Simple helper to capture window-relative mouse coordinates for any active window.
; Usage: left-click each target in order; captures 10 clicks (labeled 1..10). Press Esc to quit.

CoordMode, Mouse, Window

fields := ["1.","2.","3.","4.","5.","6.","7.","8.","9.","10."]
captureFile := A_ScriptDir "\coord.txt"
captureHeader := "Capture started " A_Now
entries := []
idx := 1

InitializeCapture(entries, captureFile, captureHeader, fields)

~LButton::
    if (idx > fields.Length())
        return
    MouseGetPos, x, y
    entries[idx] := fields[idx] ": {x: " x ", y: " y "}"
    WriteCapture(entries, captureFile, captureHeader)
    ToolTip, % "Saved " entries[idx]
    SetTimer, ClearTip, -1000
    idx++
    if (idx > fields.Length())
    {
        Sleep 1000
        ToolTip, Done. Saved to %captureFile%
        SetTimer, ClearTip, -2000
    }
return

ClearTip:
    ToolTip
return

Esc::
    WriteCapture(entries, captureFile, captureHeader)
    Sleep 1000
    ToolTip, Done. Saved to %captureFile%
    Sleep 2000
    ToolTip
    ExitApp

InitializeCapture(entries, captureFile, captureHeader, fields)
{
    if (entries.Length())
        entries.RemoveAt(1, entries.Length())  ; Clear any previous values when script reloads.
    Loop % fields.Length()
        entries[A_Index] := fields[A_Index] ": {x: , y: }"
    WriteCapture(entries, captureFile, captureHeader)
}

WriteCapture(entries, captureFile, captureHeader)
{
    FileDelete, %captureFile%
    FileAppend, % captureHeader "`n`n", %captureFile%
    Loop % entries.Length()
        FileAppend, % entries[A_Index] "`n", %captureFile%
}
