#Requires AutoHotkey v1
#NoEnv
#SingleInstance, Force
SendMode Input
SetWorkingDir %A_ScriptDir%

; Simple helper to capture window-relative mouse coordinates for any active window.
; Usage: left-click each target in order; captures 10 clicks (labeled 1..10). Press Esc to quit.

CoordMode, Mouse, Window

fields := ["1.","2.","3.","4.","5.","6.","7.","8.","9.","10."]
captureFile := A_ScriptDir "\coords-capture.txt"
idx := 1

; start fresh each run
FileDelete, %captureFile%
FileAppend, % "Capture started " A_Now "`n`n", %captureFile%

~LButton::
    if (idx > fields.Length())
        return
    MouseGetPos, x, y
    entry := fields[idx] ": {x: " x ", y: " y "}`n"
    FileAppend, %entry%, %captureFile%
    ToolTip, % "Saved " entry
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
    if (idx <= fields.Length())
    {
        Loop % (fields.Length() - idx + 1)
        {
            entry := fields[idx] ": {x: , y: }`n"
            FileAppend, %entry%, %captureFile%
            idx++
        }
    }
    Sleep 1000
    ToolTip, Done. Saved to %captureFile%
    Sleep 2000
    ToolTip
    ExitApp
