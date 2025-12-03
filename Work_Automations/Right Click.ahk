#Requires AutoHotkey v1
#NoEnv ; Prevents Unnecessary Environment Variable lookup
#Warn ; Warn All (All Warnings Enabled)
#SingleInstance, Force  ; Reload without prompt when Esc is pressed.
SendMode Input ; Send works as SendInput
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

; Scope: global right-click helper (prefers focused control).
^Esc::Reload

PrintScreen::
    ; Try to right-click the currently focused control (works when the selected item follows keyboard focus).
    ControlGetFocus, focCtrl, A
    if (focCtrl != "")
    {
        ; Anchor to a known good spot when focus resolves to the whole window.
        ControlClick, x265 y95, A,, Right, 1, NA
        return
    }
    ; Fallback: right-click at current mouse position.
    MouseGetPos, curX, curY
    MouseMove, %curX%, %curY%
    Sleep 100
    MouseClick, right
return
