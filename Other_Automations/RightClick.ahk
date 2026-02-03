#Requires AutoHotkey v1
#NoEnv ; Prevents Unnecessary Environment Variable lookup
#Warn ; Warn All (All Warnings Enabled)
#SingleInstance, Force  ; Reload without prompt when Esc is pressed.
SendMode Input ; Send works as SendInput
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

; Functions as a right click shortcut for software/active windows that don't respond with a right click upon menu key input on keyboard.

PrintScreen::
    ; VLC special case: right-click the center of the active VLC window.
    if WinActive("ahk_exe vlc.exe")
    {
        WinGetPos, winX, winY, winW, winH, A
        if (winW > 0 && winH > 0)
        {
            centerX := winX + (winW // 2)
            centerY := winY + (winH // 2)
            MouseMove, %centerX%, %centerY%
            Sleep 100
            MouseClick, right
            return
        }
    }

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
