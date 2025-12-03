#Requires AutoHotkey v1
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
; Toggle behavior
; - Alt+S enables or disables using `/` for a single left mouse click.
; - When enabled, pressing `/` sends exactly one click, even if the key is held.

ClickMode := false

; Prepare / hotkey (disabled until the toggle is turned on)
Hotkey, $/, SlashClick, Off

!s::
    ClickMode := !ClickMode
    if (ClickMode) {
        Hotkey, $/, SlashClick, On
        ToolTip, Slash => Single Left Click`nPress Alt+S to disable
    } else {
        Hotkey, $/, SlashClick, Off
        ToolTip, Slash restored to normal
    }
    SetTimer, RemoveToolTip, -3000
Return

SlashClick:
    if (!ClickMode)
        Return
    MouseClick, Left
    ToolTip, Left Click! Alt+S to disable
    SetTimer, RemoveToolTip, -500
    KeyWait, /
Return

RemoveToolTip:
    ToolTip
Return