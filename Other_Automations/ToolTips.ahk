#Requires AutoHotkey v1
#NoEnv
#Warn
#SingleInstance, Force
SendMode Input
SetWorkingDir %A_ScriptDir%
SetTitleMatchMode, 2
CoordMode, Mouse, Screen

TooltipActive := false

^Esc::Reload

#IfWinActive, ahk_exe Code.exe
^!t::
    if (TooltipActive) {
        Gosub, HideTooltips
        Return
    }
    TooltipActive := true
    tooltipText =
    (
VS Code Shortcuts
Ctrl+Alt+P - Command Palette
Ctrl+`` - Toggle terminal
Ctrl+P - Quick Open
Ctrl+Shift+E - Explorer
Ctrl+Shift+F - Search
Ctrl+Shift+G - Source Control
Ctrl+Shift+X - Extensions
Ctrl+K Ctrl+S - Keyboard Shortcuts
Ctrl+Alt+T - Show this tooltip again
    )
    Tooltip, %tooltipText%
    Hotkey, Esc, HideTooltips, On
    SetTimer, HideTooltips, -15000
return
#If

^!t::
    if (TooltipActive) {
        Gosub, HideTooltips
        Return
    }
    TooltipActive := true
    tooltipText =
    (
Global Hotkeys
Ctrl+Shift+Alt+Z - Launch Macros script
Ctrl+Shift+Alt+C - Launch Coord Capture Helper script
Ctrl+Shift+Alt+W - Toggle Window Spy
Win+Alt+E   - Open new Explorer window
-
Window Switch (Personal)
Win+E       - Focus/Minimize/Cycle Explorer
Win+Alt+E   - Open new Explorer window
Win+Alt+V   - Focus/Minimize VS Code
Win+F       - Focus/Minimize/Launch Firefox
Ctrl+Alt+T  - Show this tooltip again
    )
    Tooltip, %tooltipText%
    Hotkey, Esc, HideTooltips, On
    SetTimer, HideTooltips, -15000
return

HideTooltips:
    TooltipActive := false
    Hotkey, Esc, HideTooltips, Off
    Tooltip
return

#If (TooltipActive)
~#e::Gosub HideTooltips
~#!e::Gosub HideTooltips
~#!v::Gosub HideTooltips
~#f::Gosub HideTooltips
~^!t::Gosub HideTooltips
~^+!z::Gosub HideTooltips
~^+!w::Gosub HideTooltips
#If
