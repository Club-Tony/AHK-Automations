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
Ctrl+Shift+P - Command Palette
Restarting Extensions - Ctrl+Shift+P - type Restart Extension Host
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
Ctrl+Shift+Alt+C - Launch Coord Capture Helper script
Ctrl+Shift+Alt+W - Toggle Window Spy
Ctrl+Alt+S  - Toggle sound output device
Ctrl+Alt+E  - Reset Explorer tab scaling (reopen tab)
Ctrl+Alt+Shift+E - Multi-tab reset (1-9 tabs)
Win+Alt+E   - Open new Explorer window
-
Window Switch (Personal)
Win+E       - Focus/Minimize/Cycle Explorer
Win+Alt+E   - Open new Explorer window
Win+Alt+V   - Launch/Focus/Minimize VS Code
Win+F       - Focus/Minimize/Launch Firefox
Ctrl+Alt+T  - Show this tooltip again
-
Macros Script Hotkeys
Ctrl+Shift+Alt+Z - Launch Macros script
F5 (Macros menu) - Record macro (kb/mouse + controller)
F6 (Macros menu) - Record controller
L1+L2+R1+R2+A - Start/stop controller record or pause playback
L1+L2+R1+R2+B - Start turbo keyhold
L1+L2+R1+R2+Y - Start pure key hold
L1+L2+R1+R2+X - Kill switch (turn off macros)
Start/Options - Toggle playback pause (controller)
Share/Back - Cancel recording (controller)
Controller map: L1/LB=Left Shoulder, L2/LT=Left Trigger
R1/RB=Right Shoulder, R2/RT=Right Trigger
A/Cross, B/Circle, X/Square, Y/Triangle
F12 - Toggle playback (keyboard)
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
~^!s::Gosub HideTooltips
~^!e::Gosub HideTooltips
~^!+e::Gosub HideTooltips
#If
