#Requires AutoHotkey v1
#NoEnv
#Warn
#SingleInstance, Force
SendMode Input
SetWorkingDir %A_ScriptDir%

SetTitleMatchMode, 2
TooltipActive := false
TooltipLocked := false
TooltipCooldownMs := 30000 ; 30 seconds
ButtonsTooltipActive := false

^Esc::Reload

#If ( WinActive("Intra Desktop Client - Assign Recip")
    || WinActive("Intra Desktop Client - Update") )

^!t::
    if (TooltipActive) {
        Gosub, HideTooltips
        Return
    }
    if (TooltipLocked) {
        Return
    }
    TooltipLocked := true
    SetTimer, UnlockTooltip, % -TooltipCooldownMs
    TooltipActive := true
    TooltipText =
    (
SSJ-Intra Hotkeys
-
Assign Recip Tab:
Alt+I - Yellow Pouch to SEA22
Ctrl+I - Yellow Pouch (multi-piece)
Alt+R - Clear + Toggle Print Button & Normalize
Alt+C - Clear All
Alt+E - Focus Scan Field
Alt+A - Focus Alias Field
Alt+N - Focus Name Field
Alt+1 - Focus Package Type
Alt+2 - Focus BSC Location
Alt+3 - Print Label Toggle
Alt+D - Item Var Lookup + Apply All Toggle
-
Update Tab:
Alt+P - Status Select -> Pickup from BSC
Alt+D - Status Select -> Delivery
Alt+V - Status Select -> Void
Alt+S - Click Status Field
Ctrl+Alt+T - Show this tooltip again
-
Additional Scripts Launch Hotkeys:
Ctrl+Alt+F - Launch Intra Search Shortcuts
Ctrl+Alt+I - Launch Intra Extensive Automations
Ctrl+Alt+C - Launch DSRF to WorldShip Script
Ctrl+Alt+L - Launch Daily Audit + Smartsheet
    )
    Tooltip, %TooltipText%
    Hotkey, Esc, HideTooltips, On
    SetTimer, HideTooltips, -30000
Return

#If TooltipActive
~Esc::Gosub HideTooltips
~!s::Gosub HideTooltips
~!p::Gosub HideTooltips
~^!p::Gosub HideTooltips
~!i::Gosub HideTooltips
~!t::Gosub HideTooltips
~!r::Gosub HideTooltips
~!Space::Gosub HideTooltips
~!e::Gosub HideTooltips
~!a::Gosub HideTooltips
~!n::Gosub HideTooltips
~!1::Gosub HideTooltips
~!2::Gosub HideTooltips
~!3::Gosub HideTooltips
~!d::Gosub HideTooltips
~!c::Gosub HideTooltips
~^!c::Gosub HideTooltips
~^!f::Gosub HideTooltips
~^!i::Gosub HideTooltips
~^i::Gosub HideTooltips
#If

UnlockTooltip:
    TooltipLocked := false
Return

; Intra Buttons (Firefox Interoffice/Shipping) tooltip
#IfWinActive, Intra: Interoffice Request ahk_exe firefox.exe

^!t::
    ButtonsTooltipActive := true
    TooltipActive := true
    tooltipText =
    (
SSJ Intra - Interoffice Requests
Alt+S  - Focus envelope icon
Alt+A  - Focus alias field
Ctrl+Enter - Scroll to bottom and Submit
Ctrl+Alt+S - Fill Special Instructions
Ctrl+Alt+A - ACP preset -> Alias
Alt+P  - Load "posters" preset -> Name field
Alt+1  - Focus "# of Packages"
Alt+2  - Focus Package Type
Alt+L  - Click Load button
Alt+C  - Clear/Reset
    )
    Tooltip, %tooltipText%
    SetTimer, HideTooltips, -15000
Return

#If (ButtonsTooltipActive)
~Esc::Gosub HideTooltips
~!s::Gosub HideTooltips
~!e::Gosub HideTooltips
~!a::Gosub HideTooltips
~!p::Gosub HideTooltips
~^!a::Gosub HideTooltips
~^!s::Gosub HideTooltips
~!1::Gosub HideTooltips
~!2::Gosub HideTooltips
~!l::Gosub HideTooltips
~!c::Gosub HideTooltips
~!Space::Gosub HideTooltips
~^Enter::Gosub HideTooltips
#If

#IfWinActive, Intra: Shipping Request Form ahk_exe firefox.exe
^!t::
    if (ButtonsTooltipActive) {
        Gosub, HideTooltips
        Return
    }
    ButtonsTooltipActive := true
    TooltipActive := true
    tooltipText =
    (
Launch Hotkeys:
Ctrl+Alt+C: Launch DSRF to WorldShip Script
Ctrl+Alt+P: Personal Form
Ctrl+Alt+B: Business Form
    )
    Tooltip, %tooltipText%
    SetTimer, HideTooltips, -15000
return

#If (ButtonsTooltipActive)
~^!c::Gosub HideTooltips
~^!p::Gosub HideTooltips
~^!b::Gosub HideTooltips
#If

HideTooltips:
    SetTimer, UnlockTooltip, Off
    TooltipLocked := false
    TooltipActive := false
    ButtonsTooltipActive := false
    Hotkey, Esc, HideTooltips, Off
    Tooltip
Return
