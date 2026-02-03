#Requires AutoHotkey v1
#NoEnv
#Warn
#SingleInstance, Force
#InstallKeybdHook
SendMode Input
SetWorkingDir %A_ScriptDir%

SetTitleMatchMode, 2
TooltipActive := false
TooltipLocked := false
VSCodeTooltipActive := false
ClaudeTooltipActive := false
ClaudeKeyDismissReadyTick := 0
TooltipCooldownMs := 1500 ; 1.5 seconds (prevent accidental double-tap)
ButtonsTooltipActive := false
interofficeExes := ["firefox.exe", "chrome.exe", "msedge.exe"]

#If WinActive("Intra Desktop Client - Assign Recip")

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
Alt+4 - Focus Notes Field
Alt+L - Lost and Found
Alt+D - Item Var Lookup + Apply All Toggle
-
Bulk Tracking Export:
Alt+Shift+E - Export tracking list to Scan field
Ctrl+Shift+Alt+E - Export tracking list to Scan field (fast)
Ctrl+Alt+E - Open tracking file (1=TXT, 2=CSV)
Ctrl+Shift+Alt+Delete - Clear both tracking files
-
Additional Scripts Launch Hotkeys:
Ctrl+Alt+F - Launch Intra Search Shortcuts
Ctrl+Alt+I - Launch Intra Extensive Automations
Ctrl+Alt+T - Show this tooltip again
    )
    Tooltip, %TooltipText%
    Hotkey, Esc, HideTooltips, On
    SetTimer, HideTooltips, -30000
Return
#If

#If WinActive("Intra Desktop Client - Update")

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
Update Tab:
Alt+P - Status Select -> Pickup from BSC
Alt+D - Status Select -> Delivery
Alt+H - Status Select -> Outbound Handed-Off
Alt+V - Status Select -> Void
Alt+S - Click Status Field
Alt+4 - Focus Notes Field
-
Bulk Tracking Export:
Alt+Shift+E - Export tracking list to Scan field
Ctrl+Shift+Alt+E - Export tracking list to Scan field (fast)
Ctrl+Alt+E - Open tracking file (1=TXT, 2=CSV)
Ctrl+Shift+Alt+Delete - Clear both tracking files
-
Additional Scripts Launch Hotkeys:
Ctrl+Alt+F - Launch Intra Search Shortcuts
Ctrl+Alt+I - Launch Intra Extensive Automations
Ctrl+Alt+T - Show this tooltip again
    )
    Tooltip, %TooltipText%
    Hotkey, Esc, HideTooltips, On
    SetTimer, HideTooltips, -30000
Return
#If

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
~!+e::Gosub HideTooltips
~^+!e::Gosub HideTooltips
~!a::Gosub HideTooltips
~!n::Gosub HideTooltips
~!1::Gosub HideTooltips
~!2::Gosub HideTooltips
~!3::Gosub HideTooltips
~!d::Gosub HideTooltips
~!h::Gosub HideTooltips
~!c::Gosub HideTooltips
~^!c::Gosub HideTooltips
~^!e::Gosub HideTooltips
~^!+Delete::Gosub HideTooltips
~^!f::Gosub HideTooltips
~^!d::Gosub HideTooltips
~^+!d::Gosub HideTooltips
~^+!s::Gosub HideTooltips
~^+!w::Gosub HideTooltips
~^!w::Gosub HideTooltips
~^!t::Gosub HideTooltips
~^!i::Gosub HideTooltips
~^i::Gosub HideTooltips
~#a::Gosub HideTooltips
~#u::Gosub HideTooltips
~#p::Gosub HideTooltips
~#f::Gosub HideTooltips
~#s::Gosub HideTooltips
~#w::Gosub HideTooltips
~#i::Gosub HideTooltips
~#!m::Gosub HideTooltips
#If

; Guard tooltip when Faster Assigning script is not running (Alt+S in Assign Recip)
#IfWinActive, Intra Desktop Client - Assign Recip
!s::
    DetectHiddenWindows, On
    running := WinExist("IT_Requested_IOs-Faster_Assigning.ahk ahk_class AutoHotkey")
    DetectHiddenWindows, Off
    if (running)
        return
    tooltipText := "Faster-Assign script inactive - Ctrl+Alt+I to launch then try hotkey again."
    Tooltip, %tooltipText%
    SetTimer, HideTooltips, -3000
return
#If

; Intra Search - show SSJ search hotkeys when Search - General is active
#IfWinActive, Search - General
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
    ButtonsTooltipActive := true
    TooltipActive := true
    tooltipText =
    (
Intra SSJ Search
Ctrl+Alt+F: Load and reload search script
Alt+D: Docksided items
Ctrl+Alt+D: Delivered items
Alt+O: On-shelf items
Alt+H: Outbound - Handed Off (down 3)
Alt+A: Arrived at BSC
Alt+P: Pickup from BSC
Alt+Space: Search Windows Quick Resize
Ctrl+Alt+T: Show this tooltip again
    )
    Tooltip, %tooltipText%
    SetTimer, HideTooltips, -15000
return
#If

UnlockTooltip:
    TooltipLocked := false
Return

; Intra Buttons (Interoffice tab any browser) tooltip
#If InterofficeActive()

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
    ButtonsTooltipActive := true
    TooltipActive := true
    tooltipText =
    (
SSJ Intra: Interoffice Requests
Alt+E/Alt+S  - Focus envelope icon
Alt+A  - Focus alias field
Alt+N  - Focus SF name field
Shift+Alt+A - Focus SF alias field (1-5 quick-select)
Ctrl+Enter - Scroll to bottom and Submit
Ctrl+Alt+S - Fill Special Instructions
Ctrl+Alt+A - ACP preset -> Alias
Alt+P  - Load "posters" preset -> Name field
Alt+Z / Alt+H - Intra Online: Home button anchor
Alt+1  - Focus "# of Packages"
Alt+2  - Focus Package Type
Alt+L  - Click Load button
Alt+C  - Clear/Reset
Ctrl+Alt+Enter - Quick submit (scroll bottom + submit)
Ctrl+W - Close open tabs (bulk prompt)
Ctrl+Alt+N - Focus Name field (recipient)
-
Intra Form Full Autos
Ctrl+Win+Alt+M - IO full auto (mfrncoa)
Ctrl+Win+Alt+1 - IO full auto (Accts Payable Roxanne)
Ctrl+Win+Alt+2 - IO full auto (Payroll Roxanne)
Ctrl+Alt+P - Poster full automation
Ctrl+Win+Alt+P - Offset Y Coordinates (No envelope button)
Ctrl+Alt+T - Show this tooltip again
    )
    Tooltip, %tooltipText%
    SetTimer, HideTooltips, -15000
Return

#If (ButtonsTooltipActive)
~Esc::Gosub HideTooltips
~!s::Gosub HideTooltips
~!e::Gosub HideTooltips
~!a::Gosub HideTooltips
~+!a::Gosub HideTooltips
~!n::Gosub HideTooltips
~!p::Gosub HideTooltips
~^!a::Gosub HideTooltips
~^!s::Gosub HideTooltips
~!h::Gosub HideTooltips
~!z::Gosub HideTooltips
~!1::Gosub HideTooltips
~!2::Gosub HideTooltips
~!l::Gosub HideTooltips
~!c::Gosub HideTooltips
~!Space::Gosub HideTooltips
~^w::Gosub HideTooltips
~^#!m::Gosub HideTooltips
~^#!1::Gosub HideTooltips
~^#!2::Gosub HideTooltips
~^!t::Gosub HideTooltips
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
Ctrl+Alt+D: API Fetch -> Paste to WorldShip (New)
Ctrl+Alt+C: Launch DSRF to WorldShip Script
Ctrl+Alt+U: Launch Super-Speed version (Warning: May be unstable)
Ctrl+Alt+P: Personal Form
Ctrl+Alt+B: Business Form
Ctrl+Alt+T - Show this tooltip again
    )
    Tooltip, %tooltipText%
    SetTimer, HideTooltips, -15000
return

#If (ButtonsTooltipActive)
~^!c::Gosub HideTooltips
~^!b::Gosub HideTooltips
~^!d::Gosub HideTooltips
#If

; UPS WorldShip shortcuts (UPS_WS_Shortcuts.ahk) when WorldShip or QVN window is active
#If ( WinActive("UPS WorldShip") || WinActive("Quantum View Notify Recipients") )
^!t::
    if (TooltipActive) {
        Gosub, HideTooltips
        Return
    }
    CoordMode, Mouse, Screen
    WinGetPos, winX, winY, winW, winH, A
    if (winW && winH)
        MouseMove, % (winX + winW//2), % (winY + winH//2)
    TooltipActive := true
    tooltipText =
    (
UPS WorldShip Hotkeys
Alt+A   - Paste @amazon.com
Alt+E   - Copy Ship To email -> QVN Recipients
Alt+N   - Copy Ship From company -> Ref2
Alt+P   - Copy Ship To phone -> Ship From phone
Alt+D   - Focus Package Weight (double-click)
Ctrl+Alt+Q - Open QVN Recipients window
Alt+S   - Open UPS Service Selection
Alt+Tab - Highlight field || Tab input (6x)
Alt+1   - Ref1 || Tab input (2x)
Alt+2   - Ref2 || Tab input (8x)
Alt+3   - Service: Next Day Air
Alt+4   - Service: Next Day Air Saver
Alt+5   - Service: 2nd Day Air
Alt+6   - Service: 3-Day Select
Alt+G   - Service: Ground
Ctrl+Alt+T - Show this tooltip again
    )
    Tooltip, %tooltipText%
    Hotkey, Esc, HideTooltips, On
    SetTimer, HideTooltips, -15000
return
#If

#If (TooltipActive && (WinActive("UPS WorldShip") || WinActive("Quantum View Notify Recipients")))
~!a::Gosub HideTooltips
~!e::Gosub HideTooltips
~!Tab::Gosub HideTooltips
~!1::Gosub HideTooltips
~!2::Gosub HideTooltips
~!s::Gosub HideTooltips
~!3::Gosub HideTooltips
~!g::Gosub HideTooltips
~!p::Gosub HideTooltips
~!n::Gosub HideTooltips
#If

; VS Code shortcuts when VS Code is active
#IfWinActive, ahk_exe Code.exe
^!t::
    if (TooltipActive) {
        Gosub, HideTooltips
        Return
    }
    TooltipActive := true
    VSCodeTooltipActive := true
    tooltipText =
    (
VS Code Shortcuts
Ctrl+Shift+P - Command Palette
Ctrl+Shift+Alt+P - Restart Extension Host
Ctrl+Shift+Alt+R - Markdown: Refresh Preview
Ctrl+Shift+Alt+I - Run installers (new device setup)
Ctrl+`` - Toggle terminal
Ctrl+Alt+R - Relaunch terminal safely (opens fresh terminal)
cd + Tab - Jump to .\Repositories\AHK-Automations\ (PowerShell)
Ctrl+P - Quick Open
Ctrl+Shift+E - Explorer
Ctrl+Shift+F - Search
Ctrl+Shift+G - Source Control
Ctrl+Shift+X - Extensions
Ctrl+K Ctrl+S - Keyboard Shortcuts
Ctrl+Alt+C - Claude Code Extension
Ctrl+Alt+T - Show this tooltip again
-
PRESS 'T' FOR OTHER/AGENT SHORTCUTS
    )
    Tooltip, %tooltipText%
    Hotkey, Esc, HideTooltips, On
    Hotkey, t, ShowClaudeTooltip, On
    SetTimer, HideTooltips, -15000
return
#If

ShowClaudeTooltip:
    if (!VSCodeTooltipActive || ClaudeTooltipActive)
        return
    VSCodeTooltipActive := false
    ClaudeTooltipActive := true
    Hotkey, t, ShowClaudeTooltip, Off
    Hotkey, *t, HideClaudeTooltipWithT, On
    ClaudeKeyDismissReadyTick := A_TickCount + 250
    SetTimer, ClaudeTooltipAnyKeyDismiss, 50
    SetTimer, HideTooltips, Off
    claudeText =
    (
Claude Code Shortcuts (Windows)
-
Input Controls:
@ - Mention files/folders/URLs
Tab - Accept suggestion
Shift+Tab - Reject suggestion
Esc - Interrupt/cancel current operation
Ctrl+C - Cancel (while Claude is responding)
Up/Down - Navigate conversation history
-
Slash Commands:
/clear - Clear conversation history
/compact - Toggle compact conversation mode
/config - Open configuration settings
/cost - Show token usage and cost
/help - Show all available commands
/init - Initialize CLAUDE.md in project
/model - Switch between Claude models
/permissions - Manage tool permissions
/memory - Edit CLAUDE.md memory files
/mcp - Manage MCP server connections
/review - Review recent code changes
-
CLI Only:
/doctor - Check Claude Code health/setup
/status - Show current session status (CLI)
/terminal-setup - Configure terminal integration (CLI)
-
Press Esc to close
    )
    Tooltip, %claudeText%
    Hotkey, Esc, HideTooltips, On
    SetTimer, HideTooltips, -30000
return

HideClaudeTooltipWithT:
    if (!ClaudeTooltipActive)
        return
    Hotkey, *t, HideClaudeTooltipWithT, Off
    Gosub, HideTooltips
    SendInput, t
return

ClaudeTooltipAnyKeyDismiss:
    if (!ClaudeTooltipActive)
    {
        SetTimer, ClaudeTooltipAnyKeyDismiss, Off
        return
    }
    if (A_TickCount < ClaudeKeyDismissReadyTick)
        return
    if (A_TimeIdleKeyboard <= 150)
        Gosub, HideTooltips
return

; Slack shortcuts when Slack is active
#IfWinActive, ahk_exe slack.exe
^!t::
    if (TooltipActive) {
        Gosub, HideTooltips
        Return
    }
    CoordMode, Mouse, Screen
    WinGetPos, winX, winY, winW, winH, A
    if (winW && winH)
        MouseMove, % (winX + winW//2), % (winY + winH//2)
    TooltipActive := true
    tooltipText =
    (
Slack Hotkeys
Alt+0 - daveyuan
Alt+1 - leona-array
Alt+2 - sps-byod
Alt+3 - sea124_ouroboros
Alt+4 - acp_bsc_comms
Alt+5 - spssea124lostnfound
Alt+S - felsusad grovfred
Alt+A - tstepama grovfred
Alt+J - @jssjens
Alt+L - @leobanks
Alt+R - @grovfred
Ctrl+Alt+T - Show this tooltip again
    )
    Tooltip, %tooltipText%
    Hotkey, Esc, HideTooltips, On
    SetTimer, HideTooltips, -15000
return
#If

#If (TooltipActive && WinActive("ahk_exe slack.exe"))
~!0::Gosub HideTooltips
~!1::Gosub HideTooltips
~!2::Gosub HideTooltips
~!3::Gosub HideTooltips
~!5::Gosub HideTooltips
~!s::Gosub HideTooltips
~!a::Gosub HideTooltips
~!j::Gosub HideTooltips
~!l::Gosub HideTooltips
~!r::Gosub HideTooltips
#If

; Intra Home window hotkeys
#IfWinActive, Intra: Home
^!t::
    if (TooltipActive) {
        Gosub, HideTooltips
        Return
    }
    TooltipActive := true
    tooltipText =
    (
Intra Home Hotkeys
Alt+Z / Alt+I - Interoffice Request anchor click (340,490)
Alt+X / Alt+O - Intra Online: Outbound Shipping Requests button anchor
Ctrl+Alt+T - Show this tooltip again
    )
    Tooltip, %tooltipText%
    Hotkey, Esc, HideTooltips, On
    SetTimer, HideTooltips, -7000
return
#If

#If (TooltipActive && WinActive("Intra: Home"))
~!i::Gosub HideTooltips
~!z::Gosub HideTooltips
~!o::Gosub HideTooltips
~!x::Gosub HideTooltips
#If

; Intra Window Switch hotkeys (global trigger, guarded against other tooltip scopes)
^!t::
    ; Skip if another tooltip scope should own ^!t
    if (WinActive("Intra: Interoffice Request")
        || WinActive("Intra: Shipping Request Form")
        || WinActive("Intra Desktop Client - Assign Recip")
        || WinActive("Intra Desktop Client - Update")
        || WinActive("Intra Desktop Client - Pickup")
        || WinActive("UPS WorldShip")
        || WinActive("ahk_exe Code.exe")
        || WinActive("ahk_exe slack.exe")
        || WinActive("Intra: Home"))
        return
    if (TooltipActive) {
        Gosub, HideTooltips
        Return
    }
    ; Park cursor at active window center to keep tooltip contextually visible.
    CoordMode, Mouse, Screen
    WinGetPos, winX, winY, winW, winH, A
    if (winW && winH)
        MouseMove, % (winX + winW//2), % (winY + winH//2)
    TooltipActive := true
    tooltipText =
    (
Global Hotkeys
Ctrl+Esc - Reload all scripts (ReloadAll.ahk)
Win+A / Win+U / Win+P - Focus/Minimize Assign / Update / Pickup
Win+F - Focus/Minimize Firefox
Win+S - Focus/Minimize Slack
Win+W - Launch/Focus/Minimize UPS WorldShip
Win+E - Focus/Minimize/Cycle Explorer
Win+Alt+E - Open new Explorer window
Win+Alt+V - Launch/Focus/Minimize VS Code
Win+Alt+R - Launch RS focus helper (Alt+Z)
Win+I - Focus/Minimize all Intra windows
Win+Alt+M - Minimize all, then focus Firefox, Outlook PWA, Slack
Alt+X / Alt+O - Intra Online: Outbound Shipping Requests button anchor
Alt+Z / Alt+H - Intra Online: Home button anchor
Ctrl+Alt+E - Open tracking file (1=TXT, 2=CSV)
Ctrl+Shift+Alt+Delete - Clear both tracking files
Ctrl+Alt+L - Launch Daily Audit + Smartsheet
Ctrl+Shift+Alt+L - Auto Daily Audit + Smartsheet
Ctrl+Shift+Alt+D - Run Daily Audit
Ctrl+Shift+Alt+S - Run Daily Smartsheet
Ctrl+Alt+W - Intra Desktop Window Organizing
Ctrl+Shift+Alt+C - Launch Coord Capture helper
Ctrl+Shift+Alt+O - Toggle coord.txt open/close (Coord Capture helper)
Ctrl+Shift+Alt+R - Reset stuck modifier keys
Ctrl+Shift+Alt+W - Toggle Window Spy
Ctrl+Alt+T - Show this tooltip again
    )
    Tooltip, %tooltipText%
    Hotkey, Esc, HideTooltips, On
    SetTimer, HideTooltips, -15000
return

HideTooltips:
    SetTimer, UnlockTooltip, Off
    TooltipLocked := false
    TooltipActive := false
    VSCodeTooltipActive := false
    ClaudeTooltipActive := false
    ClaudeKeyDismissReadyTick := 0
    ButtonsTooltipActive := false
    Hotkey, Esc, HideTooltips, Off
    Hotkey, t, ShowClaudeTooltip, Off
    Hotkey, *t, HideClaudeTooltipWithT, Off
    SetTimer, ClaudeTooltipAnyKeyDismiss, Off
    Tooltip
Return

; Intra Pickup window hotkeys
#IfWinActive, Intra Desktop Client - Pickup
^!t::
    if (TooltipActive) {
        Gosub, HideTooltips
        Return
    }
    TooltipActive := true
    tooltipText =
    (
Intra Pickup Hotkeys
Alt+1 - Signature Print Name
Alt+E - Focus Item # / Scan field
Alt+C - Clear + Return to scan
Alt+4 - Focus Notes Field
Ctrl+Alt+T - Show this tooltip again
    )
    Tooltip, %tooltipText%
    Hotkey, Esc, HideTooltips, On
    SetTimer, HideTooltips, -7000
return
#If

#If (TooltipActive)
~^Esc::Gosub HideTooltips
~!1::Gosub HideTooltips
~!e::Gosub HideTooltips
~!c::Gosub HideTooltips
#If

InterofficeActive()
{
    title := "Intra: Interoffice Request"
    if (WinActive(title " ahk_exe firefox.exe"))
        return true
    if (WinActive(title " ahk_exe chrome.exe"))
        return true
    if (WinActive(title " ahk_exe msedge.exe"))
        return true
    return WinActive(title)
}
