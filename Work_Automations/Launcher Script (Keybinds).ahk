#Requires AutoHotkey v1
#NoEnv ; Prevents Unnecessary Environment Variable lookup
#Warn ; Warn All (All Warnings Enabled)
#SingleInstance, Force  ; Reload without prompt when Esc is pressed.
SendMode Input ; Send works as SendInput
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetTitleMatchMode, 2
CoordMode, Mouse, Window

intraWinTitle := "Intra: Shipping Request Form"
intraWinExes := ["firefox.exe", "chrome.exe", "msedge.exe"]  ; priority order
TooltipActive := false
searchWinTitle := "Search - General"
searchShortcutsLaunched := false
SetTimer, DetectSearchWindow, 500  ; Watch for first appearance of Search - General

; Launch: Coordinate Helper ; Keybind: Ctrl+Shift+Alt+C
^+!c::
    Run, "C:\Users\daveyuan\Documents\GitHub\Repositories\AHK-Automations\Other_Automations\Coordinate Capture Helper\Coord_Capture.ahk"
    ShowTooltip("Coord Helper: Click to capture", 4000)
return

; Launch: Smartsheets ; Keybind: Ctrl+Alt+L
^!l:: 
    Run, C:\Users\daveyuan\Documents\GitHub\Repositories\AHK-Automations\Work_Automations\Daily_Audit.ahk ; Keybind: Ctrl+Alt+D
    Sleep 150
    Run, C:\Users\daveyuan\Documents\GitHub\Repositories\AHK-Automations\Work_Automations\Daily_Smartsheet.ahk ; Keybind: Ctrl+Alt+S
    Sleep 150
    TooltipText =
    (
Ctrl+Alt+D: Daily Audit
Ctrl+Alt+S: Daily Smartsheet
    )
    ShowTooltip(TooltipText, 5000)
Return

; Launch: Super Saiyan Intra ; Keybind: Ctrl+Alt+I
^!i:: 
    FocusAssignRecipWindow()
    MouseMove, 200, 245  ; Move cursor to scan field in Assign Recip
    Sleep 75
    Run, C:\Users\daveyuan\Documents\GitHub\Repositories\AHK-Automations\Work_Automations\IT_Requested_IOs-Faster_Assigning.ahk ; Keybind: Alt+S
    Sleep 150
    Run, C:\Users\daveyuan\Documents\GitHub\Repositories\AHK-Automations\Work_Automations\Parent_Ticket_Creation-(BYOD).ahk ; Keybind: Alt+P
    Sleep 150
    Run, C:\Users\daveyuan\Documents\GitHub\Repositories\AHK-Automations\Work_Automations\Parent_Ticket_Creation-(GENERAL).ahk ; Keybind: Ctrl+Alt+P
    Sleep 150
    Run, C:\Users\daveyuan\Documents\GitHub\Repositories\AHK-Automations\Work_Automations\IT_Asset_Move.ahk ; Keybind: Alt+T
    Sleep 150
    TooltipText =
    (
SSJ-Intra Keybinds
Alt+S: Faster Assigning
Alt+P: BYOD Parent Ticket
Ctrl+Alt+P: Parent Ticket (General)
Alt+T: IT Asset Move
Ctrl+Alt+I: Relaunch Scripts + Tooltip
    )
    ShowTooltip(TooltipText, 10000)
Return

; Launch: Intra SSJ Search ; Keybind: Ctrl+Alt+F
^!f::
    Gosub, LaunchIntraSearchShortcuts
Return

; Launch: DSRF-to-UPS WorldShip ; Keybind: Ctrl+Alt+C
^!c::
    intraTitle := GetIntraWindowTitle()
    intraExe := GetIntraBrowserExe()
    if (intraTitle = "") {
        ShowTooltip("Open Intra: Shipping Request Form in Firefox/Chrome/Edge", 4000)
        Return
    }

    CloseWorldShipScripts()

    ; Ensure Intra Window is focused and Mouse cursor is positioned
    FocusIntraWindow()
    EnsureIntraWindow()
    Sleep 50
    MouseClick, left, 410, 581 ; Move to Cost Center field for tooltip visibility
    Sleep 50
    SendInput, {WheelUp 25} ; Scroll to top of form

    if (intraExe != "firefox.exe") {
        Run, "C:\Users\daveyuan\Documents\GitHub\Repositories\AHK-Automations\Work_Automations\DSRF-to-UPS_WS-Legacy.ahk"
        Sleep 150
        TooltipText =
        (
Ctrl+Alt+B: Business Form (Legacy)
Ctrl+Alt+P: Personal Form (Legacy)
Ctrl+Alt+C: Launches Legacy (Chrome/Edge)
Current Mode: Legacy (No zoom)
        )
    } else {
        Run, "C:\Users\daveyuan\Documents\GitHub\Repositories\AHK-Automations\Work_Automations\DSRF-to-UPS_WS.ahk"
        Sleep 150
        TooltipText =
        (
Ctrl+Alt+B: Business Form
Ctrl+Alt+P: Personal Form
Ctrl+Alt+C: Launches DSRF-to-UPS_WS Script
Ctrl+Alt+U: Launches Super-Speed version (Warning: May be unstable)
Current Mode: Normal Speed
        )
    }
    if (!TooltipActive)
        ShowTooltip(TooltipText, 5000)
Return

; Launch: DSRF-to-UPS WorldShip (Super-Speed) ; Keybind: Ctrl+Alt+U
^!u::
    intraTitle := GetIntraWindowTitle()
    intraExe := GetIntraBrowserExe()
    if (intraTitle = "") {
        ShowTooltip("Open Intra: Shipping Request Form in Firefox/Chrome/Edge", 4000)
        Return
    }

    CloseWorldShipScripts()

    FocusIntraWindow()
    EnsureIntraWindow()
    Sleep 50
    MouseClick, left, 410, 581
    Sleep 50
    SendInput, {WheelUp 25}

    if (intraExe != "firefox.exe") {
        Run, "C:\Users\daveyuan\Documents\GitHub\Repositories\AHK-Automations\Work_Automations\DSRF-to-UPS_WS-Legacy.ahk"
        Sleep 150
        TooltipText =
        (
Ctrl+Alt+B: Business Form (Legacy)
Ctrl+Alt+P: Personal Form (Legacy)
Ctrl+Alt+U: Launches Legacy (Chrome/Edge)
Current Mode: Legacy (No zoom)
        )
    } else {
        Run, "C:\Users\daveyuan\Documents\GitHub\Repositories\AHK-Automations\Work_Automations\DSRF-to-UPS_WS(Super-Speed).ahk"
        Sleep 75
        TooltipText =
        (
Ctrl+Alt+B: Business Form
Ctrl+Alt+P: Personal Form
Ctrl+Alt+C: Launches DSRF-to-UPS_WS Script
Ctrl+Alt+U: Launches Super-Speed version (Warning: May be unstable)
Current Mode: Super-Speed
        )
    }
    if (!TooltipActive)
        ShowTooltip(TooltipText, 5000)
Return

FocusIntraWindow()
{
    title := GetIntraWindowTitle()
    if (title = "")
        return false
    WinActivate, %title%
    WinWaitActive, %title%,, 1
    return !ErrorLevel
}

EnsureIntraWindow()
{
    title := GetIntraWindowTitle()
    if (title = "")
        return false
    ; Match the working dimensions used in Intra_Buttons. Adjust here if the target size changes.
    WinMove, %title%,, 1917, 0, 1530, 1399
    Sleep 150
    return true
}

FocusAssignRecipWindow()
{
    ; Bring forward Assign Recip if it's open before launching the search helper.
    assignTitle := "Intra Desktop Client - Assign Recip"
    if (!WinExist(assignTitle))
        return false
    WinActivate, %assignTitle%
    WinWaitActive, %assignTitle%,, 1
    return !ErrorLevel
}

^Esc::Reload

ShowTooltip(TooltipText, durationMs)
{
    global TooltipActive
    SetTimer, HideLauncherTooltip, Off
    ToolTip, %TooltipText%
    TooltipActive := true
    SetTimer, HideLauncherTooltip, % -durationMs
}

HideLauncherTooltip:
    SetTimer, HideLauncherTooltip, Off
    TooltipActive := false
    ToolTip
Return

GetIntraWindowTitle()
{
    global intraWinTitle, intraWinExes
    ; Prefer the active Intra tab in priority order (Firefox > Chrome > Edge).
    Loop % intraWinExes.Length()
    {
        activeCandidate := intraWinTitle " ahk_exe " intraWinExes[A_Index]
        if (WinActive(activeCandidate))
            return activeCandidate
    }
    ; Otherwise pick the first present in priority order.
    Loop % intraWinExes.Length()
    {
        candidate := intraWinTitle " ahk_exe " intraWinExes[A_Index]
        if (WinExist(candidate))
            return candidate
    }
    return ""
}

GetIntraBrowserExe()
{
    global intraWinTitle, intraWinExes
    Loop % intraWinExes.Length()
    {
        activeCandidate := intraWinTitle " ahk_exe " intraWinExes[A_Index]
        if (WinActive(activeCandidate))
            return intraWinExes[A_Index]
    }
    Loop % intraWinExes.Length()
    {
        candidate := intraWinTitle " ahk_exe " intraWinExes[A_Index]
        if (WinExist(candidate))
            return intraWinExes[A_Index]
    }
    return ""
}

CloseWorldShipScripts()
{
    SetTitleMatchMode, 2
    DetectHiddenWindows, On
    WinGet, pidNormal, PID, DSRF-to-UPS_WS.ahk
    if pidNormal
        Process, Close, %pidNormal%
    WinGet, pidFast, PID, DSRF-to-UPS_WS(Super-Speed).ahk
    if pidFast
        Process, Close, %pidFast%
    WinGet, pidLegacy, PID, DSRF-to-UPS_WS-Legacy.ahk
    if pidLegacy
        Process, Close, %pidLegacy%
    DetectHiddenWindows, Off
    SetTitleMatchMode, 1
}

LaunchIntraSearchShortcuts:
    global searchShortcutsLaunched
    searchShortcutsLaunched := true
    Run, C:\Users\daveyuan\Documents\GitHub\Repositories\AHK-Automations\Work_Automations\Intra_Desktop_Search_Shortcuts.ahk
    Sleep 150
    TooltipText =
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
    )
    ShowTooltip(TooltipText, 7000)
Return

DetectSearchWindow:
    global searchShortcutsLaunched, searchWinTitle
    if (searchShortcutsLaunched)
        Return
    if WinExist(searchWinTitle)
        Gosub, LaunchIntraSearchShortcuts
Return

#If (TooltipActive)
~Esc::Gosub HideLauncherTooltip
#If

