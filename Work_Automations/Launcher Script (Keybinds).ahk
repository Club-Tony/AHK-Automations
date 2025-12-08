#Requires AutoHotkey v1
#NoEnv ; Prevents Unnecessary Environment Variable lookup
#Warn ; Warn All (All Warnings Enabled)
#SingleInstance, Force  ; Reload without prompt when Esc is pressed.
SendMode Input ; Send works as SendInput
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetTitleMatchMode, 2
CoordMode, Mouse, Window

intraWinTitle := "Intra: Shipping Request Form ahk_exe firefox.exe"
TooltipActive := false

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
    FocusAssignRecipWindow()
    MouseMove, 200, 245  ; Move cursor to scan field in Assign Recip
    Sleep 75
    Run, C:\Users\daveyuan\Documents\GitHub\Repositories\AHK-Automations\Work_Automations\Intra_Desktop_Search_Shortcuts.ahk
    Sleep 150
    Sleep 250
    SendInput, ^f
    Sleep 100
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

; Launch: DSRF-to-UPS WorldShip ; Keybind: Ctrl+Alt+C
^!c::
    ; Ensure Intra Window is focused and Mouse cursor is positioned
    FocusIntraWindow()
    EnsureIntraWindow()
    Sleep 50
    MouseClick, left, 410, 581 ; Move to Cost Center field for tooltip visibility
    Sleep 50
    SendInput, {WheelUp 25} ; Scroll to top of form

    ; Launch DSRF-to-UPS_WorldShip.ahk after declaring focus window variables
    Run, C:\Users\daveyuan\Documents\GitHub\Repositories\AHK-Automations\Work_Automations\DSRF-to-UPS_WorldShip.ahk
    Sleep 150
    TooltipText =
    (
Ctrl+Alt+B: Business Form
Ctrl+Alt+P: Personal Form
Ctrl+Alt+C: Launches DSRF-to-UPS_WorldShip.ahk
    )
    ShowTooltip(TooltipText, 10000)
Return

FocusIntraWindow()
{
    global intraWinTitle
    WinActivate, %intraWinTitle%
    WinWaitActive, %intraWinTitle%,, 1
}

EnsureIntraWindow()
{
    global intraWinTitle
    ; Match the working dimensions used in Intra_Buttons. Adjust here if the target size changes.
    WinMove, %intraWinTitle%,, 1917, 0, 1530, 1399
    Sleep 150
}

FocusAssignRecipWindow()
{
    ; Bring forward Assign Recip if it's open before launching the search helper.
    WinActivate, Intra Desktop Client - Assign Recip
    WinWaitActive, Intra Desktop Client - Assign Recip,, 1
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

#If (TooltipActive)
~Esc::Gosub HideLauncherTooltip
#If
