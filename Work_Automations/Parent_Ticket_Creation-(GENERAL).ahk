#Requires AutoHotkey v1
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
#SingleInstance, Force  ; Reload without prompt when Esc is pressed.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
bscLocationWait := ""
endWaitActive := false
aliasStage := 1
finalTooltipActive := false
finalTooltipText := "Scan into Parent field - Script Done"
finalTooltipId := 20
finalTooltipMaxMs := 15000
finalTooltipStartTick := 0
finalTooltipLastIdle := 0
finalTooltipMouseX := 0
finalTooltipMouseY := 0

; Window Coordinates (Intra Desktop Client - Assign Recip):
; Window Position: x: -7 y: 0 w: 1322 h: 1339
; 200, 245 = Scan field
; 130, 850 = Name field
; 1035, 185 = enable item var lookup button
; 1060, 185 = applying to all item vars button
; 1100, 365 = Package type
; 1100, 650 = BSC Location (Destination)
; 500, 1025 = Select Preset
; 350, 900 = Docksided items preset select 
; 725, 190 = Enable Parent Item #
; 500, 120 = Print Label button

Esc::ExitApp
#IfWinActive, Intra Desktop Client - Assign Recip
SetKeyDelay 150

^!p::
    bscLocationWait := ""
    endWaitActive := false
    aliasStage := 1
    SetTitleMatchMode, 2
    DetectHiddenWindows, On
    WinGet, myPID, PID, Parent_Ticket_Creation-(BYOD)
    if myPID ; Only close if a PID was found
        Process, Close, %myPID%
    Sleep 50
    WinGet, myPID, PID, IT_Requested_IOs-Faster_Assigning
    if myPID ; Only close if a PID was found
        Process, Close, %myPID%
    Sleep 250
    SetTitleMatchMode, 1
    Sleep 100
    DetectHiddenWindows, Off
    Sleep 250

    ; Start with toggle selection/adjusting variables
    MouseClick, left, 1035, 185
    Sleep 250
    MouseClick, left, 1060, 185
    Sleep 250

    ; Click alias/starts with delimiter, wait for enter input
    MouseClick, left, 930, 830
    Sleep 250
    MouseClick, left, 925, 850, 2
    Tooltip, Type alias and press Enter to continue script

Return

; wait for enter input after typing alias
Enter::
    if (endWaitActive)
        Return
    if (bscLocationWait)
        Return
    waitMode := (aliasStage >= 2) ? "ctrl" : "enter"
    AdvanceToBscLocation(waitMode)
Return

; continue after typing BSC location from Enter hotkey
ContinueAfterBscEnter:
    bscLocationWait := ""
    aliasStage := 2
    Tooltip
    Sleep 250
    MouseClick, left, 200, 245, 2
    Sleep 250
    SendInput, ^n
    Sleep 1000
    Send, {F5}
    Sleep 5000

    ; Click alias/starts with delimiter, wait for enter input
    MouseMove, 930, 830
    Sleep 250
    MouseClick, left
    Sleep 250
    MouseMove, 925, 850
    Sleep 250
    MouseClick, left, , , 2
    Tooltip, Type alias & press Ctrl+Enter to continue script
Return

; continue after typing BSC location from Ctrl+Enter hotkey
ContinueAfterBscCtrl:
    bscLocationWait := ""
    Tooltip
    Sleep 250
    MouseClick, left, 300, 120
    Sleep 250
    MouseClick, left, 725, 190, 2
    Sleep 250
    endWaitActive := true
    ToolTip, %finalTooltipText%, , , %finalTooltipId%
    finalTooltipActive := true
    finalTooltipStartTick := A_TickCount
    finalTooltipLastIdle := A_TimeIdlePhysical
    MouseGetPos, finalTooltipMouseX, finalTooltipMouseY
    SetTimer, FinalTooltipWatch, 100
Return

; wait for enter input after typing alias
^Enter::
    if (endWaitActive)
        Return
    if (bscLocationWait = "enter")
    {
        Gosub, ContinueAfterBscEnter
        Return
    }
    if (bscLocationWait = "ctrl")
    {
        Gosub, ContinueAfterBscCtrl
        Return
    }
    AdvanceToBscLocation("ctrl")
Return

AdvanceToBscLocation(waitMode)
{
    global bscLocationWait
    Tooltip
    SendInput, {Enter}
    WinWaitActive, Intra Desktop Client - Assign Recip,, 3
    Sleep 150
    Send, {Down}
    Sleep 250
    MouseClick, left, 1100, 365, 2
    Sleep 250
    Send mid
    Sleep 250
    SendInput, {Enter}
    MouseClick, left, 1100, 650, 2
    Sleep 250
    bscLocationWait := waitMode
    Tooltip, Type BSC Location & press Ctrl+Enter to continue script
}

FinalTooltipWatch:
    if (!finalTooltipActive)
        Return
    if ((A_TickCount - finalTooltipStartTick) >= finalTooltipMaxMs)
    {
        Gosub, FinalTooltipExit
        Return
    }
    currentIdle := A_TimeIdlePhysical
    if (currentIdle < finalTooltipLastIdle)
    {
        Gosub, FinalTooltipExit
        Return
    }
    finalTooltipLastIdle := currentIdle
    MouseGetPos, curX, curY
    if (curX != finalTooltipMouseX || curY != finalTooltipMouseY)
    {
        Gosub, FinalTooltipExit
        Return
    }
Return

FinalTooltipExit:
    SetTimer, FinalTooltipWatch, Off
    finalTooltipActive := false
    endWaitActive := false
    ToolTip, , , , %finalTooltipId%
    ExitApp
Return

#IfWinActive
