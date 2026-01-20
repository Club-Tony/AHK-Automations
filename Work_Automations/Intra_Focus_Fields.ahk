#Requires AutoHotkey v1
#NoEnv ; Prevents Unnecessary Environment Variable lookup
#Warn ; Warn All (All Warnings Enabled)
#SingleInstance, Force  ; Reload without prompt when Esc is pressed.
SendMode Input ; Send works as SendInput
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

SetTitleMatchMode, 2
; Scope: Intra Assign/Update windows; focus helpers only (tooltip moved to ToolTips.ahk).

#If ( WinActive("Intra Desktop Client - Assign Recip")
    || WinActive("Intra Desktop Client - Update") )

!e:: ; focus scan field
    MouseClick, Left, 200, 245, 2 
Return

#If WinActive("Intra Desktop Client - Assign Recip")
!a:: ; focus alias field
    Sleep 250
    MouseClick, left, 930, 830, 2
    Sleep 250
    MouseClick, left, 925, 855
Return

!n:: ; focus name field
    Sleep 250
    MouseClick, left, 130, 850, 2
Return

!1:: ; focus package type field
    Sleep 250
    MouseClick, left, 1100, 365, 2
Return

!2:: ; focus BSC location field
    Sleep 250
    MouseClick, left, 1100, 650, 2
Return

!4:: ; focus notes field (Assign Recip)
    MouseClick, left, 1300, 190, 2
    Sleep 250
    MouseClick, left, 1150, 213
Return

#If ( WinActive("Intra Desktop Client - Assign Recip") && !CoordHelperActive() )
!c:: ; clear all + submit (Assign Recip)
    MouseClick, left, 68, 1345, 2
    Sleep 200
    Send, {Enter}
    Sleep 200
    Loop, 2
    {
        SendInput, {Esc}
        Sleep 50
    }
    Sleep 200
    MouseClick, left, 966, 834
    Loop, 10
    {
        MouseClick, WheelDown
    }
    Sleep 100
    MouseClick, left, 882, 873
    Sleep 150
    MouseClick, left, 200, 245, 2  ; return focus to scan field
Return
#If WinActive("Intra Desktop Client - Assign Recip")

!d:: ; click item var lookup + apply-all buttons
    Sleep 200
    MouseClick, left, 1035, 185
    Sleep 200
    MouseClick, left, 1060, 185
    Sleep 200
    MouseClick, left, 1100, 365, 2
    Sleep 200
    Loop 50
    {
        MouseClick, WheelUp
    }
Return

#If ( WinActive("Intra Desktop Client - Assign Recip")
    || WinActive("Intra Desktop Client - Update") )

CoordHelperActive()
{
    DetectHiddenWindows, On
    running := WinExist("Coord_Capture.ahk ahk_class AutoHotkey")
    DetectHiddenWindows, Off
    return running
}
