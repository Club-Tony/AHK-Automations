#Requires AutoHotkey v1
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Warn  ; Enable warnings to assist with detecting common errors.
#SingleInstance, Force  ; Reload without prompt when Esc is pressed.
SendMode Event  ; Use SendEvent to avoid aggressive Input behavior in Pickup window.
SetKeyDelay, 50, 50
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetTitleMatchMode, 2  ; Allow partial matches for Intra window names.
CoordMode, Mouse, Window  ; Use window-relative coords with ControlClick.
; Screen coordinates (Pickup):
SigPrintName := {x: 3250, y: 267}
ClearBtn     := {x: 380,  y: 1350}
; From Window Spy, Item # scan field (client coords)
ScanField    := {x: 513,  y: 246}
sigX := SigPrintName.x, sigY := SigPrintName.y
clearX := ClearBtn.x, clearY := ClearBtn.y
scanX := ScanField.x, scanY := ScanField.y
clearClass := "WindowsForms10.Window.8.app.0.38248fc_r8_ad11294"
signatureClass := "WindowsForms10.Window.8.app.0.38248fc_r8_ad11353"
scanClasses := ["WindowsForms10.EDIT.app.0.38248fc_r8_ad111", "WindowsForms10.EDIT.app.0.38248fc_r8_ad113"]

^Esc::Reload

#If WinActive("Intra Desktop Client - Pickup")
!1::  ; focus SigPrintName field
    ControlClick, x%sigX% y%sigY%, Intra Desktop Client - Pickup,, Left, 1, NA
    Sleep 100
    MouseMove, %sigX%, %sigY%
return

!e::  ; focus scan field
    FocusScanField()
    MouseClick, left,,, 2
return

!c::  ; clear all + return to scan field (Pickup)
    if (clearClass != "")
        ControlClick, %clearClass%, Intra Desktop Client - Pickup,, Left, 1, NA
    else
        ControlClick, x%clearX% y%clearY%, Intra Desktop Client - Pickup,, Left, 1, NA
    Sleep 200
    Send, {Enter}
    Sleep 200
    Loop, 2
    {
        SendInput, {Esc}
        Sleep 50
    }
    Sleep 200
    FocusScanField()
    MouseClick, left,,, 2
return
#If

ShowTempTooltip(msg, duration := 3000)
{
    ToolTip, %msg%
    SetTimer, HidePickupTooltip, -%duration%
}

HidePickupTooltip:
    ToolTip
return

FocusScanField()
{
    global scanX, scanY, scanClasses
    target := FindScanControl()
    if (target = "signature")
    {
        ShowTempTooltip("Scan focus skipped (signature)", 1500)
        return false
    }
    if (target != "")
    {
        ControlFocus, %target%, Intra Desktop Client - Pickup
        Sleep 50
        ControlClick, %target%, Intra Desktop Client - Pickup,, Left, 1, NA
    }
    else
    {
        ControlClick, x%scanX% y%scanY%, Intra Desktop Client - Pickup,, Left, 1, NA
    }
    Sleep 100
    MouseMove, %scanX%, %scanY%
    return true
}

FindScanControl()
{
    global scanClasses, signatureClass
    MouseGetPos,,, winID, ctrlUnderMouse
    if (ctrlUnderMouse != "")
    {
        if (InStr(ctrlUnderMouse, signatureClass))
            return "signature"
        Loop % scanClasses.Length()
        {
            if (InStr(ctrlUnderMouse, scanClasses[A_Index]))
                return ctrlUnderMouse
        }
    }
    ; Try to find a matching scan control explicitly
    Loop % scanClasses.Length()
    {
        class := scanClasses[A_Index]
        ControlGetPos, , , , , %class%, Intra Desktop Client - Pickup
        if (!ErrorLevel)
            return class
    }
    return ""
}
