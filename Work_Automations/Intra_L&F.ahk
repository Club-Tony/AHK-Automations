#Requires AutoHotkey v1
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Warn  ; Enable warnings to assist with detecting common errors.
#SingleInstance, Force  ; Reload without prompt when Esc is pressed.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetKeyDelay, 50
SetTitleMatchMode, 2  ; Partial matches for Intra window names.
CoordMode, Mouse, Window  ; Window-relative coordinates from Window Spy.

; Scope: Assign Recip L&F flow; abort flag lets Esc/cancel safely bail.
; Window Coordinates (Intra Desktop Client - Assign Recip):
AssignStatus := {x: 175, y: 205}
LFReceived := {x: 81, y: 91}
HistoryBtn := {x: 1300, y: 190}
NotesField := {x: 1150, y: 213}
NameField := {x: 130, y: 850}
ScanField := {x: 200, y: 245}
PackageTypeField := {x: 1175, y: 365}
BscLocationField := {x: 1100, y: 650}
CustomCarrierField := {x: 515, y: 185}
focusFieldsMsgId := 0x5557

abortHotkey := false


!l::  ; L&F received flow
    ResetAbort()
    if (!FocusAssignRecipWindow())
        return
    if (!WinActive("Intra Desktop Client - Assign Recip"))
        return
    Sleep 550
    if (AbortRequested())
        return
    ; Select Lost and Found - Received via dropdown.
    ; Click status to open dropdown, move to dropdown, scroll down, then click.
    MouseClick, left, % AssignStatus.x, % AssignStatus.y
    Sleep 250
    if (AbortRequested())
        return
    WinWait, ahk_class ComboLBox,, 2
    if (!ErrorLevel)
    {
        WinGet, lbId, ID, ahk_class ComboLBox
        if (lbId)
        {
            WinGetPos, lbX, lbY, lbW, lbH, ahk_id %lbId%
            ; Move to dropdown list at (90, 90) relative to dropdown window
            CoordMode, Mouse, Screen
            MouseMove, % lbX + 90, % lbY + 90, 0
            Sleep 100
            ; Scroll down to reach Lost and Found - Received
            SendInput, {WheelDown}
            Sleep 150
            ; Click to select
            MouseClick, left
            CoordMode, Mouse, Window
        }
        Sleep 150
    }
    Sleep 500
    if (AbortRequested())
        return
    MouseClick, left, % NameField.x, % NameField.y, 2
    Sleep 250
    if (AbortRequested())
        return
    SendInput, {Raw}sea124`,  ; Raw to avoid potential special char issues
    Sleep 250
    if (AbortRequested())
        return
    SendInput, {Enter}
    Sleep 2000
    if (AbortRequested())
        return
    SendInput, {Down}
    Sleep 200
    if (AbortRequested())
        return
    ; Trigger Intra_Focus_Fields !d action (item var lookup/apply-all).
    CallFocusFieldsAction(1)
    Sleep 200
    if (AbortRequested())
        return
    Sleep 900
    if (AbortRequested())
        return
    ; Fill package type + BSC location before scanning + notes entry.
    MouseClick, left, % PackageTypeField.x, % PackageTypeField.y, 1
    Sleep 300
    if (AbortRequested())
        return
    SendEvent, {Raw}bag
    Sleep 300
    if (AbortRequested())
        return
    ; Set Custom Carrier to Lost & Found before scanning.
    MouseClick, left, % CustomCarrierField.x, % CustomCarrierField.y, 1
    Sleep 150
    if (AbortRequested())
        return
    SendInput, {Alt up}
    KeyWait, Alt
    Sleep 50
    SendEvent, l
    Sleep 50
    SendInput, {Enter}
    Sleep 150
    if (AbortRequested())
        return
    MouseClick, left, % ScanField.x, % ScanField.y, 2
    Sleep 200
    if (AbortRequested())
        return
    SendInput, {Raw}LF
    Sleep 250
    ; Wait for barcode scan similar to IT_Asset_Move script
    ToolTip, Scan barcode on tamper-proof bag to continue script
    changed := false
    initialCaptured := false
    initialText := ""
    focusedCtrl := ""
    Loop 600  ; ~120 seconds total at 200 ms intervals
    {
        if (AbortRequested())
        {
            ToolTip
            return
        }
        Sleep 200
        ControlGetFocus, loopFocus, A
        if (loopFocus = "")
            continue
        if (!initialCaptured)
        {
            ControlGetText, initialText, %loopFocus%, A
            focusedCtrl := loopFocus
            initialCaptured := true
            continue
        }
        ControlGetText, newText, %focusedCtrl%, A
        if (newText != initialText && newText != "")
        {
            changed := true
            break
        }
    }
    if (!changed)
    {
        ToolTip, L&F script timed out waiting for scan
        Sleep 3000
        ToolTip
        return
    }
    ToolTip  ; clear scan prompt on success
    Sleep 750
    if (AbortRequested())
        return
    MouseClick, left, % HistoryBtn.x, % HistoryBtn.y, 2
    Sleep 200
    if (AbortRequested())
        return
    MouseClick, left, % NotesField.x, % NotesField.y
    ToolTip, % "Format: BagSerial#, Item Description"
    Sleep 5000
    ToolTip
return

#If WinActive("Intra Desktop Client - Assign Recip") && !CoordHelperActive()
Esc::
    abortHotkey := true
return
#If

CoordHelperActive()
{
    DetectHiddenWindows, On
    running := WinExist("Coord_Capture.ahk ahk_class AutoHotkey")
    DetectHiddenWindows, Off
    return running
}

ResetAbort()
{
    global abortHotkey
    abortHotkey := false
}

AbortRequested()
{
    global abortHotkey
    return abortHotkey
}

CallFocusFieldsAction(action := 1)
{
    global focusFieldsMsgId
    DetectHiddenWindows, On
    if (!WinExist("Intra_Focus_Fields.ahk ahk_class AutoHotkey"))
    {
        focusFieldsPath := A_ScriptDir "\Intra_Focus_Fields.ahk"
        if (FileExist(focusFieldsPath))
        {
            Run, %focusFieldsPath%
            WinWait, Intra_Focus_Fields.ahk ahk_class AutoHotkey,, 2
        }
    }
    if WinExist("Intra_Focus_Fields.ahk ahk_class AutoHotkey")
        PostMessage, %focusFieldsMsgId%, %action%, 0,, Intra_Focus_Fields.ahk ahk_class AutoHotkey
    DetectHiddenWindows, Off
}

FocusAssignRecipWindow()
{
    ; Bring forward Assign Recip if it's open before running the flow.
    assignTitle := "Intra Desktop Client - Assign Recip"
    if (!WinExist(assignTitle))
        return false
    WinActivate, %assignTitle%
    WinWaitActive, %assignTitle%,, 1
    return !ErrorLevel
}
