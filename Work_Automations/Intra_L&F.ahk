#Requires AutoHotkey v1
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Warn  ; Enable warnings to assist with detecting common errors.
#SingleInstance, Force  ; Reload without prompt when Esc is pressed.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetTitleMatchMode, 2  ; Partial matches for Intra window names.
CoordMode, Mouse, Window  ; Window-relative coordinates from Window Spy.

; Scope: Assign Recip L&F flow; abort flag lets Esc/cancel safely bail.
; Window Coordinates (Intra Desktop Client - Assign Recip):
AssignStatus := {x: 190, y: 205}
LFReceived := {x: 100, y: 108}
HistoryBtn := {x: 1300, y: 190}
NotesField := {x: 1150, y: 213}
NameField := {x: 130, y: 850}
ScanField := {x: 200, y: 245}

abortHotkey := false

^Esc::
    abortHotkey := true
return

!l::  ; L&F received flow
    ResetAbort()
    if (!FocusAssignRecipWindow())
        return
    if (!WinActive("Intra Desktop Client - Assign Recip"))
        return
    Sleep 550
    if (AbortRequested())
        return
    MouseClick, left, % AssignStatus.x, % AssignStatus.y
    Sleep 200
    if (AbortRequested())
        return
    Send, l
    Sleep 200
    MouseClick, left, % AssignStatus.x, % AssignStatus.y
    Sleep 200
    Send, {Down}
    Sleep 500
    if (AbortRequested())
        return
    MouseClick, left, % NameField.x, % NameField.y, 2
    Sleep 250
    if (AbortRequested())
        return
    SendInput, {Raw}sea124, ; Raw to avoid potential special char issues
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
    MouseClick, left, % ScanField.x, % ScanField.y, 2
    Sleep 200
    if (AbortRequested())
        return
    SendInput, ^n
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

#If WinActive("Intra Desktop Client - Assign Recip")
Esc::
    abortHotkey := true
return
#If

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
