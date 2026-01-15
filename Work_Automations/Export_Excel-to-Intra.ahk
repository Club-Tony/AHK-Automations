#Requires AutoHotkey v1
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
#SingleInstance, Force  ; Reload without prompt when Esc is pressed.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetTitleMatchMode, 2  ; Allow partial matches for Intra window title.
CoordMode, Mouse, Window  ; Window-relative coordinates from Window Spy.
SetKeyDelay, 50, 50
exportRunning := false
abortHotkey := false

; Window Coordinates (Intra Desktop Client - Assign Recip):
; Window Position: x: -7 y: 0 w: 1322 h: 1339
; 200, 245 = Scan field
; 130, 850 = Name field
; 1035, 185 = enable item var lookup button
; 1060, 185 = applying to all item vars button
; 1100, 365 = Package type
; 1100, 535 = Pieces field
; 1100, 650 = BSC Location (Destination)
; 500, 1025 = Select Preset
; 350, 900 = Docksided items preset select 
; 725, 190 = Enable Parent Item #
; 500, 120 = Print Label button
; 40, 1300 = All Button
ScanField := {x: 200, y: 245}

^Esc::Reload

!+e::  ; paste list into scan field (normal speed)
    RunExport()
return

^+!e::  ; paste list into scan field (fast speed)
    RunExport(true)
return

RunExport(isFast := false)
{
    global exportRunning, ScanField
    if (exportRunning)
    {
        MsgBox, Already running!
        return
    }
    exportRunning := true
    ResetAbort()
    if (!FocusAssignRecipWindow())
    {
        MsgBox, Could not find or activate window!
        exportRunning := false
        return
    }
    if (!WinActive("Intra Desktop Client - Assign Recip"))
    {
        MsgBox, Window not active!
        exportRunning := false
        return
    }
    trackingFile := A_ScriptDir . "\tracking_numbers.txt"
    if (!FileExist(trackingFile))
    {
        MsgBox, File not found:`n%trackingFile%
        exportRunning := false
        return
    }
    FileRead, ItemList, %trackingFile%
    if (ErrorLevel)
    {
        MsgBox, Error reading tracking file.
        exportRunning := false
        return
    }
    if (isFast)
    {
        initialSleep := 200
        clickSleep := 75
        typeSleep := 40
        enterSleep := 150
        progressEvery := 100
        modeLabel := "FAST"
    }
    else
    {
        initialSleep := 500
        clickSleep := 200
        typeSleep := 100
        enterSleep := 400
        progressEvery := 50
        modeLabel := "NORMAL"
    }
    Sleep, %initialSleep%
    itemCount := 0
    aborted := false
    Loop, Parse, ItemList, `n, `r
    {
        if (AbortRequested())
        {
            MsgBox, Aborted at item %itemCount%
            aborted := true
            break
        }
        if (A_LoopField = "")
            continue
        itemCount++
        MouseClick, left, % ScanField.x, % ScanField.y, 2
        Sleep, %clickSleep%
        if (AbortRequested())
        {
            MsgBox, Aborted at item %itemCount%
            aborted := true
            break
        }
        SendInput, % "{Raw}" A_LoopField
        Sleep, %typeSleep%
        SendInput, {Enter}
        Sleep, %enterSleep%
        if (progressEvery && Mod(itemCount, progressEvery) = 0)
            ToolTip, % "Processing " modeLabel " item " itemCount "..."
    }
    ToolTip
    if (!aborted)
        MsgBox, Complete! Processed %itemCount% items.
    exportRunning := false
}

#If ( WinActive("Intra Desktop Client - Assign Recip") && exportRunning )
Esc::
    abortHotkey := true
    ToolTip, Aborting...
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
