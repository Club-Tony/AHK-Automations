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
AliasField := {x: 945, y: 850}
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
    ; Select Lost and Found - Received via keyboard navigation
    MouseClick, left, % AssignStatus.x, % AssignStatus.y
    Sleep 250
    if (AbortRequested())
        return
    SendInput, o
    Sleep 150
    MouseClick, left, % AssignStatus.x, % AssignStatus.y
    Sleep 150
    SendInput, {Up}
    Sleep 250
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

+!a::
    DoAssignRecipShiftAltA()
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

DoAssignRecipShiftAltA()
{
    global NameField, AliasField

    ; Show sender alias selection tooltip
    prompt := "Enter Sender Alias:`n1: jssjens`n2: osterios`n3: keobeker`n4: leobanks`n5: SEA124"
    Tooltip, %prompt%

    ; Wait for single key input, 10 second timeout
    Input, key, L1 T10
    err := ErrorLevel
    Tooltip

    if (err = "Timeout")
        return
    if (key = Chr(27))  ; Esc
        return

    if (key = "1")
    {
        ; jssjens + set package type to envelope
        SendInput, jssjens
        Sleep 50
        SendInput, {Tab}
        Sleep 150
        SetAssignRecipPackageTypeEnvelope()
    }
    else if (key = "2")
    {
        SendInput, osterios
        Sleep 50
        SendInput, {Tab}
    }
    else if (key = "3")
    {
        SendInput, keobeker
        Sleep 50
        SendInput, {Tab}
    }
    else if (key = "4")
    {
        SendInput, leobanks
        Sleep 50
        SendInput, {Tab}
    }
    else if (key = "5")
    {
        ; SEA124 goes in Name field, not alias
        MouseClick, left, % NameField.x, % NameField.y, 2
        Sleep 150
        SendInput, ^a
        Sleep 50
        SendInput, sea124`,
        Sleep 2000
        SendInput, {Enter}
        Sleep 200
        SendInput, ^a
        Sleep 50
        SendInput, sea124`,
        Sleep 1500
        SendInput, {Enter}
        Sleep 250
    }
    else
    {
        ; Pass through other keys
        SendInput, %key%
        return
    }

    ; End on Alias field and show recipient quick-select
    Sleep 250
    MouseClick, left, % AliasField.x, % AliasField.y, 2
    Sleep 150
    PromptAssignRecipRecipientAlias()
}

SetAssignRecipPackageTypeEnvelope()
{
    global PackageTypeField

    MouseClick, left, % PackageTypeField.x, % PackageTypeField.y, 1
    Sleep 300
    SendEvent, e  ; Jump to "Envelope" in dropdown
    Sleep 100
    SendInput, {Enter}
    Sleep 150
}

PromptAssignRecipRecipientAlias()
{
    ; Recipient quick-select menu (1-16). Types alias then Tab.
    prompt := "Enter Recipient Alias:`n1: ouyanj`n2: yoenlee`n3: mfrncoa`n4: shrprob`n5: jirahste`n6: maloufm`n7: betharm`n8: noriekim`n9: pounan`n10: yogunn`n11: sssalta`n12: stevmura`n13: amydunc`n14: euellp`n15: josdeng`n16: eoneal"
    Tooltip, %prompt%

    Input, key1, L1 T10
    err1 := ErrorLevel
    if (err1 = "Timeout")
    {
        Tooltip
        return
    }
    if (key1 = Chr(27))
    {
        Tooltip
        return
    }

    if (!RegExMatch(key1, "^[0-9]$"))
    {
        Tooltip
        SendInput, %key1%
        return
    }
    if (key1 = "0")
    {
        Tooltip
        SendInput, %key1%
        return
    }

    choice := 0
    if (key1 = "1")
    {
        ; Wait briefly for potential second digit (10-16)
        Input, key2, L1 T0.35
        err2 := ErrorLevel
        if (err2 = "Timeout")
        {
            choice := 1
        }
        else if (key2 = Chr(27))
        {
            Tooltip
            return
        }
        else if (RegExMatch(key2, "^[0-6]$"))
        {
            choice := 10 + key2
        }
        else
        {
            choice := 1
        }
    }
    else
    {
        choice := key1 + 0
    }

    recipientAliases := {1: "ouyanj", 2: "yoenlee", 3: "mfrncoa", 4: "shrprob", 5: "jirahste", 6: "maloufm", 7: "betharm", 8: "noriekim", 9: "pounan", 10: "yogunn", 11: "sssalta", 12: "stevmura", 13: "amydunc", 14: "euellp", 15: "josdeng", 16: "eoneal"}
    alias := recipientAliases[choice]
    if (alias = "")
    {
        Tooltip
        return
    }

    Tooltip
    SendInput, %alias%
    Sleep 50
    SendInput, {Tab}
}
