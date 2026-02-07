#Requires AutoHotkey v1
#NoEnv
#Warn
#UseHook
#SingleInstance, Force
SendMode Input
SetWorkingDir %A_ScriptDir%
SetKeyDelay, 50
#Include %A_ScriptDir%\Interoffice_YOffset.ahk

SetTitleMatchMode, 2
CoordMode, Mouse, Window
SetDefaultMouseSpeed, 0

; Window targets
interofficeTitle := "Intra: Interoffice Request"
interofficeExes := ["firefox.exe", "chrome.exe", "msedge.exe"]
assignRecipTitle := "Intra Desktop Client - Assign Recip"
exportedReportTitle := "ExportedReport.pdf"
intraButtonsPath := A_ScriptDir "\Intra_Buttons.ahk"

; Intra: Interoffice Request coordinates (window-relative pixels)
NeutralClickX := 1400
NeutralClickY := 850
EnvelopeBtnX := 730
EnvelopeBtnY := 360
LoadBtnX := 880
LoadBtnY := 815
SubmitBtnX := 460
SubmitBtnY := 1313

; Intra Desktop Client - Assign Recip coordinates
AliasFieldX := 945
AliasFieldY := 850
NameFieldX := 130
NameFieldY := 850
ScanFieldX := 200
ScanFieldY := 245

abortRequested := false
workflowRunning := false

return ; end of auto-execute section

#If workflowRunning
Esc::
    abortRequested := true
    ShowTimedTooltip("Workflow aborted", 1500)
return
#If

^#!m::
    RunInterofficeWorkflow("mfrncoa", false, false, "jensen, jess", false)
return

^#!1::
    RunInterofficeWorkflow("payable-", true, true, " acp", true)
return

^#!2::
    RunInterofficeWorkflow("payroll-", true, true, " acp", true)
return

; ========== Helper Functions ==========

AbortRequested()
{
    global abortRequested
    return abortRequested
}

RunInterofficeWorkflow(searchString, useNameField := false, useShortcuts := false, offsetSenderName := "", offsetUseRecipientName := false)
{
    global interofficeTitle, assignRecipTitle, abortRequested, workflowRunning
    global NeutralClickX, NeutralClickY, EnvelopeBtnX, EnvelopeBtnY
    global LoadBtnX, LoadBtnY, SubmitBtnX, SubmitBtnY
    global AliasFieldX, AliasFieldY, NameFieldX, NameFieldY, ScanFieldX, ScanFieldY

    ; Check if Intra: Interoffice Request window is active (before setting workflowRunning)
    if (!IsInterofficeActive())
    {
        ShowTimedTooltip("Make sure Intra: Interoffice Request Tab is active", 4000)
        return
    }

    abortRequested := false
    workflowRunning := true

    ; Ensure window is positioned and scrolled to top
    EnsureIntraWindow()
    Sleep 150
    MouseClick, left, %NeutralClickX%, % IOY(NeutralClickY), 2
    Sleep 150
    SendInput, ^{Home}
    Sleep 150

    if (AbortRequested())
    {
        workflowRunning := false
        return
    }

    if (IsInterofficeYOffsetEnabled())
    {
        ; Y offset ON: No envelope button - use direct field entry
        ; Sender Name field
        MouseClick, left, 450, % IOY(560), 2
        Sleep 150
        SendInput, ^a
        Sleep 50
        SendInput, %offsetSenderName%
        Sleep 2000
        SendInput, {Enter}
        Sleep 200
        SendInput, ^a
        Sleep 50
        SendInput, %offsetSenderName%
        Sleep 1500
        SendInput, {Enter}
        Sleep 250

        if (AbortRequested())
        {
            workflowRunning := false
            return
        }

        ; Package Type field
        MouseClick, left, 480, % IOY(1246), 2
        Sleep 150
        SendInput, env
        Sleep 250
        SendInput, {Enter}
        Sleep 250

        if (AbortRequested())
        {
            workflowRunning := false
            return
        }

        ; Recipient field (Alias or Recipient Name based on offsetUseRecipientName)
        if (offsetUseRecipientName)
        {
            MouseClick, left, 467, % IOY(858), 2  ; Recipient Name field
            Sleep 150
            SendInput, ^a
            Sleep 50
            SendInput, %searchString%
            Sleep 2000
            SendInput, {Enter}
            Sleep 200
            SendInput, ^a
            Sleep 50
            SendInput, %searchString%
            Sleep 1500
            SendInput, {Enter}
        }
        else
        {
            MouseClick, left, 1005, % IOY(860), 2  ; Alias field
            Sleep 150
            SendInput, ^a
            Sleep 50
            SendInput, %searchString%
            Sleep 2000
            SendInput, {Enter}
            Sleep 200
            SendInput, ^a
            Sleep 50
            SendInput, %searchString%
            Sleep 1500
            SendInput, {Enter}
        }

        if (AbortRequested())
        {
            workflowRunning := false
            return
        }

        Sleep 250
        SendInput, {Tab 2}
        Sleep 500

        aliasText := CopyFieldText("Top", 1005, 860)
        Sleep 150
        if (aliasText != "")
            Clipboard := aliasText  ; leave alias ready to paste after submit

        ; Submit at offset coords
        MouseClick, left, 450, 1369, 2
        Sleep 150
        MouseClick, left, 450, 1369, 2
    }
    else
    {
        ; Y offset OFF: Use envelope button workflow
        MouseClick, left, %EnvelopeBtnX%, % IOY(EnvelopeBtnY)
        Sleep 1000
        SendInput, %searchString%
        Sleep 500
        SendInput, {Enter}
        Sleep 250
        MouseClick, left, %LoadBtnX%, % IOY(LoadBtnY)
        Sleep 500

        if (AbortRequested())
        {
            workflowRunning := false
            return
        }

        aliasText := CopyFieldText("Top", 1005, 860)
        Sleep 150
        if (aliasText != "")
            Clipboard := aliasText  ; leave alias ready to paste after submit

        ; Submit: scroll to bottom and click submit button
        MouseClick, left, %NeutralClickX%, % IOY(NeutralClickY), 2
        Sleep 150
        SendInput, ^{End}
        Sleep 250
        MouseClick, left, %SubmitBtnX%, % IOY(SubmitBtnY, "down"), 2
        Sleep 150
        MouseClick, left, % SubmitBtnX - 3, % IOY(SubmitBtnY, "down"), 2
    }

    ; Wait for PDF to appear and print only (don't close)
    Sleep 1500

    if (AbortRequested())
    {
        workflowRunning := false
        return
    }

    pdfTitle := WaitForExportedReportPrintOnly(15000)
    if (pdfTitle = "")
    {
        workflowRunning := false
        return
    }

    if (AbortRequested())
    {
        workflowRunning := false
        return
    }

    ; Focus Intra - Assign Recip window
    Sleep 500
    FocusAssignRecipWindow()
    if (!WinActive(assignRecipTitle))
    {
        ShowTimedTooltip("Assign Recip window not found", 3000)
        workflowRunning := false
        return
    }

    ; Click alias or name field based on useNameField, type searchString, enter, down, click scan field
    Sleep 200
    if (useNameField)
        MouseClick, left, %NameFieldX%, %NameFieldY%, 2
    else
        MouseClick, left, %AliasFieldX%, %AliasFieldY%
    Sleep 150
    SendInput, %searchString%
    Sleep 150
    SendInput, {Enter}
    Sleep 400
    SendInput, {Down}
    Sleep 400
    ; Dismiss any alert popup (double Esc pattern from !c handlers)
    Loop, 2
    {
        SendInput, {Esc}
        Sleep 50
    }
    Sleep 200
    MouseClick, left, %ScanFieldX%, %ScanFieldY%, 2

    if (AbortRequested())
    {
        workflowRunning := false
        return
    }

    ; Wait for scan input, then F5 submit
    ToolTip, Scan to continue script
    WaitForScanAndSubmit()

    if (AbortRequested())
    {
        workflowRunning := false
        return
    }

    ; Close PDF tab
    Sleep 300
    ClosePdfTab()

    workflowRunning := false
}

GetInterofficeWinTitle()
{
    global interofficeTitle, interofficeExes
    for _, exe in interofficeExes
    {
        candidate := interofficeTitle " ahk_exe " exe
        if (WinExist(candidate))
            return candidate
    }
    return ""
}

IsInterofficeActive()
{
    title := GetInterofficeWinTitle()
    return (title != "" && WinActive(title))
}

EnsureIntraWindow()
{
    title := GetInterofficeWinTitle()
    if (title = "")
        return
    WinMove, %title%,, 1917, 0, 1530, 1399
    Sleep 150
}

CopyFieldText(scrollPos, xCoord, yCoord, yMode := "up")
{
    if (scrollPos = "Top")
        SendInput, ^{Home}
    else if (scrollPos = "Bottom")
        SendInput, ^{End}

    Sleep 200
    MouseClick, left, %xCoord%, % IOY(yCoord, yMode)
    Sleep 200
    ClipSaved := ClipboardAll
    Clipboard :=
    SendInput, ^a
    Sleep 80
    SendInput, ^c
    ClipWait, 0.5
    if (ErrorLevel)
        text := ""
    else
        text := Clipboard
    Clipboard := ClipSaved
    ClipSaved := ""
    return text
}

GetExportedReportWinTitle()
{
    global exportedReportTitle, interofficeExes
    for _, exe in interofficeExes
    {
        candidate := exportedReportTitle " ahk_exe " exe
        if (WinExist(candidate))
            return candidate
    }
    return ""
}

EnsureWindowActive(title, maxRetries := 3, retryDelay := 200)
{
    Loop, %maxRetries%
    {
        if (WinActive(title))
            return true
        WinActivate, %title%
        WinWaitActive, %title%,, 1
        if (!ErrorLevel)
            return true
        Sleep %retryDelay%
    }
    return false
}

FocusAssignRecipWindow()
{
    global assignRecipTitle
    return EnsureWindowActive(assignRecipTitle)
}

WaitForExportedReportAndPrint(timeoutMs := 15000)
{
    deadline := A_TickCount + timeoutMs
    target := ""
    while (A_TickCount < deadline)
    {
        target := GetExportedReportWinTitle()
        if (target != "")
            break
        Sleep 200
    }
    if (target = "")
    {
        ToolTip, Failed to detect ExportedReport.pdf window
        SetTimer, ClearPdfFailTooltip, -5000
        return false
    }

    WinActivate, %target%
    WinWaitActive, %target%,, 2
    Sleep 300
    SendInput, ^p
    Sleep 400
    SendInput, {Enter}
    Sleep 400
    ; Close PDF and refocus Interoffice
    SendInput, ^w
    Sleep 400
    SendInput, {Tab 2}
    Sleep 200
    SendInput, {Space}
    Sleep 300
    return true
}

WaitForExportedReportPrintOnly(timeoutMs := 15000)
{
    deadline := A_TickCount + timeoutMs
    target := ""
    while (A_TickCount < deadline)
    {
        target := GetExportedReportWinTitle()
        if (target != "")
            break
        Sleep 200
    }
    if (target = "")
    {
        ToolTip, Failed to detect ExportedReport.pdf window
        SetTimer, ClearPdfFailTooltip, -5000
        return ""
    }

    WinActivate, %target%
    WinWaitActive, %target%,, 2
    Sleep 300
    SendInput, ^p
    Sleep 400
    SendInput, {Enter}
    Sleep 400
    return target
}

ClosePdfTab()
{
    target := GetExportedReportWinTitle()
    if (target = "")
        return
    WinActivate, %target%
    WinWaitActive, %target%,, 2
    Sleep 300
    SendInput, ^w
    Sleep 400
    SendInput, {Tab 2}
    Sleep 200
    SendInput, {Space}
}

ClearPdfFailTooltip:
    ToolTip
return

WaitForScanAndSubmit()
{
    ; Get window and control under mouse (proven pattern for Intra Desktop Client)
    MouseGetPos, , , winID, ctrlUnderMouse
    if (ctrlUnderMouse = "")
    {
        ShowTimedTooltip("Could not detect control under mouse", 3000)
        return
    }

    ControlGetText, oldText, %ctrlUnderMouse%, ahk_id %winID%
    lastText := oldText
    stableCount := 0

    ; Wait up to ~15s for scan input to stabilize
    Loop, 150
    {
        if (AbortRequested())
            return
        Sleep 100
        ControlGetText, newText, %ctrlUnderMouse%, ahk_id %winID%
        if (newText != "" && newText != lastText)
        {
            lastText := newText
            stableCount := 1
        }
        else if (newText != "" && newText = lastText)
        {
            stableCount++
        }
        if (stableCount >= 3)
        {
            ToolTip
            Sleep 150
            SendInput, {F5}
            Sleep 500
            return
        }
    }
    ; Timeout — scan not detected within ~8s
    ShowTimedTooltip("Scan timed out", 3000)
}

EnsureIntraButtonsScript()
{
    global intraButtonsPath
    DetectHiddenWindows, On
    running := WinExist("Intra_Buttons.ahk ahk_class AutoHotkey")
    DetectHiddenWindows, Off
    if (!running && FileExist(intraButtonsPath))
        Run, %intraButtonsPath%
}

ShowTimedTooltip(msg, duration := 3000)
{
    ToolTip, %msg%
    SetTimer, HideTimedTooltip, -%duration%
}

HideTimedTooltip:
    ToolTip
return
