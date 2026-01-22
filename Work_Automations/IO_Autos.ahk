#Requires AutoHotkey v1
#NoEnv
#Warn
#SingleInstance, Force
SendMode Input
SetWorkingDir %A_ScriptDir%
SetKeyDelay, 50

SetTitleMatchMode, 2
CoordMode, Mouse, Window
SetDefaultMouseSpeed, 0

; Window targets
interofficeTitle := "Intra: Interoffice Request"
interofficeExes := ["firefox.exe", "chrome.exe", "msedge.exe"]
assignRecipTitle := "Intra Desktop Client - Assign Recip"
exportedReportTitle := "ExportedReport.pdf"

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
ScanFieldX := 200
ScanFieldY := 245

return ; end of auto-execute section

Esc::ExitApp

^!m::
    ; Check if Intra: Interoffice Request window is active
    if (!IsInterofficeActive())
    {
        ShowTimedTooltip("Make sure Intra: Interoffice Request Tab is active", 4000)
        return
    }

    ; Ensure window is positioned and scrolled to top
    EnsureIntraWindow()
    Sleep 150
    MouseClick, left, %NeutralClickX%, %NeutralClickY%, 2
    Sleep 150
    SendInput, ^{Home}
    Sleep 150

    ; Post send: click envelope button, type mfrncoa, enter, click load
    MouseClick, left, %EnvelopeBtnX%, %EnvelopeBtnY%
    Sleep 1000
    SendInput, mfrncoa
    Sleep 500
    SendInput, {Enter}
    Sleep 250
    MouseClick, left, %LoadBtnX%, %LoadBtnY%
    Sleep 500

    ; Submit: scroll to bottom and click submit button
    MouseClick, left, %NeutralClickX%, %NeutralClickY%, 2
    Sleep 150
    SendInput, ^{End}
    Sleep 250
    MouseClick, left, %SubmitBtnX%, %SubmitBtnY%, 2
    Sleep 150
    MouseClick, left, %SubmitBtnX% - 3, %SubmitBtnY%, 2

    ; Wait for Assign Recip window to appear
    Sleep 1500

    ; Focus Intra - Assign Recip window
    FocusAssignRecipWindow()
    if (!WinActive(assignRecipTitle))
    {
        ShowTimedTooltip("Assign Recip window not found", 3000)
        return
    }

    ; Click alias field, type mfrncoa, enter, down, click scan field
    Sleep 200
    MouseClick, left, %AliasFieldX%, %AliasFieldY%
    Sleep 150
    SendInput, mfrncoa
    Sleep 150
    SendInput, {Enter}
    Sleep 400
    SendInput, {Down}
    Sleep 400
    MouseClick, left, %ScanFieldX%, %ScanFieldY%, 2

    ; Wait for ExportedReport.pdf to appear and print 1 page
    WaitForExportedReportAndPrint(15000)

    ; Focus Intra - Assign Recip again, click scan field
    FocusAssignRecipWindow()
    Sleep 200
    MouseClick, left, %ScanFieldX%, %ScanFieldY%, 2

    ; Wait for scan input to continue script
    WaitForScanAndSubmit()

    ; Focus ExportedReport.pdf and close with Ctrl+W
    expTitle := GetExportedReportWinTitle()
    if (expTitle != "")
    {
        WinActivate, %expTitle%
        WinWaitActive, %expTitle%,, 2
        Sleep 200
        SendInput, ^w
    }

    ; Check if Intra: Interoffice Request is active
    Sleep 300
    if (IsInterofficeActive())
    {
        SendInput, {Tab 2}
        Sleep 100
        SendInput, {Space}
    }
return

; ========== Helper Functions ==========

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

FocusAssignRecipWindow()
{
    global assignRecipTitle
    WinActivate, %assignRecipTitle%
    WinWaitActive, %assignRecipTitle%,, 2
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
        return false

    WinActivate, %target%
    WinWaitActive, %target%,, 2
    Sleep 300
    SendInput, ^p
    Sleep 400
    SendInput, {Enter}
    Sleep 400
    return true
}

WaitForScanAndSubmit()
{
    global assignRecipTitle, ScanFieldX, ScanFieldY

    ; Get window and control info
    WinGet, winID, ID, %assignRecipTitle%
    MouseGetPos, , , , ctrlUnderMouse
    if (ctrlUnderMouse = "")
        return

    ControlGetText, oldText, %ctrlUnderMouse%, ahk_id %winID%
    lastText := oldText
    stableCount := 0

    ; Wait up to ~8s for scan input to stabilize
    Loop, 80
    {
        Sleep 100
        ControlGetText, newText, %ctrlUnderMouse%, ahk_id %winID%
        if (newText != "" && newText != lastText)
        {
            lastText := newText
            stableCount := 1
        }
        else if (newText != "" && newText = lastText && newText != oldText)
        {
            stableCount++
        }
        if (stableCount >= 3)
        {
            Sleep 150
            SendInput, {F5}
            Sleep 500
            break
        }
    }
}

ShowTimedTooltip(msg, duration := 3000)
{
    ToolTip, %msg%
    SetTimer, HideTimedTooltip, -%duration%
}

HideTimedTooltip:
    ToolTip
return
