#Requires AutoHotkey v1
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Warn  ; Enable warnings to assist with detecting common errors.
#SingleInstance, Force  ; Reload without prompt when Esc is pressed.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

SetTitleMatchMode, 2  ; Allow partial matches on Intra tab titles.
CoordMode, Mouse, Window  ; Work with positions relative to the active Intra window.
SetDefaultMouseSpeed, 0
SetKeyDelay, 50

posterWinTitle := "Intra: Interoffice Request"
posterWinExes := ["firefox.exe", "chrome.exe", "msedge.exe"]  ; priority order
exportedReportTitle := "ExportedReport.pdf"
intraButtonsPath := A_ScriptDir "\Intra_Buttons.ahk"
posterHotkeyRunning := false
posterHotkeyCancelled := false
posterMsgId := 0x5555
posterActionAltE := 1
posterActionAltL := 2
posterActionAltP := 3
posterActionAlt2 := 4
posterActionAltN := 5
posterActionCtrlAltN := 6
posterActionCtrlAltEnter := 7
posterActionCtrlW := 8

; Scope: Intra: Interoffice Request (browser) poster automation (priority: Firefox > Chrome > Edge).

^!p::  ; TODO: full poster automation
    posterHotkeyCancelled := false
    posterHotkeyRunning := true
    startTick := A_TickCount
    if (!FocusPosterWindow())
    {
        ShowTimedTooltip("Open Intra: Interoffice Request in Firefox/Chrome/Edge", 3000)
        Gosub, PosterHotkeyCleanup
        return
    }
    ; Start persistent status tooltip
    SetTimer, ShowActiveTooltip, 100
    ; First action: capture current URL and open duplicates.
    Clipboard := ""  ; Clear clipboard first
    send, ^l
    sleep, 50
    send, ^c
    ClipWait, 2  ; Wait up to 2 seconds for clipboard
    if (ErrorLevel || Clipboard = "" || !InStr(Clipboard, "://"))
    {
        ShowTimedTooltip("Failed to copy URL - clipboard empty or invalid", 4000)
        Gosub, PosterHotkeyCleanup
        return
    }
    copiedUrl := Clipboard  ; Store for later verification

    Loop, 15
    {
        send, ^t
        sleep, 150
        send, ^v
        sleep, 100
        send, {Enter}
        sleep, 150
    }
    Sleep 500

    ; Return to the original tab.
    Loop, 15
    {
        send, ^+{Tab}
        sleep, 50
        if (!WinActive(GetPosterWindowTitle()))
        {
            send, ^l
            sleep, 100
            send, ^v
            sleep, 100
            send, {Enter}
            sleep, 200
        }
    }

    Sleep 1000
    
    curTitle := GetPosterWindowTitle()
    if (curTitle = "" || !WinActive(curTitle))
    {
        ShowTimedTooltip("Hotkey break: IO Request window N/A or Unknown", 5000)
        Gosub, PosterHotkeyCleanup
        return
    }

    EnsureIntraButtonsScript()

    ; Already at tab 1 after "Return to the original tab" loop above
    ; Add further poster automation steps here (form fills, clicks, etc.)

    ; Mid-Size Boxes (tabs 1-3)
    Loop, 3
    {
        CallIntraButtonsHotkey(posterActionAltP, true)
        Sleep 2500
        MouseClick, left, 480, 1246, 2
        Sleep 150
        Send, {Down 2}
        Sleep 400
        Send, {Enter}
        Sleep 100
        Send ^{Tab}
        Sleep 100
        if (!VerifyPosterTab())
        {
            ShowTimedTooltip("Lost Intra tab during Mid-Size Boxes loop", 5000)
            Gosub, PosterHotkeyCleanup
            return
        }
    }

    ; Envelopes (tabs 4-16)
    Loop, 13
    {
        CallIntraButtonsHotkey(posterActionAltP, true)
        Sleep 2500
        Send ^{Tab}
        Sleep 100
        if (!VerifyPosterTab())
        {
            ShowTimedTooltip("Lost Intra tab during Envelopes loop", 5000)
            Gosub, PosterHotkeyCleanup
            return
        }
    }

    ; Position back to first tab before Name Fields operations
    ; Go back 15 tabs from wherever we are to get to tab 1
    Loop, 15
    {
        Send ^+{Tab}
        Sleep 50
    }
    Sleep 500

    if (!VerifyPosterTab())
    {
        ShowTimedTooltip("Lost Intra tab after Envelopes, before Name Fields", 5000)
        Gosub, PosterHotkeyCleanup
        return
    }

    ; Name Fields (all 16 tabs)
    names := [107,83,33,129,129,129,99,99,132,125,114,111,93,109,74,69]
    maxRetries := 2  ; Up to 2 retries (3 total attempts)

    Loop % names.Length()
    {
        currentName := names[A_Index]
        isMidBox := (A_Index <= 3)  ; First 3 need Mid-Size package type
        success := false

        ; First attempt (no reset)
        if (AttemptNameEntryAndSubmit(currentName, isMidBox, false))
        {
            success := true
        }
        else
        {
            ; Retry loop with full reset
            Loop, %maxRetries%
            {
                ShowTimedTooltip("Retry " . A_Index . " for: " . currentName, 2000)
                Sleep 500
                if (AttemptNameEntryAndSubmit(currentName, isMidBox, true))
                {
                    success := true
                    break
                }
            }
        }

        if (!success)
        {
            ShowTimedTooltip("Failed after " . (maxRetries + 1) . " attempts for: " . currentName, 5000)
            Gosub, PosterHotkeyCleanup
            return
        }

        ; Only tab to next if not on the last name
        if (A_Index < names.Length())
        {
            Sleep 300
            Send ^{Tab}
            Sleep 150
            if (!VerifyPosterTab())
            {
                ShowTimedTooltip("Lost Intra tab during Name Fields loop", 5000)
                Gosub, PosterHotkeyCleanup
                return
            }
        }
    }

    ; Close relevant tabs/end the script
    CallIntraButtonsHotkey(posterActionCtrlW)
    Sleep 300
    SendInput, c
    Sleep 5000  ; Wait for bulk close to complete (32 tabs * ~120ms each)

    ; Check if we ended up on Intra: Interoffice Request
    if (!WinActive(GetPosterWindowTitle()))
    {
        ; Recovery: navigate to Intra Home
        Send ^t
        Sleep 150
        SendInput, *intra: home
        Sleep 250
        Send {Down}
        Send {Enter}
    }

    Gosub, PosterHotkeyCleanup
    ShowHotkeyRuntime(startTick)
return

PosterWindowExists()
{
    return GetPosterWindowTitle() != ""
}

FocusPosterWindow()
{
    title := GetPosterWindowTitle()
    if (title = "")
        return false
    WinActivate, %title%
    WinWaitActive, %title%,, 1
    return !ErrorLevel
}

GetPosterWindowTitle()
{
    global posterWinTitle, posterWinExes
    for _, exe in posterWinExes
    {
        candidate := posterWinTitle " ahk_exe " exe
        if (WinExist(candidate))
            return candidate
    }
    return ""
}

VerifyPosterTab()
{
    ; Check if still on Intra: Interoffice Request tab
    ; If not, send Ctrl+Shift+Tab to go back once and check again
    local checkTitle
    checkTitle := GetPosterWindowTitle()
    if (checkTitle = "" || !WinActive(checkTitle))
    {
        Send ^+{Tab}
        Sleep 150
        checkTitle := GetPosterWindowTitle()
        if (checkTitle = "" || !WinActive(checkTitle))
            return false
    }
    return true
}

GetExportedReportWindow()
{
    global exportedReportTitle, posterWinExes
    for _, exe in posterWinExes
    {
        candidate := exportedReportTitle " ahk_exe " exe
        if (WinExist(candidate))
            return candidate
    }
    return ""
}

WaitForExportedReportAndPrint(timeoutMs := 20000)
{
    deadline := A_TickCount + timeoutMs
    target := ""
    while (A_TickCount < deadline)
    {
        target := GetExportedReportWindow()
        if (target != "")
            break
        ; Check if a different (error) page opened instead
        if (!WinActive(GetPosterWindowTitle()) && target = "")
        {
            ; Error page likely opened - skip printing, Ctrl+Tab past it
            Sleep 200
            Send ^{Tab}
            Sleep 150
            return true  ; Continue without retry
        }
        Sleep 200
    }
    if (target = "")
    {
        ; 20 sec timeout - assume error page or stuck, skip and continue
        Send ^{Tab}
        Sleep 150
        return true
    }

    WinActivate, %target%
    WinWaitActive, %target%,, 2
    SendInput, ^p
    Sleep 400
    SendInput, {Enter}
    Sleep 400
    return true
}

ShowTimedTooltip(msg, duration := 3000)
{
    ToolTip, %msg%
    SetTimer, HideTimedTooltip, -%duration%
}

HideTimedTooltip:
    ToolTip
return

ShowHotkeyRuntime(startTick)
{
    elapsedMs := A_TickCount - startTick
    elapsedSec := Round(elapsedMs / 1000.0, 2)
    ToolTip, Hotkey Runtime: %elapsedSec% seconds
    SetTimer, HideRuntimeTooltip, -4000
}

HideRuntimeTooltip:
    ToolTip
return

EnsureIntraButtonsScript()
{
    global intraButtonsPath
    DetectHiddenWindows, On
    scriptRunning := WinExist("Intra_Buttons.ahk ahk_class AutoHotkey")
    DetectHiddenWindows, Off
    if (!scriptRunning && FileExist(intraButtonsPath))
        Run, %intraButtonsPath%
}

CallIntraButtonsHotkey(combo, waitForCompletion := false)
{
    global posterMsgId, intraButtonsPath
    DetectHiddenWindows, On
    if (!WinExist("Intra_Buttons.ahk ahk_class AutoHotkey") && FileExist(intraButtonsPath))
    {
        Run, %intraButtonsPath%
        WinWait, Intra_Buttons.ahk ahk_class AutoHotkey,, 2
    }
    if WinExist("Intra_Buttons.ahk ahk_class AutoHotkey")
    {
        if (waitForCompletion)
        {
            SendMessage, %posterMsgId%, %combo%, 0,, Intra_Buttons.ahk ahk_class AutoHotkey,,, 1000
            if (ErrorLevel = "FAIL")
                PostMessage, %posterMsgId%, %combo%, 0,, Intra_Buttons.ahk ahk_class AutoHotkey
        }
        else
        {
            PostMessage, %posterMsgId%, %combo%, 0,, Intra_Buttons.ahk ahk_class AutoHotkey
        }
    }
    DetectHiddenWindows, Off
    Sleep 100
}

AttemptNameEntryAndSubmit(nameValue, needsMidBox, isRetry)
{
    global posterActionAltP, posterActionCtrlAltN, posterActionCtrlAltEnter

    if (isRetry)
    {
        ; Full reset - re-trigger poster preset
        CallIntraButtonsHotkey(posterActionAltP, true)
        Sleep 2500

        if (needsMidBox)
        {
            ; Re-select Mid-Size package type
            MouseClick, left, 480, 1246, 2
            Sleep 150
            Send, {Down 2}
            Sleep 400
            Send, {Enter}
            Sleep 100
        }
    }

    ; Focus name field
    CallIntraButtonsHotkey(posterActionCtrlAltN)
    Sleep 1000

    ; Type name string
    SendInput, % nameValue "-r"
    Sleep 3500  ; Extended by 500ms for dropdown to load

    ; Select from dropdown
    Send {Enter}
    Sleep 50

    ; Submit form
    CallIntraButtonsHotkey(posterActionCtrlAltEnter)

    ; Wait for PDF tab to open (indicates successful submission)
    return WaitForExportedReportAndPrint(10000)
}

PosterHotkeyCleanup:
    SetTimer, ShowActiveTooltip, Off
    ToolTip
    posterHotkeyRunning := false
    posterHotkeyCancelled := false
return

ShowActiveTooltip:
    ToolTip, Script Active: Ctrl+Esc to cancel/reload
return

#If (posterHotkeyRunning)
Esc::
    ; Treat Esc like a reload/cancel while the poster hotkey is running.
    posterHotkeyCancelled := true
    Reload
return
#If
