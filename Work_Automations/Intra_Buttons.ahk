#Requires AutoHotkey v1
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#SingleInstance, Force  ; Reload without prompt when Esc is pressed.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#Include %A_ScriptDir%\Interoffice_YOffset.ahk

SetTitleMatchMode, 2  ; Allow partial matches on Firefox window titles.
CoordMode, Mouse, Window  ; Work with positions relative to the active Intra window.
SetDefaultMouseSpeed, 0
posterMsgId := 0x5555
OnMessage(posterMsgId, "HandlePosterMessage")
interofficeTitle := "Intra: Interoffice Request"
interofficeExes := ["firefox.exe", "chrome.exe", "msedge.exe"]
homeTitle := "Intra: Home"
searchTitle := "Intra: Search"
exportedReportTitle := "ExportedReport.pdf"
ctrlWRunning := false
exportedReportTitle := "ExportedReport.pdf"

; Scope: Intra Interoffice Request (Firefox) helpers for common form fields.

; Window Coordinates (Intra Interoffice Request - Mozilla Firefox):
; Active Window Position: x: 1721 y: 0 w: 1718 h: 1391
; Envelope button (green icon): x: 820 y: 240
; Submit button: x: 1470 y: 1066

NeutralClickX := 1400
NeutralClickY := 850
EnvelopeBtnX := 730
EnvelopeBtnY := 360
LoadBtnX := 880
LoadBtnY := 815
SubmitBtnX := 460
SubmitBtnY := 1313

NameFieldX := 485
NameFieldY := 860
AliasFieldX := 1005
AliasFieldY := 860
PackageTypeX := 480
PackageTypeY := 1246
SpecialInstrX := 510
SpecialInstrY := 1230
BuildingFieldX := 468
BuildingFieldY := 811
PackagesCountX := 480
PackagesCountY := 1154
assignRecipTitle := "Intra Desktop Client - Assign Recip"
focusFieldsMsgId := 0x5557  ; Message ID for Intra_Focus_Fields.ahk

#If IsInterofficeActive()
^#!p::
    ToggleInterofficeYOffset()
return

^Enter::
    DoCtrlEnter()
return

#If (IsInterofficeActive() && !CoordHelperActive())
!c::
    DoAltC()
return

#If IsInterofficeActive()
!s::
!e::
    DoAltE()
return

!a::
    DoAltA()
return

!n::
    DoAltN()
return

^!n::
    DoCtrlAltN()
return

^!Enter::
    DoCtrlAltEnter()
return

!p::
    DoAltP()
return

^!a::
    DoCtrlAltA()
return

+!a::
    DoShiftAltA()
return

HandleEnvelopeClick:
    MouseClick, left, 730, % IOY(360)
    Sleep 150
return

^!s::
    DoCtrlAltS()
return

!h::
!z::
    DoAnchorClick()
return

!1::
    DoAlt1()
return

!2::
    DoAlt2()
return

!l::
    DoAltL()
return

!Space::
    DoAltSpace()
return

^w::
    DoCtrlW()
return

EnsureIntraWindow()
{
    title := GetInterofficeWinTitle()
    if (title = "")
        return
    WinMove, %title%,, 1917, 0, 1530, 1399
    Sleep 150
}

DoAltE()
{
    EnsureIntraWindow()
    Sleep 150
    MouseClick, left, 1400, % IOY(850), 2  ; neutral click to defocus controls
    Sleep 150
    SendInput, ^{Home}  ; scroll to top
    Sleep 150
    MouseClick, left, 730, % IOY(360)
}

DoAltL()
{
    Send {Tab 4}
    Sleep 50
    Send {Space}
}

DoAltC()
{
    EnsureIntraWindow()
    Sleep 150
    MouseClick, left, 1400, % IOY(850), 2
    Sleep 150
    SendInput, ^{End}
    Sleep 150
    MouseClick, left, 320, % IOY(1313, "down"), 2  ; Clear/Reset
    Sleep 150
    MouseClick, left, 1400, % IOY(850), 2
    Sleep 150
    SendInput, ^{Home}
}

DoAltA()
{
    EnsureIntraWindow()
    Sleep 50
    MouseClick, left, 1400, % IOY(850), 2
    Sleep 150
    SendInput, ^{Home}
    Sleep 150
    MouseClick, left, 1005, % IOY(860), 2
    Sleep 150
    aliasText := CopyFieldText("Top", AliasFieldX, AliasFieldY)
    Sleep 150
    if (aliasText != "")
        Clipboard := aliasText  ; leave alias ready to paste after submit
    MouseClick, left, 1005, % IOY(860)
}

DoAltN()
{
    EnsureIntraWindow()
    Sleep 150
    MouseClick, left, 1400, % IOY(850), 2
    Sleep 150
    SendInput, ^{Home}
    Sleep 150
    MouseClick, left, 450, % IOY(560), 2
}

DoCtrlAltN()
{
    EnsureIntraWindow()
    Sleep 150
    MouseClick, left, 1400, % IOY(850), 2
    Sleep 150
    SendInput, ^{Home}
    Sleep 150
    MouseClick, left, 467, % IOY(858), 2
}

DoShiftAltA()
{
    EnsureIntraWindow()
    Sleep 150
    MouseClick, left, 1400, % IOY(850), 2
    Sleep 150
    SendInput, ^{Home}
    Sleep 150
    MouseClick, left, 950, % IOY(563), 2
    Sleep 100

    ; Show alias selection tooltip
    prompt := "Enter Sender Alias:`n1: jssjens`n2: osterios`n3: keobeker`n4: leobanks`n5: SEA124"
    Tooltip, %prompt%

    ; Wait for any single key input, 10 second timeout
    Input, key, L1 T10
    err := ErrorLevel
    Tooltip  ; hide tooltip

    ; Timeout or Esc - do nothing
    if (err = "Timeout")
        return
    if (key = Chr(27))  ; Esc character
        return

    ; Handle quick-select options
    if (key = "1")
    {
        ; jssjens - alias + Tab + set package type to envelope + recipient quick-select
        SendInput, jssjens
        Sleep 50
        SendInput, {Tab}
        SetPackageTypeEnvelope()
        Sleep 250
        MouseClick, left, 1005, % IOY(860), 2
        PromptRecipientAliasQuickSelect()
        return
    }
    else if (key = "2")
    {
        ; osterios - alias + Tab, then clear field + SEA124
        SendInput, osterios
        Sleep 50
        SendInput, {Tab}
        Sleep 250
        SendInput, ^a
        Sleep 50
        SendInput, {Delete}
        Sleep 250
        SendInput, SEA124
    }
    else if (key = "3")
    {
        ; keobeker - simple alias + Tab
        SendInput, keobeker
        Sleep 50
        SendInput, {Tab}
    }
    else if (key = "4")
    {
        ; leobanks - alias + Tab, then clear field + SEA124
        SendInput, leobanks
        Sleep 50
        SendInput, {Tab}
        Sleep 250
        SendInput, ^a
        Sleep 50
        SendInput, {Delete}
        Sleep 250
        SendInput, SEA124
    }
    else if (key = "5")
    {
        ; team, ouroboros - goes in Sender NAME field (not alias)
        ; Click SF name field directly (same coords as !n)
        MouseClick, left, 450, % IOY(560), 2
        Sleep 150
        ; Autocomplete pattern (same as " acp")
        SendInput, ^a
        Sleep 50
        SendInput, team, ouroboros-sea124-bsc
        Sleep 2000
        SendInput, {Enter}
        ; Failsafe repeat
        Sleep 200
        SendInput, ^a
        Sleep 50
        SendInput, team, ouroboros-sea124-bsc
        Sleep 1500
        SendInput, {Enter}
        Sleep 250
    }
    else
    {
        ; Any other key: pass it through to the field for manual typing
        SendInput, %key%
        return  ; Don't click recipient alias for manual typing
    }

    ; Options 2-5: end on Recipient Alias field
    Sleep 250
    MouseClick, left, 1005, % IOY(860), 2
}

SetPackageTypeEnvelope()
{
    if (!IsInterofficeActive())
        return

    EnsureIntraWindow()
    Sleep 100
    MouseClick, left, 1400, % IOY(850), 2
    Sleep 100
    if (!IsInterofficeYOffsetEnabled())
    {
        Loop 5 {
            SendInput, {WheelUp}
        }
        Sleep 100
    }
    MouseClick, left, 480, % IOY(1246), 2
    Sleep 150
    SendInput, env
    Sleep 250
    SendInput, {Enter}
    Sleep 250
}

PromptRecipientAliasQuickSelect()
{
    ; Recipient quick-select menu (1-9, Alt+0-6 for 10-16). Types alias then Tab.
    prompt := "Enter Recipient Alias:`n1: amydunc`n2: betharm`n3: eoneal`n4: euellp`n5: jirahste`n6: josdeng`n7: maloufm`n8: mfrncoa`n9: noriekim`nAlt+0: ouyanj`nAlt+1: pounan`nAlt+2: shrprob`nAlt+3: sssalta`nAlt+4: stevmura`nAlt+5: yoenlee`nAlt+6: yogunn"
    Tooltip, %prompt%

    ; Wait for input - capture regular keys and Alt+number combinations
    Input, key, L1 T20 M, {Esc}
    err := ErrorLevel

    if (err = "Timeout")
    {
        Tooltip
        return
    }
    if (InStr(err, "EndKey:Esc"))
    {
        Tooltip
        return
    }

    choice := 0
    ; Check for Alt+number (10-16)
    if (GetKeyState("Alt", "P"))
    {
        if (key = "0")
            choice := 10
        else if (key = "1")
            choice := 11
        else if (key = "2")
            choice := 12
        else if (key = "3")
            choice := 13
        else if (key = "4")
            choice := 14
        else if (key = "5")
            choice := 15
        else if (key = "6")
            choice := 16
    }
    else if (RegExMatch(key, "^[1-9]$"))
    {
        choice := key + 0
    }

    if (choice = 0)
    {
        ; Non-matching key - pass through
        Tooltip
        SendInput, %key%
        return
    }

    recipientAliases := {1: "amydunc", 2: "betharm", 3: "eoneal", 4: "euellp", 5: "jirahste", 6: "josdeng", 7: "maloufm", 8: "mfrncoa", 9: "noriekim", 10: "ouyanj", 11: "pounan", 12: "shrprob", 13: "sssalta", 14: "stevmura", 15: "yoenlee", 16: "yogunn"}
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

DoAltP()
{
    EnsureIntraWindow()
    Sleep 150
    MouseClick, left, 1400, % IOY(850), 2
    Sleep 150
    SendInput, ^{Home}
    Sleep 250

    if (IsInterofficeYOffsetEnabled())
    {
        ; No envelope button - use direct field entry
        ; Click Sender Name field (same coords as !n)
        MouseClick, left, 450, % IOY(560), 2
        Sleep 150
        SendInput, ^a
        Sleep 50
        SendInput, {Space}acp
        Sleep 2000
        SendInput, {Enter}
        ; Failsafe: clear and repeat if dropdown didn't load in time
        Sleep 200
        SendInput, ^a
        Sleep 50
        SendInput, {Space}acp
        Sleep 1500
        SendInput, {Enter}
        Sleep 250
        ; Click Package Type field (same coords as !2)
        MouseClick, left, 480, % IOY(1246), 2
        Sleep 150
        SendInput, env
        Sleep 250
        SendInput, {Enter}
        Sleep 250
        ; End on Recipient Name field (same coords as ^!n)
        MouseClick, left, 467, % IOY(858), 2
    }
    else
    {
        Gosub, HandleEnvelopeClick
        Sleep 1000
        SendInput, post
        Sleep 500
        SendInput, {Enter}
        Sleep 250
        MouseClick, left, 880, % IOY(815)
        Sleep 250
        MouseClick, left, 485, % IOY(860)
    }
}

DoCtrlAltA()
{
    EnsureIntraWindow()
    Sleep 150
    MouseClick, left, 1400, % IOY(850), 2
    Sleep 150
    SendInput, ^{Home}
    Sleep 150

    if (IsInterofficeYOffsetEnabled())
    {
        ; No envelope button - use direct field entry
        ; Click Sender Name field (same coords as !n)
        MouseClick, left, 450, % IOY(560), 2
        Sleep 150
        SendInput, ^a
        Sleep 50
        SendInput, {Space}acp
        Sleep 2000
        SendInput, {Enter}
        ; Failsafe: clear and repeat if dropdown didn't load in time
        Sleep 200
        SendInput, ^a
        Sleep 50
        SendInput, {Space}acp
        Sleep 1500
        SendInput, {Enter}
        Sleep 250
        ; Click Package Type field (same coords as !2)
        MouseClick, left, 480, % IOY(1246), 2
        Sleep 150
        SendInput, mid
        Sleep 250
        SendInput, {Enter}
        Sleep 250
        ; End on Alias field (same coords as !a)
        MouseClick, left, 1005, % IOY(860), 2
    }
    else
    {
        Gosub, HandleEnvelopeClick
        Sleep, 500
        SendInput, ACP
        Sleep 250
        SendInput, {Enter}
        Sleep 250
        MouseClick, left, 880, % IOY(815)
        Sleep 500
        MouseClick, left, 1005, % IOY(860)
    }
}

DoAlt1()
{
    EnsureIntraWindow()
    Sleep 150
    MouseClick, left, 1400, % IOY(850), 2
    Sleep 150
    Loop 5 {
        SendInput, {WheelUp}
        Sleep 25
    }
    Sleep 150
    MouseClick, left, 480, % IOY(1154)
}

DoAlt2()
{
    EnsureIntraWindow()
    Sleep 50
    MouseClick, left, 1400, % IOY(850), 2
    Sleep 100
    Loop 5 {
        SendInput, {WheelUp}
    }
    Sleep 100
    MouseClick, left, 480, % IOY(1246), 2
    Sleep 50
}

DoCtrlEnter()
{
    global SubmitBtnX, SubmitBtnY
    EnsureIntraWindow()
    Sleep 150
    MouseClick, left, 1005, % IOY(860), 2
    Sleep 150
    aliasText := CopyFieldText("Top", AliasFieldX, AliasFieldY)
    Sleep 150
    if (aliasText != "")
        Clipboard := aliasText  ; leave alias ready to paste after submit

    if (IsInterofficeYOffsetEnabled())
    {
        ; No scroll needed when offset ON, submit button visible at fixed coords
        MouseClick, left, 450, 1369, 2
        Sleep 150
        MouseClick, left, 450, 1369, 2
    }
    else
    {
        MouseClick, left, 1400, % IOY(850), 2
        Sleep 150
        SendInput, ^{End}
        Sleep 250
        submitY := IOY(SubmitBtnY, "down")
        ; This click, the Sleep 250 before, and the following double click ensure a successful press.
        MouseClick, left, %SubmitBtnX%, %submitY%, 2
        Sleep 150
        MouseClick, left, %SubmitBtnX%, %submitY%, 2
    }
}

DoCtrlAltS()
{
    EnsureIntraWindow()
    Sleep 150
    MouseClick, left, 1400, % IOY(850), 2
    Sleep 150
    SendInput, ^{End}
    Sleep 150
    MouseClick, left, 510, % IOY(1230, "down")
    Sleep 150
    SendInput, ^a
    Sleep 80
    SendInput, {Backspace}
    Sleep 120
    SendInput, Order:{Space}
    Tooltip, Alt+Space to enter building code + recipient alias.
    SetTimer, HideSpecialTooltip, -4000
}

DoAltSpace()
{
    EnsureIntraWindow()
    Sleep 150
    MouseClick, left, 1400, % IOY(850), 2
    Sleep 150
    SendInput, ^{End}
    Sleep 150
    buildingText := CopyFieldText("", 468, 811, "build")
    Sleep 150
    MouseClick, left, 510, % IOY(1230, "down")
    Sleep 100
    SendInput, {Space 2}
    Sleep 100
    ClipSaved := ClipboardAll
    Clipboard := buildingText
    SendInput, ^v
    Sleep 100
    SendInput, {Space 2}
    Sleep 150
    ; capture alias after finishing and return focus to Special Instructions
    MouseClick, left, 1400, % IOY(850), 2
    Sleep 150
    aliasText := CopyFieldText("Top", 1005, 860)
    Sleep 150
    MouseClick, left, 1400, % IOY(850), 2
    Sleep 150
    SendInput, ^{End}
    Sleep 200
    MouseClick, left, 1400, % IOY(850), 2
    Sleep 150
    MouseClick, left, 510, % IOY(1230, "down")
    Clipboard := aliasText  ; leave alias on clipboard after finishing
    Sleep 100
    SendInput, ^v
    Sleep 100
    SendInput, @
}

ToggleInterofficeYOffset()
{
    global coordToggleIni, coordToggleSection, coordToggleKey
    IniRead, enabled, %coordToggleIni%, %coordToggleSection%, %coordToggleKey%, 0
    newValue := (enabled = 1) ? 0 : 1
    IniWrite, %newValue%, %coordToggleIni%, %coordToggleSection%, %coordToggleKey%
    state := newValue ? "ON" : "OFF"
    Tooltip, % "Offset Y Coordinates (No envelope button): " state
    SetTimer, HideInterofficeToggleTooltip, -1500
}

HideInterofficeToggleTooltip:
    Tooltip
return

DoCtrlAltEnter()
{
    global SubmitBtnX, SubmitBtnY
    EnsureIntraWindow()
    Sleep 150

    if (IsInterofficeYOffsetEnabled())
    {
        ; No scroll needed when offset ON, submit button visible at fixed coords
        MouseClick, left, 450, 1369, 2
        Sleep 150
        MouseClick, left, 450, 1369, 2
    }
    else
    {
        MouseClick, left, 1400, % IOY(850), 2
        Sleep 150
        SendInput, ^{End}
        Sleep 250
        submitY := IOY(SubmitBtnY, "down")
        ; This click, the Sleep 250 before, and the following double click ensure a successful press.
        MouseClick, left, %SubmitBtnX%, %submitY%, 2
        Sleep 150
        MouseClick, left, %SubmitBtnX%, %submitY%, 2
    }

    ; After submit, wait for PDF and focus alias
    Sleep 1500
    WaitForExportedReportAndPrint(15000)
    Sleep 500
    FocusAssignRecipAndAlias()
}

HandlePosterMessage(wParam, lParam, msg, hwnd)
{
    global posterMsgId
    if (msg != posterMsgId)
        return
    if (wParam = 1)
        DoAltE()
    else if (wParam = 2)
        DoAltL()
    else if (wParam = 3)
        DoAltP()
    else if (wParam = 4)
        DoAlt2()
    else if (wParam = 5)
        DoAltN()
    else if (wParam = 6)
        DoCtrlAltN()
    else if (wParam = 7)
        DoCtrlAltEnter()
    else if (wParam = 8)
        DoCtrlW()
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

DoAnchorClick()
{
    EnsureIntraWindow()
    Sleep 100
    MouseClick, left, 75, 150, 2
}

DoHomeAlt1()
{
    EnsureHomeWindow()
    Sleep 100
    MouseClick, left, 340, 490, 2
}

DoHomeAlt2()
{
    EnsureHomeWindow()
    Sleep 100
    CoordMode, Mouse, Screen
    MouseClick, left, 2600, 1000, 2
    CoordMode, Mouse, Window
}

DoSearchAltZ()
{
    EnsureSearchWindow()
    Sleep 100
    CoordMode, Mouse, Screen
    MouseClick, left, 2000, 150, 2
    CoordMode, Mouse, Window
}

GetFieldText(scrollPos, xRatio, yRatio, promptText := "")
{
    text := CopyFieldText(scrollPos, xRatio, yRatio)
    if (text = "" && promptText != "")
    {
        InputBox, userText, Field Required, %promptText%
        if (!ErrorLevel)
            text := userText
    }
    return text
}

HideSpecialTooltip()
{
    Tooltip
}

EnsureHomeWindow()
{
    title := GetHomeWinTitle()
    if (title = "")
        return
    WinActivate, %title%
    WinWaitActive, %title%,, 1
}

EnsureSearchWindow()
{
    title := GetSearchWinTitle()
    if (title = "")
        return
    WinActivate, %title%
    WinWaitActive, %title%,, 1
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

GetHomeWinTitle()
{
    global homeTitle, interofficeExes
    for _, exe in interofficeExes
    {
        candidate := homeTitle " ahk_exe " exe
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

IsHomeActive()
{
    title := GetHomeWinTitle()
    return (title != "" && WinActive(title))
}

GetSearchWinTitle()
{
    global searchTitle, interofficeExes
    for _, exe in interofficeExes
    {
        candidate := searchTitle " ahk_exe " exe
        if (WinExist(candidate))
            return candidate
    }
    return ""
}

IsSearchActive()
{
    title := GetSearchWinTitle()
    return (title != "" && WinActive(title))
}

CoordHelperActive()
{
    DetectHiddenWindows, On
    running := WinExist("Coord_Capture.ahk ahk_class AutoHotkey")
    DetectHiddenWindows, Off
    return running
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

ClearPdfFailTooltip:
    ToolTip
return

FocusAssignRecipAndAlias()
{
    global assignRecipTitle, focusFieldsMsgId
    WinActivate, %assignRecipTitle%
    WinWaitActive, %assignRecipTitle%,, 2
    if (ErrorLevel)
        return false
    Sleep 200
    ; Use PostMessage to trigger alias focus via Intra_Focus_Fields.ahk (wParam=2)
    DetectHiddenWindows, On
    if WinExist("Intra_Focus_Fields.ahk ahk_class AutoHotkey")
        PostMessage, %focusFieldsMsgId%, 2, 0,, Intra_Focus_Fields.ahk ahk_class AutoHotkey
    DetectHiddenWindows, Off
    return true
}

IsCloseableWindow()
{
    interTitle := GetInterofficeWinTitle()
    expTitle := GetExportedReportWinTitle()
    return ( (interTitle != "" && WinActive(interTitle))
          || (expTitle != "" && WinActive(expTitle)) )
}

DoCtrlW()
{
    global ctrlWRunning
    if (ctrlWRunning)
        return
    ctrlWRunning := true

    if (!IsCloseableWindow())
    {
        ctrlWRunning := false
        return
    }

    prompt := "Close tabs?`nEnter = 15x`nC = 32x`nW = only current"
    Tooltip, %prompt%

    action := WaitForCloseTabsChoice()
    Tooltip
    ctrlWRunning := false

    if (action = "single")
    {
        if (IsCloseableWindow())
            SendInput, ^w
        return
    }
    else if (action != "bulk" && action != "bulk32")
    {
        return
    }

    maxCount := (action = "bulk32") ? 32 : 15
    ; Bulk close up to maxCount tabs, aborting on Esc or window loss.
    Loop, %maxCount%
    {
        if (GetKeyState("Esc", "P") || GetKeyState("Escape", "P"))
            break
        if (!IsCloseableWindow())
            break
        SendInput, ^w
        Sleep 120
    }

    ; If the active tab isn't Intra Home or Interoffice after bulk close, re-open Intra Home.
    if (!IsHomeActive() && !IsInterofficeActive())
    {
        SendInput, ^t
        Sleep 150
        SendInput, ^l
        Sleep 100
        SendInput, {Raw}*Intra: Home
        Sleep 350
        SendInput, {Down}
        Sleep 100
        SendInput, {Enter}
    }
}

WaitForCloseTabsChoice()
{
    action := ""
    Loop
    {
        Input, key, L1 M V, {Enter}{Esc}{LControl}{RControl}{w}{W}{c}{C}
        err := ErrorLevel
        if (SubStr(err, 1, 6) = "EndKey")
        {
            keyName := SubStr(err, 8)
            if (keyName = "Enter")
                action := "bulk"
            else if (keyName = "c" || keyName = "C")
                action := "bulk32"
            else if (keyName = "w" || keyName = "W")
                action := "single"
            else if (keyName = "LControl" || keyName = "RControl")
            {
                action := "cancel"
            }
            else
                action := "cancel"
            break
        }
        else
        {
            action := "cancel"
            break
        }
    }
    return action
}

#If IsHomeActive()
!i::
!z::
    DoHomeAlt1()
return

!o::
!x::
    DoHomeAlt2()
return
#If

#If IsSearchActive()
!z::
    DoSearchAltZ()
return

!h::
    DoSearchAltZ()
return
#If

#IfWinActive
