#Requires AutoHotkey v1
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#SingleInstance, Force  ; Reload without prompt when Esc is pressed.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

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
^Esc::Reload

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

#If IsInterofficeActive()
^Enter::
    DoCtrlEnter()
return

!c::
    DoAltC()
return

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

HandleEnvelopeClick:
    MouseClick, left, 730, 360
    Sleep 150
return

^!s::
    DoCtrlAltS()
return

^!t::
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
    MouseClick, left, 1400, 850, 2  ; neutral click to defocus controls
    Sleep 150
    SendInput, ^{Home}  ; scroll to top
    Sleep 150
    MouseClick, left, 730, 360
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
    MouseClick, left, 1400, 850, 2
    Sleep 150
    SendInput, ^{End}
    Sleep 150
    MouseClick, left, 320, 1315, 2  ; Clear/Reset
    Sleep 150
    MouseClick, left, 1400, 850, 2
    Sleep 150
    SendInput, ^{Home}
}

DoAltA()
{
    EnsureIntraWindow()
    Sleep 50
    MouseClick, left, 1400, 850, 2
    Sleep 150
    SendInput, ^{Home}
    Sleep 150
    MouseClick, left, 1005, 860, 2
    Sleep 150
    aliasText := CopyFieldText("Top", AliasFieldX, AliasFieldY)
    Sleep 150
    if (aliasText != "")
        Clipboard := aliasText  ; leave alias ready to paste after submit
    MouseClick, left, 1005, 860
}

DoAltN()
{
    EnsureIntraWindow()
    Sleep 150
    MouseClick, left, 1400, 850, 2
    Sleep 150
    SendInput, ^{Home}
    Sleep 150
    MouseClick, left, 450, 560, 2
}

DoCtrlAltN()
{
    EnsureIntraWindow()
    Sleep 150
    MouseClick, left, 1400, 850, 2
    Sleep 150
    SendInput, ^{Home}
    Sleep 150
    MouseClick, left, 467, 858, 2
}

DoAltP()
{
    EnsureIntraWindow()
    Sleep 50
    MouseClick, left, 1400, 850, 2
    Sleep 50
    SendInput, ^{Home}
    Sleep 150
    Gosub, HandleEnvelopeClick
    Sleep, 350
    SendInput, post
    Sleep 150
    SendInput, {Enter}
    Sleep 200
    MouseClick, left, 880, 815
    Sleep 150
    MouseClick, left, 485, 860
}

DoCtrlAltA()
{
    EnsureIntraWindow()
    Sleep 150
    MouseClick, left, 1400, 850, 2
    Sleep 150
    SendInput, ^{Home}
    Sleep 150
    Gosub, HandleEnvelopeClick
    Sleep, 500
    SendInput, ACP
    Sleep 150
    SendInput, {Enter}
    Sleep 200
    MouseClick, left, 880, 815
    Sleep 500
    MouseClick, left, 1005, 860
}

DoAlt1()
{
    EnsureIntraWindow()
    Sleep 150
    MouseClick, left, 1400, 850, 2
    Sleep 150
    Loop 5 {
        SendInput, {WheelUp}
        Sleep 25
    }
    Sleep 150
    MouseClick, left, 480, 1154
}

DoAlt2()
{
    EnsureIntraWindow()
    Sleep 50
    MouseClick, left, 1400, 850, 2
    Sleep 100
    Loop 5 {
        SendInput, {WheelUp}
    }
    Sleep 100
    MouseClick, left, 480, 1246, 2
}

DoCtrlEnter()
{
    EnsureIntraWindow()
    Sleep 150
    MouseClick, left, 1005, 860, 2
    Sleep 150
    aliasText := CopyFieldText("Top", AliasFieldX, AliasFieldY)
    Sleep 150
    if (aliasText != "")
        Clipboard := aliasText  ; leave alias ready to paste after submit
    MouseClick, left, 1400, 850, 2
    Sleep 150
    SendInput, ^{End}
    Sleep 250
    MouseClick, left, 460, 1313, 2 ; this Mouseclick, the Sleep 250 before, and the following double mouseclick ensures successful button press
    Sleep 150
    MouseClick, left, 457, 1313, 2
}

DoCtrlAltS()
{
    EnsureIntraWindow()
    Sleep 150
    MouseClick, left, 1400, 850, 2
    Sleep 150
    SendInput, ^{End}
    Sleep 150
    MouseClick, left, 510, 1230
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
    MouseClick, left, 1400, 850, 2
    Sleep 150
    SendInput, ^{End}
    Sleep 150
    buildingText := CopyFieldText("", 468, 811)
    Sleep 150
    MouseClick, left, 510, 1230
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
    MouseClick, left, 1400, 850, 2
    Sleep 150
    aliasText := CopyFieldText("Top", 1005, 860)
    Sleep 150
    MouseClick, left, 1400, 850, 2
    Sleep 150
    SendInput, ^{End}
    Sleep 200
    MouseClick, left, 1400, 850, 2
    Sleep 150
    MouseClick, left, 510, 1230
    Clipboard := aliasText  ; leave alias on clipboard after finishing
    Sleep 100
    SendInput, ^v
    Sleep 100
    SendInput, @
}

DoCtrlAltEnter()
{
    EnsureIntraWindow()
    Sleep 150
    MouseClick, left, 1400, 850, 2
    Sleep 150
    SendInput, ^{End}
    Sleep 250
    MouseClick, left, 460, 1313, 2 ; this Mouseclick, the Sleep 250 before, and the following double mouseclick ensures successful button press
    Sleep 150
    MouseClick, left, 457, 1313, 2
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

CopyFieldText(scrollPos, xCoord, yCoord)
{
    if (scrollPos = "Top")
        SendInput, ^{Home}
    else if (scrollPos = "Bottom")
        SendInput, ^{End}

    Sleep 200
    MouseClick, left, %xCoord%, %yCoord%
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
