#Requires AutoHotkey v1
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#SingleInstance, Force  ; Reload without prompt when Esc is pressed.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

SetTitleMatchMode, 2  ; Allow partial matches on Firefox window titles.
CoordMode, Mouse, Window  ; Work with positions relative to the active Intra window.
SetDefaultMouseSpeed, 0

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

#IfWinActive, Intra: Interoffice Request ahk_exe firefox.exe
^Enter::
    EnsureIntraWindow()
    Sleep 150
    MouseClick, left, 1400, 850, 2
    Sleep 150
    SendInput, ^{End}
    Sleep 250
    MouseClick, left, 460, 1313, 2 ; this Mouseclick, the Sleep 250 before, and the following double mouseclick ensures successful button press
    Sleep 150
    MouseClick, left, 457, 1313, 2
return

!c::
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
return

!s::
!e::
    EnsureIntraWindow()
    Sleep 150
    MouseClick, left, 1400, 850, 2  ; neutral click to defocus controls
    Sleep 150
    SendInput, ^{Home}  ; scroll to top
    Sleep 150
    MouseClick, left, 730, 360
return

!a::
    EnsureIntraWindow()
    Sleep 150
    MouseClick, left, 1400, 850, 2
    Sleep 150
    SendInput, ^{Home}
    Sleep 150
    MouseClick, left, 1005, 860, 2
return

!n::
    EnsureIntraWindow()
    Sleep 150
    MouseClick, left, 1400, 850, 2
    Sleep 150
    SendInput, ^{Home}
    Sleep 150
    MouseClick, left, 450, 560, 2
return

!p::
    EnsureIntraWindow()
    Sleep 150
    MouseClick, left, 1400, 850, 2
    Sleep 150
    SendInput, ^{Home}
    Sleep 150
    Gosub, HandleEnvelopeClick
    Sleep, 500
    SendInput, poster
    Sleep 150
    SendInput, {Enter}
    Sleep 200
    MouseClick, left, 880, 815
    Sleep 150
    MouseClick, left, 485, 860
return

^!a::
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
return

HandleEnvelopeClick:
    MouseClick, left, 730, 360
    Sleep 150
return

^!s::
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
return

^!t::
return

!1::
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
return

!2::
    EnsureIntraWindow()
    Sleep 150
    MouseClick, left, 1400, 850, 2
    Sleep 150
    Loop 5 {
        SendInput, {WheelUp}
        Sleep 25
    }
    Sleep 150
    MouseClick, left, 480, 1246
return

!l::
    Send {Tab 4}
    Sleep 50
    Send {Space}
return

!Space::
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
return

EnsureIntraWindow()
{
    WinMove, Intra: Interoffice Request ahk_exe firefox.exe,, 1917, 0, 1530, 1399
    Sleep 150
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

#IfWinActive
