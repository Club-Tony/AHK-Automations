#Requires AutoHotkey v1
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Warn  ; Enable warnings to assist with detecting common errors.
#SingleInstance, Force  ; Reload without prompt when Esc is pressed.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetTitleMatchMode, 2  ; Partial matches for Intra window names.

; Please note the Mouse Window coordinates for the following fields when #WinActive, Intra - General:
TrackingNumField := {x: 350, y: 50}
LoadSavedSearchBtn := {x: 120, y: 720}
ClearSearchBtn := {x: 745, y: 770}
SearchBtn := {x: 880, y: 765}
DocksidedPreset := {x: 200, y: 600}
StatusSelect := {x: 375, y: 160}
; Post StatusSelect Click:
; OnShelf: send o, down, space, enter
; Delivered: send d, down, space, enter
; PickupFromBSC: send p, down 3, space, enter
; ArrivedAtBSC: send a, space, enter

; Quick restart without single-instance prompts.
^Esc::Reload

IsIntraSearchWin()
{
    return WinActive("Search - General")
        || WinActive("Search Results:")
}

; Scope any search hotkeys to the intended Intra windows.
#If IsIntraSearchWin()
!d:: ; Docksided items
    MouseClick, left, % ClearSearchBtn.x, % ClearSearchBtn.y, 3
    Sleep 250
    MouseClick, left, % LoadSavedSearchBtn.x, % LoadSavedSearchBtn.y
    Sleep 500
    MouseClick, left, % DocksidedPreset.x, % DocksidedPreset.y
    Sleep 500
    MouseClick, left, % StatusSelect.x, % StatusSelect.y
return

^!d::  ; Delivered items
    MouseClick, left, % ClearSearchBtn.x, % ClearSearchBtn.y, 3
    Sleep 250
    MouseClick, left, % LoadSavedSearchBtn.x, % LoadSavedSearchBtn.y
    Sleep 500
    MouseClick, left, % DocksidedPreset.x, % DocksidedPreset.y
    Sleep 500
    MouseClick, left, % StatusSelect.x, % StatusSelect.y
    Sleep 250
    SendInput, d
    Sleep 250
    SendInput, {Down}
    Sleep 250
    SendInput, {Space}
    Sleep 250
    SendInput, {Enter}
    Sleep 500
    WinWait, Search Results:, , 5
    if (ErrorLevel)
        return
    WinActivate, Search Results:
    WinWaitActive, Search Results:, , 2
    Sleep 100
    SendInput, !{Space}
    Sleep 100
    SendInput, s
    Sleep 100
    SendInput, {Right}
    Sleep 100
    SendInput, {Down}
    MouseMove, 2050, 1025
    Sleep 100
    SendInput, {Enter}
    Sleep 100
    MouseMove, 945, 70
return

!o::  ; On-shelf items
    MouseClick, left, % ClearSearchBtn.x, % ClearSearchBtn.y, 3
    Sleep 250
    MouseClick, left, % LoadSavedSearchBtn.x, % LoadSavedSearchBtn.y
    Sleep 500
    MouseClick, left, % DocksidedPreset.x, % DocksidedPreset.y
    Sleep 500
    MouseClick, left, % StatusSelect.x, % StatusSelect.y
    Sleep 250
    SendInput, o
    Sleep 250
    SendInput, {Down}
    Sleep 250
    SendInput, {Space}
    Sleep 250
    SendInput, {Enter}
    Sleep 500
    WinWait, Search Results:, , 5
    if (ErrorLevel)
        return
    WinActivate, Search Results:
    WinWaitActive, Search Results:, , 2
    Sleep 100
    SendInput, !{Space}
    Sleep 100
    SendInput, s
    Sleep 100
    SendInput, {Right}
    Sleep 100
    SendInput, {Down}
    MouseMove, 2050, 1025
    Sleep 100
    SendInput, {Enter}
    Sleep 100
    MouseMove, 945, 70
return

!a::  ; Arrived at BSC items
    MouseClick, left, % ClearSearchBtn.x, % ClearSearchBtn.y, 3
    Sleep 250
    MouseClick, left, % LoadSavedSearchBtn.x, % LoadSavedSearchBtn.y
    Sleep 500
    MouseClick, left, % DocksidedPreset.x, % DocksidedPreset.y
    Sleep 500
    MouseClick, left, % StatusSelect.x, % StatusSelect.y
    Sleep 250
    SendInput, a
    Sleep 250
    SendInput, {Space}
    Sleep 250
    SendInput, {Enter}
    Sleep 500
    WinWait, Search Results:, , 5
    if (ErrorLevel)
        return
    WinActivate, Search Results:
    WinWaitActive, Search Results:, , 2
    Sleep 100
    SendInput, !{Space}
    Sleep 100
    SendInput, s
    Sleep 100
    SendInput, {Right}
    Sleep 100
    SendInput, {Down}
    MouseMove, 2050, 1025
    Sleep 100
    SendInput, {Enter}
    Sleep 100
    MouseMove, 945, 70
return

!r::  ; Status selection using "r"
    MouseClick, left, % ClearSearchBtn.x, % ClearSearchBtn.y, 3
    Sleep 250
    MouseClick, left, % LoadSavedSearchBtn.x, % LoadSavedSearchBtn.y
    Sleep 500
    MouseClick, left, % DocksidedPreset.x, % DocksidedPreset.y
    Sleep 500
    MouseClick, left, % StatusSelect.x, % StatusSelect.y
    Sleep 250
    SendInput, r
    Sleep 250
    SendInput, {Down}
    Sleep 250
    SendInput, {Space}
    Sleep 250
    SendInput, {Enter}
    Sleep 500
    WinWait, Search Results:, , 5
    if (ErrorLevel)
        return
    WinActivate, Search Results:
    WinWaitActive, Search Results:, , 2
    Sleep 100
    SendInput, !{Space}
    Sleep 100
    SendInput, s
    Sleep 100
    SendInput, {Right}
    Sleep 100
    SendInput, {Down}
    MouseMove, 2050, 1025
    Sleep 100
    SendInput, {Enter}
    Sleep 100
    MouseMove, 945, 70
return

!p::  ; Pickup from BSC items
    MouseClick, left, % ClearSearchBtn.x, % ClearSearchBtn.y, 3
    Sleep 250
    MouseClick, left, % LoadSavedSearchBtn.x, % LoadSavedSearchBtn.y
    Sleep 500
    MouseClick, left, % DocksidedPreset.x, % DocksidedPreset.y
    Sleep 500
    MouseClick, left, % StatusSelect.x, % StatusSelect.y
    Sleep 250
    SendInput, p
    Sleep 250
    Loop 3
    {
        SendInput, {Down}
        Sleep 150
    }
    SendInput, {Space}
    Sleep 250
    SendInput, {Enter}
    Sleep 500
    WinWait, Search Results:, , 5
    if (ErrorLevel)
        return
    WinActivate, Search Results:
    WinWaitActive, Search Results:, , 2
    Sleep 100
    SendInput, !{Space}
    Sleep 100
    SendInput, s
    Sleep 100
    SendInput, {Right}
    Sleep 100
    SendInput, {Down}
    MouseMove, 2050, 1025
    Sleep 100
    SendInput, {Enter}
    Sleep 100
    MouseMove, 945, 70
return

!s::
    MouseClick, left, % StatusSelect.x, % StatusSelect.y
return
#If
