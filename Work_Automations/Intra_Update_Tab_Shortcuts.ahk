#Requires AutoHotkey v1
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Warn  ; Enable warnings to assist with detecting common errors.
#SingleInstance, Force  ; Reload without prompt when Esc is pressed.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetTitleMatchMode, 2  ; Partial matches for Intra window names.
CoordMode, Mouse, Window  ; Use window-relative coordinates.

; Scope: Update tab + Search window helpers; abort flag lets Esc/cancel safely bail.
; More Coordinates (ratios for Update window so they scale with size/position)
; Current client size from latest window spy: w=1718, h=1360
baseW := 1718.0
baseH := 1360.0
StatusBar := {rx: 186/baseW, ry: 205/baseH}
PickupFromBSC := {rx: 174/baseW, ry: 328/baseH}
FirstDropdown := {rx: 185/baseW, ry: 225/baseH}
PackageType := {x: 1375, y: 360}    ; absolute for Update window
BSCLocation := {x: 1375, y: 652}    ; absolute for Update window
PrintLabel := {rx: 300/baseW, ry: 120/baseH}
ScanField := {rx: 200/baseW, ry: 245/baseH}

; Borrowed from Intra_Desktop_Search_Shortcuts (Search - General window)
LoadSavedSearchBtn := {x: 120, y: 720}
DocksidedPreset := {x: 200, y: 600}
StatusSelect := {x: 375, y: 160}

abortHotkey := false

^Esc::
    abortHotkey := true
    Reload
return

#If ( WinActive("Intra Desktop Client - Update") || WinActive("Search - General") )
Esc::
    abortHotkey := true
return

!p::
    ResetAbort()
    MouseClick, left, 70, 1345, 2
    Sleep 200
    ClickScaled(StatusBar)
    Sleep 200
    Loop 25 {
        MouseClick, WheelUp
        Sleep 20
    }
    Loop 20 {
        MouseClick, WheelDown
        Sleep 20
    }
    ClickScaled(FirstDropdown)
    Sleep 3000
    if (AbortRequested())
        return
    SendInput, ^f
    Sleep 1500
    MouseClick, left, % LoadSavedSearchBtn.x, % LoadSavedSearchBtn.y
    Sleep 500
    MouseClick, left, % DocksidedPreset.x, % DocksidedPreset.y
    Sleep 750
    MouseClick, left, % StatusSelect.x, % StatusSelect.y
    Sleep 250
    SendInput, a
    Sleep 100
    SendInput, {Space}
    Sleep 100
    SendInput, {Enter}
    ; Resize/position the Search Results window like the search script does.
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

!d::
    MouseClick, left, 70, 1345, 2
    Sleep 200
    ClickScaled(StatusBar)
    Sleep 200
    Loop 25 {
        MouseClick, WheelUp
        Sleep 20
    }
    ClickScaled(FirstDropdown)
    Sleep 200
return

^!d::
    MouseClick, left, 70, 1345, 2
    Sleep 200
    ClickScaled(StatusBar)
    Sleep 200
    Loop 25 {
        MouseClick, WheelUp
        Sleep 20
    }
    Loop 5 {
        MouseClick, WheelDown
        Sleep 20
    }
    ClickScaled(FirstDropdown)
    Sleep 200
return

!c::
    MouseClick, left, 70, 1345, 2
    Sleep 200
    Send, {Enter}
    Sleep 200
    Loop, 2
    {
        SendInput, {Esc}
        Sleep 50
    }
    Sleep 200
    WinGetPos,,, winW, winH, A
    mX := Floor(winW * StatusBar.rx)
    mY := Floor(winH * StatusBar.ry)
    MouseMove, %mX%, %mY%
return
#If

#If WinActive("Intra Desktop Client - Update")
!1::
    MouseClick, left, % PackageType.x, % PackageType.y
return

!2::
    MouseClick, left, % BSCLocation.x, % BSCLocation.y
return

!3::
    ClickScaled(PrintLabel)
    Sleep 150
    ClickScaled(ScanField, 2)
return

!s::
    ClickScaled(StatusBar)
return

!v::
    MouseClick, left, 70, 1345, 2
    Sleep 200
    ClickScaled(StatusBar)
    Sleep 200
    SendInput, v
    Sleep 200
return
#If

ClickScaled(coord, clicks := 1)
{
    local tx, ty
    WinGetPos,,, winW, winH, A
    tx := Floor(winW * coord.rx)
    ty := Floor(winH * coord.ry)
    MouseClick, left, %tx%, %ty%, clicks
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
