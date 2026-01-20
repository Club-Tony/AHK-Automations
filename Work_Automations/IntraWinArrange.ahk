#Requires AutoHotkey v1
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Warn  ; Enable warnings to assist with detecting common errors.
#SingleInstance, Force  ; Reload without prompt when Esc is pressed.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetTitleMatchMode, 2  ; Partial matches for Intra window names.

assignTitle := "Intra Desktop Client - Assign Recip"
updateTitle := "Intra Desktop Client - Update"
pickupTitle := "Intra Desktop Client - Pickup"

; Window coordinates taken from existing Intra scripts/window spy for consistency.
assignPos := {x: -7, y: 0, w: 1322, h: 1399}      ; from latest Assign Recip window spy
updatePos := {x: 1713, y: 0, w: 1734, h: 1399}    ; from latest Update window spy
assignScanPos := {x: 200, y: 245}                  ; scan field (Assign Recip)

^!w::  ; Align Assign/Update/Pickup windows
    ArrangeIntraWindows()
return

ArrangeIntraWindows()
{
    global assignTitle, updateTitle, pickupTitle, assignPos, updatePos, assignScanPos

    assignCount := GetWindowCount(assignTitle)
    updateCount := GetWindowCount(updateTitle)
    pickupCount := GetWindowCount(pickupTitle)
    totalCount := assignCount + updateCount + pickupCount

    if (totalCount < 3)
    {
        ShowTimedTooltip("Make sure 3 Intra Desktop Client instances are open first", 3000)
        return
    }

    ; Special-case: three Assign windows open with no other Intra windows.
    if (totalCount = 3 && assignCount = 3)
    {
        CycleAssignWindows(assignTitle)
        return
    }

    if !(assignCount && updateCount && pickupCount)
    {
        ShowTimedTooltip("Make sure Assign, Update, and Pickup windows are open", 3000)
        return
    }

    ; 1) Maximize Pickup.
    WinActivate, %pickupTitle%
    WinWaitActive, %pickupTitle%,, 1
    WinRestore, %pickupTitle%
    Sleep 50
    WinMaximize, %pickupTitle%
    Sleep 100

    ; 2) Put Update on the right (window spy coordinates).
    WinActivate, %updateTitle%
    WinWaitActive, %updateTitle%,, 1
    WinRestore, %updateTitle%
    Sleep 75
    WinMove, %updateTitle%,, % updatePos.x, % updatePos.y, % updatePos.w, % updatePos.h
    Sleep 100

    ; 3) Put Assign Recip on the left (same sizing used by prior scripts).
    WinActivate, %assignTitle%
    WinWaitActive, %assignTitle%,, 1
    WinRestore, %assignTitle%
    Sleep 75
    WinMove, %assignTitle%,, % assignPos.x, % assignPos.y, % assignPos.w, % assignPos.h

    ; 4) Park cursor on Assign scan field, then confirm completion.
    prevCoordMode := A_CoordModeMouse
    CoordMode, Mouse, Window
    MouseMove, % assignScanPos.x, % assignScanPos.y
    CoordMode, Mouse, %prevCoordMode%
    ShowTimedTooltip("Windows aligned.", 3000)
}

GetWindowCount(title)
{
    WinGet, winList, List, %title%
    return winList
}

CycleAssignWindows(title)
{
    global updatePos, updateTitle, pickupTitle, assignPos, assignScanPos

    prevCoordMode := A_CoordModeMouse
    CoordMode, Mouse, Window

    winInfo := []
    WinGet, winList, List, %title%
    Loop, %winList%
    {
        thisId := winList%A_Index%
        creation := GetProcessCreationTimeForWindow(thisId)
        winInfo.Push({id: thisId, time: creation})
    }

    SortWindowsByCreation(winInfo)

    for index, info in winInfo
    {
        hwnd := info.id
        WinActivate, ahk_id %hwnd%
        WinWaitActive, ahk_id %hwnd%,, 1
        if (index = 1)
        {
            ; Match Update window sizing/position, then switch to Update tab.
            WinMove, ahk_id %hwnd%,, % updatePos.x, % updatePos.y, % updatePos.w, % updatePos.h
            Sleep 100
            MouseClick, left, 65, 75
            WinWaitActive, %updateTitle%,, 2
            Sleep 1000
        }
        else if (index = 2)
        {
            ; Ensure maximized, then switch to Pickup tab.
            WinGet, isMax, MinMax, ahk_id %hwnd%
            if (isMax != 1)
                WinMaximize, ahk_id %hwnd%
            MouseClick, left, 290, 75, 2
            WinWaitActive, %pickupTitle%,, 2
            Sleep 1000
        }
        else if (index = 3)
        {
            ; Return to Assign sizing/position.
            WinRestore, ahk_id %hwnd%
            WinMove, ahk_id %hwnd%,, % assignPos.x, % assignPos.y, % assignPos.w, % assignPos.h
            Sleep 100
            MouseMove, % assignScanPos.x, % assignScanPos.y
            ShowTimedTooltip("Intra Desktop switched and aligned", 3000)
        }
    }

    CoordMode, Mouse, %prevCoordMode%
}

GetProcessCreationTimeForWindow(hwnd)
{
    WinGet, pid, PID, ahk_id %hwnd%
    return GetProcessCreationTime(pid)
}

GetProcessCreationTime(pid)
{
    PROCESS_QUERY_LIMITED_INFORMATION := 0x1000
    creation := 0, exitT := 0, kernelT := 0, userT := 0
    hProc := DllCall("OpenProcess", "UInt", PROCESS_QUERY_LIMITED_INFORMATION, "Int", False, "UInt", pid, "Ptr")
    if (!hProc)
        return 0

    success := DllCall("GetProcessTimes", "Ptr", hProc, "Int64P", creation, "Int64P", exitT, "Int64P", kernelT, "Int64P", userT)
    DllCall("CloseHandle", "Ptr", hProc)

    return success ? creation : 0
}

SortWindowsByCreation(ByRef winInfo)
{
    count := winInfo.MaxIndex()
    if (!count)
        return

    ; Simple stable bubble sort; keeps z-order when creation times tie.
    Loop, % count - 1
    {
        swapped := false
        Loop, % count - A_Index
        {
            i := A_Index
            j := i + 1
            if (winInfo[i].time > winInfo[j].time)
            {
                temp := winInfo[i]
                winInfo[i] := winInfo[j]
                winInfo[j] := temp
                swapped := true
            }
        }
        if (!swapped)
            break
    }
}

ShowTimedTooltip(msg, duration := 3000)
{
    ToolTip, %msg%
    Sleep %duration%
    ToolTip
}
