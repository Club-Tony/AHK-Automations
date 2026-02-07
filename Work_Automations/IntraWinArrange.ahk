#Requires AutoHotkey v1
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Warn  ; Enable warnings to assist with detecting common errors.
#SingleInstance, Force  ; Reload without prompt when Esc is pressed.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetTitleMatchMode, 2  ; Partial matches for Intra window names.

; Startup tooltip:
; - Uses a separate tooltip ID (20) to avoid interfering with normal status tooltips.
; - Auto-hides after 30s and also hides on the first new key press.
; - Suppressible via ReloadAll by passing /SkipReadyTooltip.
readyTooltipActive := false
readyTooltipKeysDown := {}
skipReadyTooltip := false
for _, arg in A_Args
{
    if (arg = "/SkipReadyTooltip")
    {
        skipReadyTooltip := true
        break
    }
}

if (!skipReadyTooltip)
{
    readyTooltipActive := true
    ToolTip, Intra Window Alignment Script Ready: Open 3 Desktop Clients - (Ctrl + Alt + W), , , 20
    SetTimer, HideReadyTooltip, -30000
    SetTimer, StartReadyTooltipDismissCheck, -1
}

assignTitle := "Intra Desktop Client - Assign Recip"
updateTitle := "Intra Desktop Client - Update"
pickupTitle := "Intra Desktop Client - Pickup"
loginTitle := "Desktop Client Login"

; Automatic tracking of login windows in order of appearance
trackedLoginWindows := []
SetTimer, TrackLoginWindows, 500

; Window coordinates taken from existing Intra scripts/window spy for consistency.
assignPos := {x: -7, y: 0, w: 1322, h: 1399}      ; from latest Assign Recip window spy
updatePos := {x: 1713, y: 0, w: 1734, h: 1399}    ; from latest Update window spy
assignScanPos := {x: 200, y: 245}                  ; scan field (Assign Recip)

^!w::  ; Align Assign/Update/Pickup windows
    ArrangeIntraWindows()
return

^!0::  ; Clear tracked login windows
    trackedLoginWindows := []
    ShowTimedTooltip("Login tracking cleared", 1500)
return

StartReadyTooltipDismissCheck:
    if (!readyTooltipActive)
        return
    readyTooltipKeysDown := {}
    Loop, 256
    {
        key := A_Index
        ; Skip mouse buttons.
        if (key = 1 || key = 2 || key = 4 || key = 5 || key = 6)
            continue
        if (GetKeyState(Format("vk{:02X}", key), "P"))
            readyTooltipKeysDown[key] := true
    }
    SetTimer, CheckReadyTooltipDismiss, 50
return

CheckReadyTooltipDismiss:
    if (!readyTooltipActive)
        return
    Loop, 256
    {
        key := A_Index
        ; Skip mouse buttons.
        if (key = 1 || key = 2 || key = 4 || key = 5 || key = 6)
            continue
        isPressed := GetKeyState(Format("vk{:02X}", key), "P")
        if (isPressed)
        {
            if (!readyTooltipKeysDown.HasKey(key))
            {
                Gosub, HideReadyTooltip
                return
            }
        }
        else if (readyTooltipKeysDown.HasKey(key))
            readyTooltipKeysDown.Delete(key)
    }
return

HideReadyTooltip:
    SetTimer, HideReadyTooltip, Off
    SetTimer, StartReadyTooltipDismissCheck, Off
    SetTimer, CheckReadyTooltipDismiss, Off
    readyTooltipActive := false
    readyTooltipKeysDown := {}
    ToolTip, , , , 20
return

TrackLoginWindows:
    TrackLoginWindowsFunc()
return

TrackLoginWindowsFunc()
{
    global loginTitle, trackedLoginWindows

    ; Get current login windows
    currentHwnds := {}
    WinGet, winList, List, %loginTitle%
    Loop, %winList%
    {
        hwnd := winList%A_Index%
        currentHwnds[hwnd] := true
    }

    ; Remove any tracked windows that no longer exist as login windows
    ; (they've been logged in and transformed)
    newTracked := []
    for i, info in trackedLoginWindows
    {
        if (currentHwnds[info.hwnd])
            newTracked.Push(info)
    }
    trackedLoginWindows := newTracked

    ; Add any new login windows we haven't seen before
    alreadyTracked := {}
    for i, info in trackedLoginWindows
        alreadyTracked[info.hwnd] := true

    Loop, %winList%
    {
        hwnd := winList%A_Index%
        if (!alreadyTracked[hwnd])
        {
            WinGet, pid, PID, ahk_id %hwnd%
            trackedLoginWindows.Push({hwnd: hwnd, pid: pid})
        }
    }
}

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
    global trackedLoginWindows, updatePos, updateTitle, pickupTitle, assignPos, assignScanPos, assignTitle

    prevCoordMode := A_CoordModeMouse
    CoordMode, Mouse, Window

    ; Get current Assign windows with their PIDs
    winInfo := []
    WinGet, winList, List, %title%
    Loop, %winList%
    {
        thisId := winList%A_Index%
        WinGet, pid, PID, ahk_id %thisId%
        winInfo.Push({id: thisId, pid: pid})
    }

    ; Try to match by tracked login order (using PID)
    useTrackedOrder := false
    if (trackedLoginWindows.MaxIndex() >= 3)
    {
        ordered := []
        for i, tracked in trackedLoginWindows
        {
            for j, win in winInfo
            {
                if (win.pid = tracked.pid)
                {
                    ordered.Push(win)
                    break
                }
            }
        }
        if (ordered.MaxIndex() = 3)
        {
            winInfo := ordered
            useTrackedOrder := true
        }
    }

    ; Fallback to process creation time sorting if tracking didn't work
    if (!useTrackedOrder)
    {
        for i, win in winInfo
            win.time := GetProcessCreationTimeForWindow(win.id)
        SortWindowsByCreation(winInfo)
    }

    for index, info in winInfo
    {
        hwnd := info.id
        WinActivate, ahk_id %hwnd%
        WinWaitActive, ahk_id %hwnd%,, 1
        if (index = 1)
        {
            ; First opened → Update (RIGHT)
            WinMove, ahk_id %hwnd%,, % updatePos.x, % updatePos.y, % updatePos.w, % updatePos.h
            Sleep 100
            MouseClick, left, 65, 75
            WinWaitActive, %updateTitle%,, 2
        }
        else if (index = 2)
        {
            ; Second opened → Pickup (MAXIMIZED)
            WinGet, isMax, MinMax, ahk_id %hwnd%
            if (isMax != 1)
                WinMaximize, ahk_id %hwnd%
            MouseClick, left, 290, 75, 2
            WinWaitActive, %pickupTitle%,, 2
        }
        else if (index = 3)
        {
            ; Third opened → Assign (LEFT)
            WinRestore, ahk_id %hwnd%
            WinMove, ahk_id %hwnd%,, % assignPos.x, % assignPos.y, % assignPos.w, % assignPos.h
            Sleep 100
            ; Stay on Assign tab (no click needed)
        }
    }

    ; Focus back on Assign for scanning
    Sleep 250
    WinActivate, %assignTitle%
    MouseMove, % assignScanPos.x, % assignScanPos.y
    ShowTimedTooltip("Intra Desktop switched and aligned", 3000)

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
