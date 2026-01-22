#Requires AutoHotkey v1
#NoEnv
#Warn
#SingleInstance, Force
SendMode Input
SetWorkingDir %A_ScriptDir%
SetTitleMatchMode, 2
CoordMode, Mouse, Window

#e::ToggleExplorer()
#!e::Run, explorer.exe
#f::ToggleFirefox()
; Win+Alt+V - Focus/Minimize/Launch VS Code
#!v::ToggleVsCode()
; Ctrl+Win+Alt+V - Focus/Minimize/Position VLC
^#!v::ToggleVlc()

; Win+Alt+R - Launch RS focus helper
#!r::
    Run, C:\Users\Davey\Documents\AutoHotkey\Focus_RS_Window.ahk
    Sleep, 150
    ToolTip, RS Window Focus: Alt+Z
    SetTimer, HideRSTip, -5000
return

HideRSTip:
    ToolTip
return

ToggleExplorer()
{
    ; Cycle Explorer windows if multiple; toggle minimize if single and active; open if none.
    WinGet, idList, List, ahk_class CabinetWClass
    count := idList
    if (count = 0)
    {
        Run, explorer.exe
        return
    }

    WinGet, activeId, ID, A
    activeIsExplorer := WinActive("ahk_class CabinetWClass")

    if (count = 1 && activeIsExplorer)
    {
        WinMinimize, ahk_id %activeId%
        return
    }

    if (!activeIsExplorer)
    {
        WinRestore, ahk_id %idList1%
        WinActivate, ahk_id %idList1%
        WinWaitActive, ahk_id %idList1%,, 1
        return
    }

    nextIndex := 1
    Loop %count%
    {
        idx := A_Index
        thisId := idList%idx%
        if (thisId = activeId)
        {
            nextIndex := (idx = count) ? 1 : (idx + 1)
            break
        }
    }
    nextId := idList%nextIndex%
    WinRestore, ahk_id %nextId%
    WinActivate, ahk_id %nextId%
    WinWaitActive, ahk_id %nextId%,, 1
}

ToggleFirefox()
{
    if WinActive("ahk_exe firefox.exe")
    {
        WinMinimize, ahk_exe firefox.exe
        return
    }
    if WinExist("ahk_exe firefox.exe")
    {
        WinActivate  ; last found window
        WinWaitActive, ahk_exe firefox.exe,, 1
        return
    }
    Run, firefox.exe
}

ToggleVsCode()
{
    if WinActive("ahk_exe Code.exe")
    {
        WinMinimize, ahk_exe Code.exe
        return
    }
    if WinExist("ahk_exe Code.exe")
    {
        WinActivate  ; last found window
        WinWaitActive, ahk_exe Code.exe,, 1
        return
    }
    vsExe := GetVsCodeExe()
    if (vsExe = "code")
        Run, %vsExe%,, Hide
    else
        Run, % """" vsExe """"
}

ToggleVlc()
{
    local candidate := "ahk_exe vlc.exe"

    if WinActive(candidate)
    {
        WinMinimize, %candidate%
        return
    }
    if WinExist(candidate)
    {
        WinRestore, %candidate%
        WinActivate  ; last found window
        WinWaitActive, %candidate%,, 1
    }
    else
    {
        Run, vlc.exe
        WinWaitActive, %candidate%,, 2
    }
    if WinExist(candidate)
    {
        EnsureVlcWindowed(candidate)
        CenterWindowOnMonitor(candidate, 0.5)
    }
}

EnsureVlcWindowed(winTitle)
{
    ; Ensure VLC has a normal resizable captioned window (fixes missing drag bar).
    WinSet, Style, +0xC00000, %winTitle%  ; WS_CAPTION
    WinSet, Style, +0x40000, %winTitle%   ; WS_THICKFRAME
    WinSet, Style, +0x80000, %winTitle%   ; WS_SYSMENU
    WinSet, Style, +0x10000, %winTitle%   ; WS_MAXIMIZEBOX
    WinSet, Style, +0x20000, %winTitle%   ; WS_MINIMIZEBOX
    WinSet, Redraw,, %winTitle%
    ; Exit fullscreen if VLC is in that mode.
    ControlSend,, {Esc}, %winTitle%
}

CenterWindowOnMonitor(winTitle, scale := 0.5)
{
    global monLeft, monTop, monRight, monBottom
    ; Use implicit locals to avoid SysGet scope issues with MonitorWorkArea.
    winX := ""
    winY := ""
    winW := ""
    winH := ""

    WinGetPos, winX, winY, winW, winH, %winTitle%
    if (winW = "" || winH = "")
        return

    centerX := winX + (winW / 2)
    centerY := winY + (winH / 2)

    ; Defaults in case SysGet fails
    monLeft := 0
    monTop := 0
    monRight := A_ScreenWidth
    monBottom := A_ScreenHeight

    SysGet, monitorCount, MonitorCount
    Loop %monitorCount%
    {
        SysGet, mon, MonitorWorkArea, %A_Index%
        if (centerX >= monLeft && centerX <= monRight && centerY >= monTop && centerY <= monBottom)
            break
    }
    ; If not on any monitor, fall back to primary monitor work area.
    if !(centerX >= monLeft && centerX <= monRight && centerY >= monTop && centerY <= monBottom)
        SysGet, mon, MonitorWorkArea, 1

    monW := monRight - monLeft
    monH := monBottom - monTop
    if (monW <= 0 || monH <= 0)
    {
        monLeft := 0
        monTop := 0
        monRight := A_ScreenWidth
        monBottom := A_ScreenHeight
        monW := monRight - monLeft
        monH := monBottom - monTop
    }
    newW := Floor(monW * scale)
    newH := Floor(monH * scale)
    newX := Floor(monLeft + (monW - newW) / 2)
    newY := Floor(monTop + (monH - newH) / 2)

    WinMove, %winTitle%,, %newX%, %newY%, %newW%, %newH%
}

GetVsCodeExe()
{
    local localAppData, path
    EnvGet, localAppData, LOCALAPPDATA
    if (localAppData != "")
    {
        path := localAppData "\Programs\Microsoft VS Code\Code.exe"
        if FileExist(path)
            return path
    }
    path := A_ProgramFiles "\Microsoft VS Code\Code.exe"
    if FileExist(path)
        return path
    path := A_ProgramFiles " (x86)\Microsoft VS Code\Code.exe"
    if FileExist(path)
        return path
    return "code"
}

ToggleFocusOrMinimizeExe(exe)
{
    candidate := "ahk_exe " exe
    if WinActive(candidate)
    {
        WinMinimize, %candidate%
        return
    }
    if WinExist(candidate)
    {
        WinActivate  ; last found window
        WinWaitActive, %candidate%,, 1
    }
}
