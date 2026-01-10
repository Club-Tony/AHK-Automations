#Requires AutoHotkey v1
#NoEnv ; Prevents Unnecessary Environment Variable lookup
#Warn ; Warn All (All Warnings Enabled)
#SingleInstance, Force ; Removes script already open warning when reloading scripts
SendMode Input
SetWorkingDir, %A_ScriptDir%

^Esc::Reload

; ; Keybind: Ctrl+Alt+R Opens RS-related scripts
^!r:: 
    Run, C:\Users\Davey\Documents\AutoHotkey\Focus_RS_Window.ahk ; Keybind: Alt+Z
    Sleep 150
    Tooltip, RS Window Focus: Alt+Z
    Sleep 5000
    Tooltip
Return

^+!c::
    Run, "C:\Users\Davey\Documents\GitHub\Repositories\AHK-Automations\Other_Automations\Coordinate Capture Helper\Coord_Capture.ahk"
    ToolTip, Coord Helper: Alt+C to capture, Alt+I for more info, Alt+G toggle bg
    SetTimer, HideCoordTip, -4000
Return

; Launch: Window Spy ; Keybind: Ctrl+Shift+Alt+W
^+!w::
    ToggleWindowSpy()
Return

^+!o::
    Run, "C:\Users\Davey\Documents\GitHub\Repositories\AHK-Automations\Other_Automations\Coordinate Capture Helper\coord.txt"
Return

HideCoordTip:
    ToolTip
Return

ToggleWindowSpy()
{
    DetectHiddenWindows, On
    hwnd := WinExist("Window Spy")
    DetectHiddenWindows, Off
    if (hwnd)
    {
        WinClose, ahk_id %hwnd%
        return
    }

    spyExe := ""
    spyPath := ""
    if (GetWindowSpyLaunch(spyExe, spyPath))
        Run, "%spyExe%" "%spyPath%"
    else
        Run, WindowSpy.ahk
}

GetWindowSpyLaunch(ByRef spyExe, ByRef spyPath)
{
    v2Base := "C:\Program Files\AutoHotkey"
    v2Spy := v2Base "\WindowSpy.ahk"
    if (FileExist(v2Spy))
    {
        if (FileExist(v2Base "\v2\AutoHotkey64.exe"))
            spyExe := v2Base "\v2\AutoHotkey64.exe"
        else if (FileExist(v2Base "\v2\AutoHotkey32.exe"))
            spyExe := v2Base "\v2\AutoHotkey32.exe"
        else if (FileExist(v2Base "\v2\AutoHotkey.exe"))
            spyExe := v2Base "\v2\AutoHotkey.exe"
        if (spyExe != "")
        {
            spyPath := v2Spy
            return true
        }
    }

    SplitPath, A_AhkPath,, ahkDir
    candidate := ahkDir "\WindowSpy.ahk"
    if (FileExist(candidate))
    {
        spyExe := A_AhkPath
        spyPath := candidate
        return true
    }
    return false
}
