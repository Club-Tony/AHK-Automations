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
