#Requires AutoHotkey v1
#NoEnv
#Warn
#SingleInstance, Force
SendMode Input
SetWorkingDir %A_ScriptDir%

; Master reload script - reloads all currently running AHK scripts
; Keybind: Ctrl+Esc
; Dynamically detects running scripts from managed directories

; Directories to scan for running scripts
ManagedDirs := []
; Home machine (Davey)
ManagedDirs.Push("C:\Users\Davey\Documents\GitHub\Repositories\AHK-Automations\Work_Automations")
ManagedDirs.Push("C:\Users\Davey\Documents\GitHub\Repositories\AHK-Automations\Other_Automations")
ManagedDirs.Push("C:\Users\Davey\Documents\GitHub\Repositories\Macros-Script")
; Work machine (daveyuan)
ManagedDirs.Push("C:\Users\daveyuan\Documents\GitHub\Repositories\AHK-Automations\Work_Automations")
ManagedDirs.Push("C:\Users\daveyuan\Documents\GitHub\Repositories\AHK-Automations\Other_Automations")
ManagedDirs.Push("C:\Users\daveyuan\Documents\GitHub\Repositories\Macros-Script")

^Esc::
    reloadedScripts := []
    DetectHiddenWindows, On

    ; Get list of all running AHK scripts
    WinGet, ahkWindows, List, ahk_class AutoHotkey

    Loop, %ahkWindows%
    {
        hwnd := ahkWindows%A_Index%
        WinGetTitle, scriptPath, ahk_id %hwnd%

        ; Skip this script (ReloadAll.ahk)
        if (scriptPath = A_ScriptFullPath)
            continue

        ; Check if this script is in one of our managed directories
        isManaged := false
        for idx, dir in ManagedDirs
        {
            if (InStr(scriptPath, dir) = 1)  ; Script path starts with managed dir
            {
                isManaged := true
                break
            }
        }

        if (isManaged)
        {
            SplitPath, scriptPath, scriptName
            WinClose, ahk_id %hwnd%
            Sleep 50
            Run, "%scriptPath%"
            reloadedScripts.Push(scriptName)
        }
    }

    DetectHiddenWindows, Off

    ; Show tooltip with results
    if (reloadedScripts.Length() > 0)
    {
        msg := "Reloaded " reloadedScripts.Length() " script(s):`n"
        for i, name in reloadedScripts
            msg .= "  " name "`n"
        ToolTip, %msg%
        SetTimer, ClearReloadTooltip, -2000
    }
    else
    {
        ToolTip, No managed scripts were running
        SetTimer, ClearReloadTooltip, -1500
    }

    ; Reload this script last
    Reload
return

ClearReloadTooltip:
    ToolTip
return
