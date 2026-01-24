#Requires AutoHotkey v1
#NoEnv
#Warn
#SingleInstance, Force
#InstallKeybdHook
SendMode Input
SetWorkingDir %A_ScriptDir%

; Master reload script - reloads all currently running AHK scripts
; Keybind: Ctrl+Esc
; Dynamically detects running scripts from managed directories

reloadTooltipActive := false
reloadedScriptsList := []
tooltipMouseX := 0
tooltipMouseY := 0

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
        WinGetTitle, winTitle, ahk_id %hwnd%

        ; Window title format: "C:\path\script.ahk - AutoHotkey v1.x.x"
        ; Extract just the script path by removing " - AutoHotkey..." suffix
        scriptPath := winTitle
        ahkPos := InStr(winTitle, " - AutoHotkey")
        if (ahkPos > 0)
            scriptPath := SubStr(winTitle, 1, ahkPos - 1)

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
    reloadedScriptsList := reloadedScripts
    if (reloadedScripts.Length() > 0)
    {
        msg := "Reloaded " reloadedScripts.Length() " script(s)`nPress T to view list"
        ToolTip, %msg%
        reloadTooltipActive := true
        MouseGetPos, tooltipMouseX, tooltipMouseY
        SetTimer, ClearReloadTooltip, -5000
        SetTimer, CheckTooltipDismiss, 50
    }
    else
    {
        ToolTip, No managed scripts were running
        reloadTooltipActive := true
        MouseGetPos, tooltipMouseX, tooltipMouseY
        SetTimer, ClearReloadTooltip, -1500
        SetTimer, CheckTooltipDismiss, 50
    }
return

#If reloadTooltipActive
t::
    SetTimer, ClearReloadTooltip, Off
    SetTimer, CheckTooltipDismiss, Off
    msg := "Reloaded " reloadedScriptsList.Length() " script(s):`n"
    for i, name in reloadedScriptsList
        msg .= "  " name "`n"
    msg .= "`nPress Esc to close"
    ToolTip, %msg%
    MouseGetPos, tooltipMouseX, tooltipMouseY
    SetTimer, ClearReloadTooltip, -30000
    SetTimer, CheckTooltipDismiss, 50
return

Esc::
    Gosub, ClearReloadTooltip
return
#If

CheckTooltipDismiss:
    MouseGetPos, currentX, currentY
    if (Abs(currentX - tooltipMouseX) > 20 || Abs(currentY - tooltipMouseY) > 20)
    {
        Gosub, ClearReloadTooltip
        return
    }
    ; Check for any key press (except T which is handled separately)
    Loop, 256
    {
        key := A_Index
        if (key = 84)  ; Skip T
            continue
        if (GetKeyState(Format("vk{:02X}", key), "P"))
        {
            Gosub, ClearReloadTooltip
            return
        }
    }
return

ClearReloadTooltip:
    SetTimer, ClearReloadTooltip, Off
    SetTimer, CheckTooltipDismiss, Off
    reloadTooltipActive := false
    ToolTip
return
