#Requires AutoHotkey v1
#NoEnv
#Warn
#SingleInstance, Force
#InstallKeybdHook
#UseHook
#InputLevel 1  ; Only respond to physical keypresses, improves game compatibility
SendMode Input
SetWorkingDir %A_ScriptDir%

; Master reload script - reloads all currently running AHK scripts
; Keybind: Ctrl+Esc
; Dynamically detects running scripts from managed directories

reloadTooltipActive := false
reloadedScriptsList := []
tooltipMouseX := 0
tooltipMouseY := 0
reloadElapsedText := ""
allowInputDismiss := false
reloadTooltipShowingList := false

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

; Legacy script-name aliases after filename cleanup.
ScriptRenameMap := {}
ScriptRenameMap["Right Click.ahk"] := "RightClick.ahk"
ScriptRenameMap["Right_Click.ahk"] := "RightClick.ahk"
ScriptRenameMap["Right-Click.ahk"] := "RightClick.ahk"
ScriptRenameMap["Launcher Script (Keybinds).ahk"] := "Launcher_Script(WorkAutos).ahk"
ScriptRenameMap["Launcher_Script_(Keybinds).ahk"] := "Launcher_Script(WorkAutos).ahk"
ScriptRenameMap["Launcher_Script_Other.ahk"] := "Launcher_Script.ahk"
ScriptRenameMap["Alt+Q Map to Alt+F4.ahk"] := "Alt+Q=Alt+F4.ahk"
ScriptRenameMap["Alt+Q_Map_to_Alt+F4.ahk"] := "Alt+Q=Alt+F4.ahk"
ScriptRenameMap["Ctrl+Home key = Sleep.ahk"] := "Ctrl+Home=Sleep.ahk"
ScriptRenameMap["Ctrl+Home_key_=_Sleep.ahk"] := "Ctrl+Home=Sleep.ahk"
ScriptRenameMap["RK71_Key_Fixes.ahk"] := "RK71KeyFixes.ahk"

$^Esc::
    startTick := A_TickCount
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
            resolvedPath := ResolveManagedScriptPath(scriptPath)
            if (resolvedPath = "")
                continue
            SplitPath, resolvedPath, scriptName
            WinClose, ahk_id %hwnd%
            Sleep 50
            if (scriptName = "IntraWinArrange.ahk")
                Run, "%A_AhkPath%" "%resolvedPath%" /SkipReadyTooltip
            else
                Run, "%resolvedPath%"
            reloadedScripts.Push(scriptName)
        }
    }

    DetectHiddenWindows, Off
    Sleep 1000
    elapsedMs := A_TickCount - startTick
    reloadElapsedText := "Reload time: " Round(elapsedMs / 1000.0, 2) "s"

    ; Show tooltip with results
    reloadedScriptsList := reloadedScripts
    if (reloadedScripts.Length() > 0)
    {
        msg := "Reloaded " reloadedScripts.Length() " script(s)`n" reloadElapsedText "`nPress T to view list"
        ToolTip, %msg%
        reloadTooltipActive := true
        allowInputDismiss := false
        reloadTooltipShowingList := false
        MouseGetPos, tooltipMouseX, tooltipMouseY
        SetTimer, StartTooltipDismissCheck, -1
        SetTimer, ClearReloadTooltip, -5000
    }
    else
    {
        ToolTip, No managed scripts were running`n%reloadElapsedText%
        reloadTooltipActive := true
        allowInputDismiss := false
        reloadTooltipShowingList := false
        MouseGetPos, tooltipMouseX, tooltipMouseY
        SetTimer, StartTooltipDismissCheck, -1
        SetTimer, ClearReloadTooltip, -3000
    }
return

#If reloadTooltipActive
t::
    ; If the expanded list is showing, treat "t" as normal typing: dismiss tooltip and pass the key through.
    if (reloadTooltipShowingList)
    {
        Gosub, ClearReloadTooltip
        SendInput, t
        return
    }
    SetTimer, ClearReloadTooltip, Off
    SetTimer, StartTooltipDismissCheck, Off
    SetTimer, CheckTooltipDismiss, Off
    msg := "Reloaded " reloadedScriptsList.Length() " script(s):`n" reloadElapsedText "`n"
    for i, name in reloadedScriptsList
        msg .= "  " name "`n"
    msg .= "`nPress Esc to close"
    ToolTip, %msg%
    allowInputDismiss := false
    reloadTooltipShowingList := true
    MouseGetPos, tooltipMouseX, tooltipMouseY
    SetTimer, StartTooltipDismissCheck, -1
    SetTimer, ClearReloadTooltip, -30000
return

Esc::
    Gosub, ClearReloadTooltip
return
#If

StartTooltipDismissCheck:
    ; Allow dismiss via mouse-move or key press after tooltip is shown.
    allowInputDismiss := true
    SetTimer, CheckTooltipDismiss, 50
return

CheckTooltipDismiss:
    if (!allowInputDismiss || !reloadTooltipActive)
        return
    ; Check for any key press (except T and modifier keys)
    Loop, 256
    {
        key := A_Index
        ; Skip T (84), Esc (27), Ctrl (17), Shift (16), Alt (18), LWin (91), RWin (92)
        ; Skip mouse buttons: LButton (1), RButton (2), MButton (4), XButton1 (5), XButton2 (6)
        if (key = 84 || key = 27 || key = 16 || key = 17 || key = 18 || key = 91 || key = 92
            || key = 1 || key = 2 || key = 4 || key = 5 || key = 6)
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
    SetTimer, StartTooltipDismissCheck, Off
    SetTimer, CheckTooltipDismiss, Off
    reloadTooltipActive := false
    allowInputDismiss := false
    reloadTooltipShowingList := false
    ToolTip
return

ResolveManagedScriptPath(scriptPath)
{
    global ScriptRenameMap
    if (FileExist(scriptPath))
        return scriptPath
    SplitPath, scriptPath, fileName, fileDir
    if (!ScriptRenameMap.HasKey(fileName))
        return ""
    candidate := fileDir "\" ScriptRenameMap[fileName]
    if (FileExist(candidate))
        return candidate
    return ""
}
