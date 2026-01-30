#Requires AutoHotkey v1
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Warn  ; Enable warnings to assist with detecting common errors.
#SingleInstance, Force  ; Reload without prompt when Esc is pressed.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetTitleMatchMode, 2  ; Allow partial title matches for Intra windows.
SetWinDelay, 0  ; Speed up window commands

assignTitle := "Intra Desktop Client - Assign Recip"
updateTitle := "Intra Desktop Client - Update"
pickupTitle := "Intra Desktop Client - Pickup"
updatePos := {x: 1713, y: 0, w: 1734, h: 1399}    ; Match Ctrl+Alt+W sizing for Update
worldShipTitle := "UPS WorldShip"
worldShipPos := {x: 1204, y: 711, w: 814, h: 688} ; Window Spy sizing for UPS WorldShip
qvnTitle := "Quantum View Notify Recipients"
exportedReportTitle := "ExportedReport.pdf"
browserExes := ["firefox.exe", "chrome.exe", "msedge.exe"]
vsCodePos := {x: 0, y: 0, w: 1724, h: 1392}

#a::ToggleFocusOrMinimize(assignTitle)
#u::ToggleFocusOrMinimize(updateTitle)
#p::ToggleFocusOrMinimize(pickupTitle)
#f::ToggleFirefox()
#s::ToggleFocusOrMinimizeExe("slack.exe")
#w::ToggleWorldShip()
#o::ToggleOutlookPwa()
#e::ToggleExplorer()
#!e::OpenNewExplorer()

#i::ToggleIntraGroup()
#!v::ToggleVsCode()

#!m::
    WinMinimizeAll
    Sleep 150
    FocusExeWindow("firefox.exe")
    Sleep 100
    FocusOutlookPwa()
    Sleep 100
    FocusExeWindow("slack.exe")
return

#If (WinActive(assignTitle) || WinActive(updateTitle) || IsExportedReportActive())
$!Tab::
    ; Only intercept Alt+Tab while in Assign/Update; otherwise let other scripts/OS handle it.
    Send, {Alt up}
    if WinActive(assignTitle)
    {
        pdfWin := GetExportedReportWindow()
        if (pdfWin != "")
        {
            WinActivate, %pdfWin%
            WinWaitActive, %pdfWin%,, 1
        }
        else if WinExist(updateTitle)
        {
            WinActivate, %updateTitle%
            WinWaitActive, %updateTitle%,, 1
        }
        else
        {
            Send, !{Tab}
        }
    }
    else if (IsExportedReportActive())
    {
        if WinExist(assignTitle)
        {
            WinActivate, %assignTitle%
            WinWaitActive, %assignTitle%,, 1
        }
        else
        {
            Send, !{Tab}
        }
    }
    else if WinActive(updateTitle)
    {
        if WinExist(assignTitle)
        {
            WinActivate, %assignTitle%
            WinWaitActive, %assignTitle%,, 1
        }
        else
        {
            Send, !{Tab}
        }
    }
    else
    {
        Send, !{Tab}
    }
return
#If

GetExportedReportWindow()
{
    global exportedReportTitle, browserExes
    Loop % browserExes.Length()
    {
        candidate := exportedReportTitle " ahk_exe " browserExes[A_Index]
        if (WinExist(candidate))
            return candidate
    }
    return ""
}

IsExportedReportActive()
{
    win := GetExportedReportWindow()
    return (win != "" && WinActive(win))
}

FocusWindow(title)
{
    if WinExist(title)
    {
        WinActivate, %title%
        WinWaitActive, %title%,, 1
    }
}

ToggleFocusOrMinimize(title)
{
    global pickupTitle, updateTitle
    if WinActive(title)
    {
        WinMinimize, %title%
    }
    else if WinExist(title)
    {
        if (title = updateTitle)
            EnsureUpdatePlacement()
        WinActivate, %title%
        WinWaitActive, %title%,, 1
        if (title = pickupTitle)
            WinMaximize, %title%
    }
}

MinimizeIfExists(title)
{
    if WinExist(title)
        WinMinimize, %title%
}

ActivateFast(title, timeout := 0.35)
{
    WinActivate, %title%
    WinWaitActive, %title%,, %timeout%
}

EnsureUpdatePlacement()
{
    global updateTitle, updatePos
    if !WinExist(updateTitle)
        return
    WinRestore, %updateTitle%
    WinMove, %updateTitle%,, % updatePos.x, % updatePos.y, % updatePos.w, % updatePos.h
}

EnsureWorldShipPlacement()
{
    global worldShipTitle, worldShipPos
    if !WinExist(worldShipTitle)
        return
    WinRestore, %worldShipTitle%
    WinMove, %worldShipTitle%,, % worldShipPos.x, % worldShipPos.y, % worldShipPos.w, % worldShipPos.h
}

EnsureVsCodePlacement()
{
    global vsCodePos
    if !WinExist("ahk_exe Code.exe")
        return
    WinRestore, ahk_exe Code.exe
    WinMove, ahk_exe Code.exe,, % vsCodePos.x, % vsCodePos.y, % vsCodePos.w, % vsCodePos.h
}

ToggleIntraGroup()
{
    global assignTitle, updateTitle, pickupTitle
    if (WinActive(assignTitle) || WinActive(updateTitle) || WinActive(pickupTitle))
    {
        MinimizeIfExists(assignTitle)
        MinimizeIfExists(updateTitle)
        MinimizeIfExists(pickupTitle)
        return
    }

    titles := [assignTitle, updateTitle, pickupTitle]
    primary := ""
    assignExists := WinExist(assignTitle)
    updateExists := WinExist(updateTitle)
    pickupExists := WinExist(pickupTitle)

    ; When all three are present (typical minimize/restore), bring them up in a predictable order.
    if (assignExists && updateExists && pickupExists)
    {
        WinRestore, %pickupTitle%
        WinMaximize, %pickupTitle%
        ActivateFast(pickupTitle)

        EnsureUpdatePlacement()
        ActivateFast(updateTitle)

        WinRestore, %assignTitle%
        ActivateFast(assignTitle)
        return
    }

    Loop % titles.Length()
    {
        t := titles[A_Index]
        if (WinExist(t))
        {
            if (t = updateTitle)
            {
                EnsureUpdatePlacement()
            }
            else
            {
                WinRestore, %t%
                if (t = pickupTitle)
                    WinMaximize, %t%
            }
            primary := (primary = "") ? t : primary
        }
    }
    if (updateExists)
        EnsureUpdatePlacement()

    if (assignExists)
    {
        ActivateFast(assignTitle)
    }
    else if (updateExists)
    {
        ActivateFast(updateTitle)
    }
    else if (primary != "")
    {
        ActivateFast(primary)
        if (primary = pickupTitle)
            WinMaximize, %primary%
    }
}

FocusExeWindow(exe)
{
    if WinExist("ahk_exe " exe)
    {
        WinActivate  ; activates last found window
        WinWaitActive, ahk_exe %exe%,, 1
    }
}

ToggleFocusOrMinimizeExe(exe)
{
    candidate := "ahk_exe " exe
    if WinActive(candidate)
    {
        WinMinimize, %candidate%
    }
    else if WinExist(candidate)
    {
        WinActivate  ; last found window
        WinWaitActive, %candidate%,, 1
    }
}

ToggleFirefox()
{
    ; Cycle Firefox windows if multiple; toggle minimize if single and active; open if none.
    WinGet, idList, List, ahk_exe firefox.exe
    count := idList
    if (count = 0)
    {
        Run, firefox.exe
        return
    }

    WinGet, activeId, ID, A
    activeIsFirefox := WinActive("ahk_exe firefox.exe")

    if (count = 1 && activeIsFirefox)
    {
        WinMinimize, ahk_id %activeId%
        return
    }

    if (!activeIsFirefox)
    {
        WinRestore, ahk_id %idList1%
        WinActivate, ahk_id %idList1%
        WinWaitActive, ahk_id %idList1%,, 1
        EnsureFirefoxWindow()
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
    EnsureFirefoxWindow()
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
        EnsureVsCodePlacement()
        return
    }
    vsExe := GetVsCodeExe()
    if (vsExe = "code")
        Run, %vsExe%,, Hide
    else
        Run, % """" vsExe """"
    WinWaitActive, ahk_exe Code.exe,, 2
    EnsureVsCodePlacement()
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

EnsureFirefoxWindow()
{
    ; Align Firefox to the standard Intra size/position used by DSRF scripts.
    WinMove, ahk_exe firefox.exe,, 1917, 0, 1530, 1399
}

FocusOutlookPwa()
{
    ; Try common PWA hosts first (title starts with "Outlook (PWA"), then native Outlook as fallback.
    candidates := ["Outlook (PWA) ahk_exe msedgewebview2.exe"
                 , "Outlook (PWA) ahk_exe msedge.exe"
                 , "Outlook (PWA) ahk_exe chrome.exe"
                 , "Outlook ahk_exe outlook.exe"]
    Loop % candidates.Length()
    {
        candidate := candidates[A_Index]
        if (WinExist(candidate))
        {
            WinActivate  ; last found window
            WinWaitActive, %candidate%,, 1
            return
        }
    }
}

ToggleOutlookPwa()
{
    candidates := ["Outlook (PWA) ahk_exe msedgewebview2.exe"
                 , "Outlook (PWA) ahk_exe msedge.exe"
                 , "Outlook (PWA) ahk_exe chrome.exe"
                 , "Outlook ahk_exe outlook.exe"]

    Loop % candidates.Length()
    {
        candidate := candidates[A_Index]
        if (WinActive(candidate))
        {
            WinMinimize, %candidate%
            return
        }
    }

    Loop % candidates.Length()
    {
        candidate := candidates[A_Index]
        if (WinExist(candidate))
        {
            WinActivate  ; last found window
            WinWaitActive, %candidate%,, 1
            return
        }
    }
}

ToggleWorldShip()
{
    global worldShipTitle
    if WinExist(worldShipTitle)
    {
        if WinActive(worldShipTitle)
        {
            WinMinimize, %worldShipTitle%
            return
        }
        EnsureWorldShipPlacement()
        WinActivate, %worldShipTitle%
        WinWaitActive, %worldShipTitle%,, 1
        return
    }
    Run, WorldShipTD.exe
}

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
    WinActivate, ahk_id %nextId%
    WinWaitActive, ahk_id %nextId%,, 1
}

OpenNewExplorer()
{
    Run, explorer.exe
}
