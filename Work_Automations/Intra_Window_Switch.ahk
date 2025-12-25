#Requires AutoHotkey v1
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Warn  ; Enable warnings to assist with detecting common errors.
#SingleInstance, Force  ; Reload without prompt when Esc is pressed.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetTitleMatchMode, 2  ; Allow partial title matches for Intra windows.

assignTitle := "Intra Desktop Client - Assign Recip"
updateTitle := "Intra Desktop Client - Update"
pickupTitle := "Intra Desktop Client - Pickup"
worldShipTitle := "UPS WorldShip"
qvnTitle := "Quantum View Notify Recipients"
exportedReportTitle := "ExportedReport.pdf"
browserExes := ["firefox.exe", "chrome.exe", "msedge.exe"]

^Esc::Reload

#a::ToggleFocusOrMinimize(assignTitle)
#u::ToggleFocusOrMinimize(updateTitle)
#p::ToggleFocusOrMinimize(pickupTitle)
#f::ToggleFirefox()
#s::ToggleFocusOrMinimizeExe("slack.exe")
#w::ToggleFocusOrMinimize(worldShipTitle)
#o::ToggleOutlookPwa()

#i::ToggleIntraGroup()
#!v::ToggleFocusOrMinimizeExe("Code.exe")

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
    if WinActive(title)
    {
        WinMinimize, %title%
    }
    else if WinExist(title)
    {
        WinActivate, %title%
        WinWaitActive, %title%,, 1
    }
}

MinimizeIfExists(title)
{
    if WinExist(title)
        WinMinimize, %title%
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
    Loop % titles.Length()
    {
        t := titles[A_Index]
        if (WinExist(t))
        {
            WinRestore, %t%
            primary := (primary = "") ? t : primary
        }
    }
    if (primary != "")
    {
        WinActivate, %primary%
        WinWaitActive, %primary%,, 1
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
    if WinActive("ahk_exe firefox.exe")
    {
        WinMinimize, ahk_exe firefox.exe
        return
    }
    if WinExist("ahk_exe firefox.exe")
    {
        WinActivate  ; last found window
        WinWaitActive, ahk_exe firefox.exe,, 1
        EnsureFirefoxWindow()
    }
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
