#Requires AutoHotkey v1
#NoEnv ; Prevents Unnecessary Environment Variable lookup
#Warn ; Warn All (All Warnings Enabled)
#SingleInstance, Force  ; Reload without prompt when Esc is pressed.
SendMode Input ; Send works as SendInput
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetTitleMatchMode, 2
CoordMode, Mouse, Window

intraWinTitle := "Intra: Shipping Request Form"
intraWinExes := ["firefox.exe", "chrome.exe", "msedge.exe"]  ; priority order
TooltipActive := false
searchWinTitle := "Search - General"
searchShortcutsLaunched := false
showSearchTooltip := true
fastAssignPrevRunning := false
autoSmartsheetRunning := false
autoSmartsheetCancelled := false
tooltipMouseX := 0
tooltipMouseY := 0
SetTimer, DetectSearchWindow, 500  ; Watch for first appearance of Search - General
SetTimer, DetectFastAssignClosed, 1000

; Launch: Coordinate Helper ; Keybind: Ctrl+Shift+Alt+C
^+!c::
    Run, "C:\Users\daveyuan\Documents\GitHub\Repositories\AHK-Automations\Other_Automations\Coordinate Capture Helper\Coord_Capture.ahk"
    ShowTooltip("Coord Helper: Alt+C to capture, Alt+I for more info, Alt+G toggle bg", 4000)
return

; Launch: Window Spy ; Keybind: Ctrl+Shift+Alt+W
^+!w::
    ToggleWindowSpy()
return

^+!o::
    Run, "C:\Users\daveyuan\Documents\GitHub\Repositories\AHK-Automations\Other_Automations\Coordinate Capture Helper\coord.txt"
return

; Toggle: Tracking Numbers File ; Keybind: Ctrl+Alt+E
^!e::
    ToggleTrackingNumbersFile()
return

; Toggle: Tracking CSV File ; Keybind: Alt+Shift+C
!+c::
    ToggleTrackingCSVFile()
return

; Clear: Both Tracking Files ; Keybind: Ctrl+Shift+Alt+Delete
^!+Delete::
    ClearTrackingFiles()
return

; Launch: Smartsheets ; Keybind: Ctrl+Alt+L
^!l:: 
    if !(WinActive("Daily BSC Audit") || WinActive("Ouroboros BSC - SEA124"))
    {
        ShowTooltip("Smartsheets must be open in browser before continuing", 5000)
        Return
    }
    Run, C:\Users\daveyuan\Documents\GitHub\Repositories\AHK-Automations\Work_Automations\Daily_Audit.ahk ; Keybind: Ctrl+Alt+D
    Sleep 150
    Run, C:\Users\daveyuan\Documents\GitHub\Repositories\AHK-Automations\Work_Automations\Daily_Smartsheet.ahk ; Keybind: Ctrl+Alt+S
    Sleep 150
    TooltipText =
    (
Ctrl+Alt+D: Daily Audit
Ctrl+Alt+S: Daily Smartsheet
    )
    ShowTooltip(TooltipText, 5000)
Return

; Launch: Smartsheets (Auto) ; Keybind: Ctrl+Shift+Alt+L
^+!l::
    RunSmartsheetsAuto()
Return

; Launch: Super Saiyan Intra ; Keybind: Ctrl+Alt+I
^!i:: 
    FocusAssignRecipWindow()
    MouseMove, 200, 245  ; Move cursor to scan field in Assign Recip
    Sleep 75
    Run, C:\Users\daveyuan\Documents\GitHub\Repositories\AHK-Automations\Work_Automations\IT_Requested_IOs-Faster_Assigning.ahk ; Keybind: Alt+S
    Sleep 150
    Run, C:\Users\daveyuan\Documents\GitHub\Repositories\AHK-Automations\Work_Automations\Parent_Ticket_Creation-(BYOD).ahk ; Keybind: Alt+P
    Sleep 150
    Run, C:\Users\daveyuan\Documents\GitHub\Repositories\AHK-Automations\Work_Automations\Parent_Ticket_Creation-(GENERAL).ahk ; Keybind: Ctrl+Alt+P
    Sleep 150
    Run, C:\Users\daveyuan\Documents\GitHub\Repositories\AHK-Automations\Work_Automations\IT_Asset_Move.ahk ; Keybind: Alt+T
    Sleep 150
    TooltipText =
    (
SSJ-Intra Keybinds
Alt+S: Faster Assigning
Alt+P: BYOD Parent Ticket
Ctrl+Alt+P: Parent Ticket (General)
Alt+T: IT Asset Move
Ctrl+Alt+I: Relaunch Scripts + Tooltip
    )
    ShowTooltip(TooltipText, 10000)
Return

; Launch: Intra SSJ Search ; Keybind: Ctrl+Alt+F
^!f::
    showSearchTooltip := true
    Gosub, LaunchIntraSearchShortcuts
Return

; Launch: DSRF-to-UPS WorldShip ; Keybind: Ctrl+Alt+C
^!c::
    intraTitle := GetIntraWindowTitle()
    intraExe := GetIntraBrowserExe()
    if (intraTitle = "") {
        ShowTooltip("Open Intra: Shipping Request Form in Firefox/Chrome/Edge", 4000)
        Return
    }

    CloseWorldShipScripts()

    ; Ensure Intra Window is focused and Mouse cursor is positioned
    FocusIntraWindow()
    EnsureIntraWindow()
    Sleep 50
    MouseClick, left, 410, 581 ; Move to Cost Center field for tooltip visibility
    Sleep 50
    SendInput, {WheelUp 25} ; Scroll to top of form

    if (intraExe != "firefox.exe") {
        Run, "C:\Users\daveyuan\Documents\GitHub\Repositories\AHK-Automations\Work_Automations\DSRF-to-UPS_WS\DSRF-to-UPS_WS-Legacy.ahk"
        Sleep 150
        TooltipText =
        (
Ctrl+Alt+B: Business Form (Legacy)
Ctrl+Alt+P: Personal Form (Legacy)
Ctrl+Alt+C: Launches Legacy (Chrome/Edge)
Current Mode: Legacy (No zoom)
        )
    } else {
        Run, "C:\Users\daveyuan\Documents\GitHub\Repositories\AHK-Automations\Work_Automations\DSRF-to-UPS_WS\DSRF-to-UPS_WS.ahk"
        Sleep 150
        TooltipText =
        (
Ctrl+Alt+B: Business Form
Ctrl+Alt+P: Personal Form
Ctrl+Alt+C: Launches DSRF-to-UPS_WS Script
Ctrl+Alt+U: Launches Super-Speed version (Warning: May be unstable)
Current Mode: Normal Speed
        )
    }
    if (!TooltipActive)
        ShowTooltip(TooltipText, 5000)
Return

; Launch: DSRF-to-UPS WorldShip (Super-Speed) ; Keybind: Ctrl+Alt+U
^!u::
    intraTitle := GetIntraWindowTitle()
    intraExe := GetIntraBrowserExe()
    if (intraTitle = "") {
        ShowTooltip("Open Intra: Shipping Request Form in Firefox/Chrome/Edge", 4000)
        Return
    }

    CloseWorldShipScripts()

    FocusIntraWindow()
    EnsureIntraWindow()
    Sleep 50
    MouseClick, left, 410, 581
    Sleep 50
    SendInput, {WheelUp 25}

    if (intraExe != "firefox.exe") {
        Run, "C:\Users\daveyuan\Documents\GitHub\Repositories\AHK-Automations\Work_Automations\DSRF-to-UPS_WS\DSRF-to-UPS_WS-Legacy.ahk"
        Sleep 150
        TooltipText =
        (
Ctrl+Alt+B: Business Form (Legacy)
Ctrl+Alt+P: Personal Form (Legacy)
Ctrl+Alt+U: Launches Legacy (Chrome/Edge)
Current Mode: Legacy (No zoom)
        )
    } else {
        Run, "C:\Users\daveyuan\Documents\GitHub\Repositories\AHK-Automations\Work_Automations\DSRF-to-UPS_WS\DSRF-to-UPS_WS(Super-Speed).ahk"
        Sleep 75
        TooltipText =
        (
Ctrl+Alt+B: Business Form
Ctrl+Alt+P: Personal Form
Ctrl+Alt+C: Launches DSRF-to-UPS_WS Script
Ctrl+Alt+U: Launches Super-Speed version (Warning: May be unstable)
Current Mode: Super-Speed
        )
    }
    if (!TooltipActive)
        ShowTooltip(TooltipText, 5000)
Return

FocusIntraWindow()
{
    title := GetIntraWindowTitle()
    if (title = "")
        return false
    WinActivate, %title%
    WinWaitActive, %title%,, 1
    return !ErrorLevel
}

EnsureIntraWindow()
{
    title := GetIntraWindowTitle()
    if (title = "")
        return false
    ; Match the working dimensions used in Intra_Buttons. Adjust here if the target size changes.
    WinMove, %title%,, 1917, 0, 1530, 1399
    Sleep 150
    return true
}

FocusAssignRecipWindow()
{
    ; Bring forward Assign Recip if it's open before launching the search helper.
    assignTitle := "Intra Desktop Client - Assign Recip"
    if (!WinExist(assignTitle))
        return false
    WinActivate, %assignTitle%
    WinWaitActive, %assignTitle%,, 1
    return !ErrorLevel
}

ShowTooltip(TooltipText, durationMs)
{
    global TooltipActive, tooltipMouseX, tooltipMouseY
    SetTimer, HideLauncherTooltip, Off
    SetTimer, TooltipMouseCheck, Off
    MouseGetPos, tooltipMouseX, tooltipMouseY
    ToolTip, %TooltipText%
    TooltipActive := true
    SetTimer, HideLauncherTooltip, % -durationMs
    SetTimer, TooltipMouseCheck, 100
}

HideLauncherTooltip:
    SetTimer, HideLauncherTooltip, Off
    SetTimer, TooltipMouseCheck, Off
    TooltipActive := false
    ToolTip
Return

GetIntraWindowTitle()
{
    global intraWinTitle, intraWinExes
    ; Prefer the active Intra tab in priority order (Firefox > Chrome > Edge).
    Loop % intraWinExes.Length()
    {
        activeCandidate := intraWinTitle " ahk_exe " intraWinExes[A_Index]
        if (WinActive(activeCandidate))
            return activeCandidate
    }
    ; Otherwise pick the first present in priority order.
    Loop % intraWinExes.Length()
    {
        candidate := intraWinTitle " ahk_exe " intraWinExes[A_Index]
        if (WinExist(candidate))
            return candidate
    }
    return ""
}

GetIntraBrowserExe()
{
    global intraWinTitle, intraWinExes
    Loop % intraWinExes.Length()
    {
        activeCandidate := intraWinTitle " ahk_exe " intraWinExes[A_Index]
        if (WinActive(activeCandidate))
            return intraWinExes[A_Index]
    }
    Loop % intraWinExes.Length()
    {
        candidate := intraWinTitle " ahk_exe " intraWinExes[A_Index]
        if (WinExist(candidate))
            return intraWinExes[A_Index]
    }
    return ""
}

CloseWorldShipScripts()
{
    SetTitleMatchMode, 2
    DetectHiddenWindows, On
    WinGet, pidNormal, PID, DSRF-to-UPS_WS.ahk
    if pidNormal
        Process, Close, %pidNormal%
    WinGet, pidFast, PID, DSRF-to-UPS_WS(Super-Speed).ahk
    if pidFast
        Process, Close, %pidFast%
    WinGet, pidLegacy, PID, DSRF-to-UPS_WS-Legacy.ahk
    if pidLegacy
        Process, Close, %pidLegacy%
    DetectHiddenWindows, Off
    SetTitleMatchMode, 1
}

ToggleCoordTxt()
{
    capturePath := "C:\Users\daveyuan\Documents\GitHub\Repositories\AHK-Automations\Other_Automations\Coordinate Capture Helper\coord.txt"
    captureTitle := "coord.txt - Notepad"
    DetectHiddenWindows, On
    hwnd := WinExist(captureTitle)
    DetectHiddenWindows, Off
    if (hwnd)
    {
        if (WinActive("ahk_id " hwnd))
        {
            SendInput, ^w  ; Close the active coord.txt tab
            Sleep 150
        }
        else
        {
            WinActivate, ahk_id %hwnd%
            WinWaitActive, ahk_id %hwnd%,, 1
        }
        return
    }
    Run, notepad.exe "%capturePath%"
    WinWaitActive, %captureTitle%,, 2
}

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

RunSmartsheetsAuto()
{
    global autoSmartsheetRunning, autoSmartsheetCancelled
    if (autoSmartsheetRunning)
        return
    autoSmartsheetRunning := true
    autoSmartsheetCancelled := false
    if (!FocusPreferredBrowser())
    {
        ShowTooltip("Failed, ensure web browser active", 5000)
        autoSmartsheetRunning := false
        return
    }
    if (!SleepWithCancel(2500))
    {
        autoSmartsheetRunning := false
        return
    }
    if (!EnsureFirefoxActive())
    {
        ShowTooltip("Failed, ensure Firefox active", 5000)
        autoSmartsheetRunning := false
        return
    }

    if (!OpenDailyTabs())
    {
        autoSmartsheetRunning := false
        return
    }

    if (!RunDailyAuditThenSmartsheet())
    {
        autoSmartsheetRunning := false
        return
    }
    autoSmartsheetRunning := false
}

FocusPreferredBrowser()
{
    if WinExist("ahk_exe firefox.exe")
    {
        if (!WinActive("ahk_exe firefox.exe"))
        {
            WinRestore, ahk_exe firefox.exe
            WinActivate, ahk_exe firefox.exe
        }
        WinWaitActive, ahk_exe firefox.exe,, 1
        return WinActive("ahk_exe firefox.exe")
    }
    return false
}

EnsureFirefoxActive()
{
    if WinActive("ahk_exe firefox.exe")
        return true
    return false
}

EnsureFirefoxActiveOrWarn()
{
    if WinActive("ahk_exe firefox.exe")
        return true
    WinActivate, ahk_exe firefox.exe
    WinWaitActive, ahk_exe firefox.exe,, 1
    if WinActive("ahk_exe firefox.exe")
        return true
    ShowTooltip("Failed, ensure Firefox active", 5000)
    return false
}

OpenDailyTabs()
{
    if (!EnsureFirefoxActiveOrWarn())
        return false
    SendEvent, ^t
    if (!SleepWithCancel(250))
        return false
    if (!EnsureFirefoxActiveOrWarn())
        return false
    MouseClick, left, 607, 63, 2
    if (!SleepWithCancel(250))
        return false

    if (!EnsureFirefoxActiveOrWarn())
        return false
    SendInput, {Raw}*daily bsc audit
    if (!SleepWithCancel(250))
        return false
    SendEvent, {Down}
    if (!SleepWithCancel(250))
        return false
    SendEvent, {Enter}
    if (!SleepWithCancel(250))
        return false

    if (!EnsureFirefoxActiveOrWarn())
        return false
    SendEvent, ^t
    if (!SleepWithCancel(250))
        return false

    if (!EnsureFirefoxActiveOrWarn())
        return false
    MouseClick, left, 607, 63, 2
    if (!SleepWithCancel(250))
        return false

    if (!EnsureFirefoxActiveOrWarn())
        return false
    SendInput, {Raw}*smartsheet daily task
    if (!SleepWithCancel(250))
        return false
    SendEvent, {Down}
    if (!SleepWithCancel(250))
        return false
    SendEvent, {Enter}
    if (!SleepWithCancel(250))
        return false

    SendEvent, ^+{Tab}
    if (!WaitForAuditWindowActive(10000))
        return false
    return true
}

RunDailyAuditThenSmartsheet()
{
    global autoSmartsheetCancelled
    if (autoSmartsheetCancelled)
        return false
    PostSendHotkey("^!l")
    if (!SleepWithCancel(1500))
        return false
    if (!WaitForAuditWindowActive(10000))
        return false
    PostSendHotkey("^!d")
    if (!SleepWithCancel(8000))
        return false
    SendEvent, ^{Tab}
    if (!SleepWithCancel(600))
        return false
    if (!WaitForSmartsheetWindowActive(10000))
        return false
    PostSendHotkey("^!s")
    if (!SleepWithCancel(5000))
        return false
    return true
}

RunDailyForActiveTab()
{
    global autoSmartsheetCancelled
    if (autoSmartsheetCancelled)
        return false
    dailyType := GetDailyPageType()
    if (dailyType = "")
    {
        ShowTooltip("Smartsheets must be open in browser before continuing", 5000)
        return false
    }

    EnsureDailyScriptsRunningAll()
    if (!SleepWithCancel(2000))
        return false
    EnsureDailyScriptRunning(dailyType)
    if (!SleepWithCancel(1500))
        return false
    if (dailyType = "audit")
        PostSendHotkey("^!d")
    else
        PostSendHotkey("^!s")

    if (!SleepWithCancel(5000))
        return false
    return true
}

GetDailyPageType()
{
    WinGetTitle, title, A
    if (InStr(title, "Daily BSC Audit"))
        return "audit"
    if (InStr(title, "Ouroboros BSC - SEA124"))
        return "smartsheet"
    return ""
}

WaitForDailyWindowActive(timeoutMs)
{
    global autoSmartsheetCancelled
    startTick := A_TickCount
    Loop
    {
        if (autoSmartsheetCancelled)
            return false
        if (WinActive("Daily BSC Audit") || WinActive("Ouroboros BSC - SEA124"))
            return true
        if ((A_TickCount - startTick) >= timeoutMs)
            break
        Sleep 100
    }
    ShowTooltip("Smartsheets must be open in browser before continuing", 5000)
    return false
}

WaitForAuditWindowActive(timeoutMs)
{
    return WaitForWindowActiveTitle("Daily BSC Audit", timeoutMs)
}

WaitForSmartsheetWindowActive(timeoutMs)
{
    return WaitForWindowActiveTitle("Ouroboros BSC - SEA124", timeoutMs)
}

WaitForWindowActiveTitle(title, timeoutMs)
{
    global autoSmartsheetCancelled
    startTick := A_TickCount
    Loop
    {
        if (autoSmartsheetCancelled)
            return false
        if (WinActive(title))
            return true
        if ((A_TickCount - startTick) >= timeoutMs)
            break
        Sleep 100
    }
    ShowTooltip("Smartsheets must be open in browser before continuing", 5000)
    return false
}

EnsureDailyScriptRunning(dailyType)
{
    DetectHiddenWindows, On
    if (dailyType = "audit")
    {
        if (!WinExist("Daily_Audit.ahk ahk_class AutoHotkey"))
            Run, C:\Users\daveyuan\Documents\GitHub\Repositories\AHK-Automations\Work_Automations\Daily_Audit.ahk
    }
    else
    {
        if (!WinExist("Daily_Smartsheet.ahk ahk_class AutoHotkey"))
            Run, C:\Users\daveyuan\Documents\GitHub\Repositories\AHK-Automations\Work_Automations\Daily_Smartsheet.ahk
    }
    DetectHiddenWindows, Off
}

EnsureDailyScriptsRunningAll()
{
    DetectHiddenWindows, On
    if (!WinExist("Daily_Audit.ahk ahk_class AutoHotkey"))
        Run, C:\Users\daveyuan\Documents\GitHub\Repositories\AHK-Automations\Work_Automations\Daily_Audit.ahk
    if (!WinExist("Daily_Smartsheet.ahk ahk_class AutoHotkey"))
        Run, C:\Users\daveyuan\Documents\GitHub\Repositories\AHK-Automations\Work_Automations\Daily_Smartsheet.ahk
    DetectHiddenWindows, Off
}

PostSendHotkey(keys)
{
    ; Use SendEvent so another script can receive the hotkey reliably.
    SendEvent, %keys%
}

SleepWithCancel(totalMs)
{
    global autoSmartsheetCancelled
    elapsed := 0
    while (elapsed < totalMs)
    {
        if (autoSmartsheetCancelled)
            return false
        Sleep 50
        elapsed += 50
    }
    return true
}

LaunchIntraSearchShortcuts:
    global searchShortcutsLaunched, showSearchTooltip
    searchShortcutsLaunched := true
    Run, C:\Users\daveyuan\Documents\GitHub\Repositories\AHK-Automations\Work_Automations\Intra_Desktop_Search_Shortcuts.ahk
    Sleep 150
    TooltipText =
    (
Intra SSJ Search
Ctrl+Alt+F: Load and reload search script
Alt+D: Docksided items
Ctrl+Alt+D: Delivered items
Alt+O: On-shelf items
Alt+H: Outbound - Handed Off (down 3)
Alt+A: Arrived at BSC
Alt+P: Pickup from BSC
Alt+Space: Search Windows Quick Resize
    )
    if (showSearchTooltip)
        ShowTooltip(TooltipText, 7000)
Return

DetectSearchWindow:
    global searchShortcutsLaunched, searchWinTitle, showSearchTooltip
    if (searchShortcutsLaunched)
        Return
    if WinExist(searchWinTitle)
    {
        showSearchTooltip := false
        Gosub, LaunchIntraSearchShortcuts
    }
Return

DetectFastAssignClosed:
    global fastAssignPrevRunning
    DetectHiddenWindows, On
    running := WinExist("IT_Requested_IOs-Faster_Assigning.ahk ahk_class AutoHotkey")
    DetectHiddenWindows, Off
    if (fastAssignPrevRunning && !running)
        ShowTooltip("Fast-Assign Script Closed", 1500)
    fastAssignPrevRunning := running
Return

#If (autoSmartsheetRunning)
Esc::
    autoSmartsheetCancelled := true
return
#If

#If (TooltipActive)
~Esc::Gosub HideLauncherTooltip
~!s::Gosub HideLauncherTooltip
~!p::Gosub HideLauncherTooltip
~^!p::Gosub HideLauncherTooltip
~!t::Gosub HideLauncherTooltip
~^!i::Gosub HideLauncherTooltip
#If

TooltipMouseCheck:
    global tooltipMouseX, tooltipMouseY
    MouseGetPos, curX, curY
    if (curX != tooltipMouseX || curY != tooltipMouseY)
        Gosub, HideLauncherTooltip
return

ToggleTrackingNumbersFile()
{
    ; Close any open tracking files first
    CloseTrackingFiles()
    Sleep 200

    ; Prompt user to choose TXT or CSV
    ToolTip, Press 1 for TXT or 2 for CSV (5s timeout)
    startTime := A_TickCount
    choice := ""

    while (A_TickCount - startTime < 5000)
    {
        if GetKeyState("1", "P")
        {
            choice := "txt"
            break
        }
        if GetKeyState("2", "P")
        {
            choice := "csv"
            break
        }
        Sleep 50
    }
    ToolTip

    ; Default to TXT on timeout
    if (choice = "")
        choice := "txt"

    ; Call appropriate toggle function
    if (choice = "txt")
        ToggleTrackingTxtFile()
    else
        ToggleTrackingCSVFile()
}

CloseTrackingFiles()
{
    ; Close tracking_numbers.txt if open in Notepad (use Ctrl+W to close tab, not whole window)
    DetectHiddenWindows, On
    txtTitle := "tracking_numbers.txt - Notepad"
    hwndTxt := WinExist(txtTitle)
    if (hwndTxt)
    {
        WinActivate, ahk_id %hwndTxt%
        Sleep 50
        Send ^w
        Sleep 100
    }

    ; Close tracking_numbers.csv if open in Notepad (use Ctrl+W to close tab, not whole window)
    csvTitleNotepad := "tracking_numbers.csv - Notepad"
    hwndCsvNotepad := WinExist(csvTitleNotepad)
    if (hwndCsvNotepad)
    {
        WinActivate, ahk_id %hwndCsvNotepad%
        Sleep 50
        Send ^w
        Sleep 100
    }

    ; Close tracking_numbers.csv if open in Excel
    csvTitleExcel := "tracking_numbers.csv"
    hwndCsvExcel := WinExist(csvTitleExcel " ahk_exe EXCEL.EXE")
    if (hwndCsvExcel)
    {
        WinClose, ahk_id %hwndCsvExcel%
        Sleep 100
    }
    DetectHiddenWindows, Off
}

ToggleTrackingTxtFile()
{
    trackingFile := A_ScriptDir "\Bulk_Tracking_Export_to_Intra\tracking_numbers.txt"
    trackingTitle := "tracking_numbers.txt - Notepad"

    ; Check if file exists, create if not
    if (!FileExist(trackingFile))
    {
        FileAppend,, %trackingFile%
        if (ErrorLevel)
        {
            ShowTooltip("Error creating tracking_numbers.txt", 3000)
            return
        }
    }

    ; Open in Notepad (CloseTrackingFiles already closed it if it was open)
    Run, notepad.exe "%trackingFile%"
    WinWaitActive, %trackingTitle%,, 2
    if (!ErrorLevel)
        ShowTooltip("Opened tracking_numbers.txt", 1500)
}

ToggleTrackingCSVFile()
{
    csvFile := A_ScriptDir "\Bulk_Tracking_Export_to_Intra\tracking_numbers.csv"

    ; Create if doesn't exist
    if (!FileExist(csvFile))
    {
        FileAppend,, %csvFile%
        if (ErrorLevel)
        {
            ShowTooltip("Error creating tracking_numbers.csv", 3000)
            return
        }
    }

    ; Try to open in Excel first (CloseTrackingFiles already closed it if it was open)
    Run, "%csvFile%",, UseErrorLevel
    if (ErrorLevel)
    {
        ; Excel not available, fallback to Notepad
        csvTitleNotepad := "tracking_numbers.csv - Notepad"
        Run, notepad.exe "%csvFile%"
        WinWaitActive, %csvTitleNotepad%,, 2
        if (!ErrorLevel)
            ShowTooltip("Opened tracking_numbers.csv in Notepad", 1500)
        return
    }

    ; Wait longer for Excel to open and check if it succeeded
    Sleep 1500
    DetectHiddenWindows, On
    excelHwnd := WinExist("ahk_exe EXCEL.EXE")
    DetectHiddenWindows, Off

    if (excelHwnd)
    {
        ShowTooltip("Opened tracking_numbers.csv in Excel", 1500)
    }
    else
    {
        ; Excel didn't open, fallback to Notepad
        csvTitleNotepad := "tracking_numbers.csv - Notepad"
        Run, notepad.exe "%csvFile%"
        WinWaitActive, %csvTitleNotepad%,, 2
        if (!ErrorLevel)
            ShowTooltip("Opened tracking_numbers.csv in Notepad", 1500)
    }
}

ClearTrackingFiles()
{
    trackingDir := A_ScriptDir "\Bulk_Tracking_Export_to_Intra"
    txtFile := trackingDir "\tracking_numbers.txt"
    csvFile := trackingDir "\tracking_numbers.csv"

    txtCleared := false
    csvCleared := false

    ; Clear TXT file
    if FileExist(txtFile)
    {
        FileDelete, %txtFile%
        FileAppend,, %txtFile%
        if (!ErrorLevel)
            txtCleared := true
    }

    ; Clear CSV file
    if FileExist(csvFile)
    {
        FileDelete, %csvFile%
        FileAppend,, %csvFile%
        if (!ErrorLevel)
            csvCleared := true
    }

    ; Create CSV if doesn't exist
    if (!FileExist(csvFile))
    {
        FileAppend,, %csvFile%
        csvCleared := true
    }

    ; Show result
    msg := "Cleared tracking files:`n"
    if (txtCleared)
        msg .= "✓ tracking_numbers.txt`n"
    if (csvCleared)
        msg .= "✓ tracking_numbers.csv"
    ShowTooltip(msg, 3000)
}
