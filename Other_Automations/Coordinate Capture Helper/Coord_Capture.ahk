#Requires AutoHotkey v1
#NoEnv
#SingleInstance, Force
#InstallMouseHook
#Persistent
SendMode Input
SetWorkingDir %A_ScriptDir%
SetBatchLines, -1
#MaxThreadsPerHotkey 1
SetTitleMatchMode, 2

CoordMode, Mouse, Screen

fields := ["1.","2.","3.","4.","5.","6.","7.","8.","9.","10."]
captureFile := A_ScriptDir "\coord.txt"
captureHeader := "Capture started " A_Now
entries := []
idx := 1
escConfirm := false
detailsVisible := false
captureEnabled := false
solidBackground := true
guiHwnd := 0

InitializeCapture()

; overlay with live coords
Gui, -Caption +AlwaysOnTop +ToolWindow +LastFound
Gui, Color, Black
Gui, Font, s20 q3, Arial
Gui, Add, Text, vTitleText cFF0000, Window Coordinates
Gui, Font, s32 q3, Arial
Gui, Add, Text, vCoordText cFF0000, X: XXXXX  Y: YYYYY
Gui, Font, s18 q3, Arial
Gui, Add, Text, vHintText cFF0000, Alt+C to capture | Alt+I for more info | Alt+G toggle bg
Gui, Font, s18 q3, Arial
Gui, Add, Text, vDetailsText cFF0000 w420 r22 +Wrap,
GuiControl, Hide, DetailsText
ApplyBackground()
SetTimer, Update, 50
Gui, Show, x0 y0 NA AutoSize
guiHwnd := WinExist()

return

^+!c::Reload
!c::StartCapture()
!i::ToggleDetails()
!g::ToggleBackground()
Esc::
    savedCapture := captureEnabled
    if (savedCapture)
    {
        WriteCapture()
        captureEnabled := false
    }
    if (!escConfirm)
    {
        escConfirm := true
        if (savedCapture)
            ToolTip, Done. Saved to coord.txt (Shortcut: Ctrl+Shift+Alt+O)
        SetTimer, ClearTip, -3000
        return
    }
    ExitApp

~LButton::
    if (!captureEnabled || idx > fields.Length())
        return
    CloseCaptureFile(captureFile)
    CoordMode, Mouse, Screen
    MouseGetPos, screenX, screenY, mouseWinId
    if (mouseWinId)
    {
        WinGetPos, winX, winY,,, ahk_id %mouseWinId%
        x := screenX - winX
        y := screenY - winY
    }
    else
    {
        x := screenX
        y := screenY
    }
    entries[idx] := fields[idx] ": {x: " x ", y: " y "}"
    WriteCapture()
    ToolTip, % "Saved " entries[idx] " (" idx " of " fields.Length() ")"
    SetTimer, ClearTip, -800
    idx++
    if (idx > fields.Length())
    {
        Sleep 300
        ToolTip, % "Done. Saved to " captureFile
        SetTimer, ClearTip, -3000
    }
return

ClearTip:
    ToolTip
return

Update() {
    global detailsVisible
    CoordMode, Mouse, Screen
    MouseGetPos, screenX, screenY, mouseWinId, controlClassNN
    mouseWinX := ""
    mouseWinY := ""
    mouseWinW := ""
    mouseWinH := ""
    winRelX := "N/A"
    winRelY := "N/A"
    clientX := "N/A"
    clientY := "N/A"

    if (mouseWinId)
    {
        WinGetPos, mouseWinX, mouseWinY, mouseWinW, mouseWinH, ahk_id %mouseWinId%
        winRelX := screenX - mouseWinX
        winRelY := screenY - mouseWinY
    }

    GuiControl,, CoordText, % "X: " winRelX "  Y: " winRelY

    if (!detailsVisible)
        return

    if (mouseWinId)
    {
        VarSetCapacity(pt, 8, 0)
        NumPut(screenX, pt, 0, "Int")
        NumPut(screenY, pt, 4, "Int")
        if DllCall("ScreenToClient", "ptr", mouseWinId, "ptr", &pt)
        {
            clientX := NumGet(pt, 0, "Int")
            clientY := NumGet(pt, 4, "Int")
        }
    }

    WinGet, activeId, ID, A
    WinGetTitle, activeTitle, ahk_id %activeId%
    WinGetClass, activeClass, ahk_id %activeId%
    WinGet, activeExe, ProcessName, ahk_id %activeId%
    WinGetPos, activeX, activeY, activeW, activeH, ahk_id %activeId%
    activeIdHex := (activeId != "") ? Format("0x{:X}", activeId) : "N/A"

    activeTitle := SanitizeText(activeTitle)
    activeClass := SanitizeText(activeClass)
    activeExe := SanitizeText(activeExe)

    controlText := "N/A"
    if (controlClassNN = "")
    {
        controlClassNN := "N/A"
    }
    else
    {
        ControlGetText, controlText, %controlClassNN%, ahk_id %mouseWinId%
        controlText := SanitizeText(controlText)
    }

    detailsText := "Active window:`n"
        . "Title: " activeTitle "`n"
        . "Class: " activeClass "`n"
        . "Exe: " activeExe "`n"
        . "ahk_id: " activeIdHex "`n`n"
        . "Mouse coords:`n"
        . "Screen: " screenX ", " screenY "`n"
        . "Window: " winRelX ", " winRelY "`n"
        . "Client: " clientX ", " clientY "`n`n"
        . "Control under mouse:`n"
        . "ClassNN: " controlClassNN "`n"
        . "Text: " controlText "`n`n"
        . "Active window position/size:`n"
        . "X: " activeX "  Y: " activeY "`n"
        . "W: " activeW "  H: " activeH

    GuiControl,, DetailsText, %detailsText%
}

ToggleDetails()
{
    global detailsVisible
    detailsVisible := !detailsVisible
    if (detailsVisible)
        GuiControl, Show, DetailsText
    else
        GuiControl, Hide, DetailsText
    Gui, Show, x0 y0 NA AutoSize
}

ToggleBackground()
{
    global solidBackground
    solidBackground := !solidBackground
    ApplyBackground()
}

ApplyBackground()
{
    global solidBackground, guiHwnd
    target := (guiHwnd ? "ahk_id " guiHwnd : "A")
    if (solidBackground)
    {
        WinSet, TransColor, Off, %target%
        WinSet, Transparent, 255, %target%
    }
    else
    {
        WinSet, TransColor, Black, %target%
    }
}

StartCapture()
{
    global captureEnabled, escConfirm, idx
    captureEnabled := true
    escConfirm := false
    idx := 1
    InitializeCapture()
    ToolTip, Capture started. Click to save coords.
    SetTimer, ClearTip, -1500
}

SanitizeText(text)
{
    StringReplace, text, text, `r, , All
    StringReplace, text, text, `n,  , All
    return text
}

InitializeCapture()
{
    global entries, captureFile, captureHeader, fields
    CloseCaptureFile(captureFile)
    if (entries.Length())
        entries.RemoveAt(1, entries.Length())
    Loop % fields.Length()
        entries[A_Index] := fields[A_Index] ": {x: , y: }"
    WriteCapture()
}

WriteCapture()
{
    global entries, captureFile, captureHeader
    FileDelete, %captureFile%
    FileAppend, % captureHeader "`n`n", %captureFile%
    Loop % entries.Length()
        FileAppend, % entries[A_Index] "`n", %captureFile%
}

CloseCaptureFile(captureFile)
{
    SplitPath, captureFile, fileName
    captureTitle := fileName " - Notepad"
    DetectHiddenWindows, On
    hwnd := WinExist(captureTitle)
    DetectHiddenWindows, Off
    if (hwnd && WinActive("ahk_id " hwnd))
    {
        SendInput, ^w
        Sleep 150
    }
}
