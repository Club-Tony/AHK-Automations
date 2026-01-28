#Requires AutoHotkey v2.0
#SingleInstance Force
InstallMouseHook()
#MaxThreadsPerHotkey 1

SendMode("Input")
SetWorkingDir(A_ScriptDir)
SetTitleMatchMode(2)
CoordMode("Mouse", "Screen")

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
myGui := Gui("-Caption +AlwaysOnTop +ToolWindow +LastFound")
myGui.BackColor := "Black"
myGui.SetFont("s20 q3", "Arial")
titleCtrl := myGui.Add("Text", "vTitleText cFF0000", "Window Coordinates")
myGui.SetFont("s32 q3", "Arial")
coordCtrl := myGui.Add("Text", "vCoordText cFF0000", "X: XXXXX  Y: YYYYY")
myGui.SetFont("s18 q3", "Arial")
hintCtrl := myGui.Add("Text", "vHintText cFF0000", "Alt+C to capture | Alt+I for more info | Alt+G toggle bg")
myGui.SetFont("s18 q3", "Arial")
detailsCtrl := myGui.Add("Text", "vDetailsText cFF0000 w420 r22 +Wrap", "")
detailsCtrl.Visible := false
ApplyBackground()
SetTimer(Update, 50)
myGui.Show("x0 y0 NA AutoSize")
guiHwnd := WinExist()

^+!c::Reload()
!c::StartCapture()
!i::ToggleDetails()
!g::ToggleBackground()

Esc:: {
    global captureEnabled, escConfirm
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
            ToolTip("Done. Saved to coord.txt (Shortcut: Ctrl+Shift+Alt+O)")
        SetTimer(ClearTip, -3000)
        return
    }
    ExitApp()
}

~LButton:: {
    global captureEnabled, idx, fields, entries, captureFile
    if (!captureEnabled || idx > fields.Length)
        return
    CloseCaptureFile(captureFile)
    CoordMode("Mouse", "Screen")
    MouseGetPos(&screenX, &screenY, &mouseWinId)
    if (mouseWinId)
    {
        WinGetPos(&winX, &winY, , , "ahk_id " mouseWinId)
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
    ToolTip("Saved " entries[idx] " (" idx " of " fields.Length ")")
    SetTimer(ClearTip, -800)
    idx++
    if (idx > fields.Length)
    {
        Sleep(300)
        ToolTip("Done. Saved to " captureFile)
        SetTimer(ClearTip, -3000)
    }
}

ClearTip() {
    ToolTip()
}

Update() {
    global detailsVisible, coordCtrl, detailsCtrl
    CoordMode("Mouse", "Screen")
    MouseGetPos(&screenX, &screenY, &mouseWinId, &controlClassNN)
    winRelX := "N/A"
    winRelY := "N/A"
    clientX := "N/A"
    clientY := "N/A"

    if (mouseWinId)
    {
        WinGetPos(&mouseWinX, &mouseWinY, &mouseWinW, &mouseWinH, "ahk_id " mouseWinId)
        winRelX := screenX - mouseWinX
        winRelY := screenY - mouseWinY
    }

    coordCtrl.Value := "X: " winRelX "  Y: " winRelY

    if (!detailsVisible)
        return

    if (mouseWinId)
    {
        pt := Buffer(8, 0)
        NumPut("Int", screenX, pt, 0)
        NumPut("Int", screenY, pt, 4)
        if DllCall("ScreenToClient", "ptr", mouseWinId, "ptr", pt)
        {
            clientX := NumGet(pt, 0, "Int")
            clientY := NumGet(pt, 4, "Int")
        }
    }

    pixelColor := PixelGetColor(screenX, screenY)
    pixelColor := StrReplace(pixelColor, "0x", "")

    try {
        activeId := WinGetID("A")
        activeTitle := WinGetTitle("ahk_id " activeId)
        activeClass := WinGetClass("ahk_id " activeId)
        activeExe := WinGetProcessName("ahk_id " activeId)
        WinGetPos(&activeX, &activeY, &activeW, &activeH, "ahk_id " activeId)
        activeIdHex := Format("0x{:X}", activeId)
    } catch {
        activeTitle := "N/A"
        activeClass := "N/A"
        activeExe := "N/A"
        activeX := "N/A"
        activeY := "N/A"
        activeW := "N/A"
        activeH := "N/A"
        activeIdHex := "N/A"
    }

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
        try
            controlText := SanitizeText(ControlGetText(controlClassNN, "ahk_id " mouseWinId))
        catch
            controlText := "N/A"
    }

    detailsText := "Active window:`n"
        . "Title: " activeTitle "`n"
        . "Class: " activeClass "`n"
        . "Exe: " activeExe "`n"
        . "ahk_id: " activeIdHex "`n`n"
        . "Mouse coords:`n"
        . "Screen: " screenX ", " screenY "`n"
        . "Window: " winRelX ", " winRelY "`n"
        . "Client: " clientX ", " clientY "`n"
        . "Color: " pixelColor (pixelColor = "FFFFFF" ? " (Use Window Spy [Ctrl+Shift+Alt+W] for browser color captures)" : "") "`n`n"
        . "Control under mouse:`n"
        . "ClassNN: " controlClassNN "`n"
        . "Text: " controlText "`n`n"
        . "Active window position/size:`n"
        . "X: " activeX "  Y: " activeY "`n"
        . "W: " activeW "  H: " activeH

    detailsCtrl.Value := detailsText
}

ToggleDetails() {
    global detailsVisible, detailsCtrl, myGui
    detailsVisible := !detailsVisible
    detailsCtrl.Visible := detailsVisible
    myGui.Show("x0 y0 NA AutoSize")
}

ToggleBackground() {
    global solidBackground
    solidBackground := !solidBackground
    ApplyBackground()
}

ApplyBackground() {
    global solidBackground, guiHwnd
    target := guiHwnd ? "ahk_id " guiHwnd : "A"
    if (solidBackground)
    {
        WinSetTransColor("Off", target)
        WinSetTransparent(255, target)
    }
    else
    {
        WinSetTransColor("Black", target)
    }
}

StartCapture() {
    global captureEnabled, escConfirm, idx
    captureEnabled := true
    escConfirm := false
    idx := 1
    InitializeCapture()
    ToolTip("Capture started. Click to save coords.")
    SetTimer(ClearTip, -1500)
}

SanitizeText(text) {
    text := StrReplace(text, "`r", "")
    text := StrReplace(text, "`n", " ")
    return text
}

InitializeCapture() {
    global entries, captureFile, captureHeader, fields
    CloseCaptureFile(captureFile)
    entries := []
    Loop fields.Length
        entries.Push(fields[A_Index] ": {x: , y: }")
}

WriteCapture() {
    global entries, captureFile, captureHeader
    try FileDelete(captureFile)
    FileAppend(captureHeader "`n`n", captureFile)
    Loop entries.Length
        FileAppend(entries[A_Index] "`n", captureFile)
}

CloseCaptureFile(path) {
    SplitPath(path, &fileName)
    captureTitle := fileName " - Notepad"
    DetectHiddenWindows(true)
    hwnd := WinExist(captureTitle)
    DetectHiddenWindows(false)
    if (hwnd && WinActive("ahk_id " hwnd))
    {
        SendInput("^w")
        Sleep(150)
    }
}
