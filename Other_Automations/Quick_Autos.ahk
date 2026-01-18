#Requires AutoHotkey v1
#NoEnv ; Prevents Unnecessary Environment Variable lookup
#Warn ; Warn All (All Warnings Enabled)
#SingleInstance, Force ; Removes script already open warning when reloading scripts
SendMode Input
SetWorkingDir, %A_ScriptDir%
SetKeyDelay, 25

; Switch sound output Ctrl+Alt+S
^!s::
    ToggleAudioOutput()
return

; Explorer tab reset Ctrl+Alt+E
^!e::
    if (!ExplorerWindowExists())
    {
        ToolTip, No Explorer window running.
        SetTimer, HideQuickTip, -1200
        return
    }
    if (!IsExplorerActive())
    {
        ToolTip, Focus Explorer window and try again.
        SetTimer, HideQuickTip, -1500
        return
    }

    originalHwnd := WinExist("A")

    ; Capture path from address bar
    path := ExplorerCapturePathSilent(originalHwnd)

    if (path = "")
    {
        ToolTip, Could not get current path.
        SetTimer, HideQuickTip, -1200
        return
    }

    ; Resolve special folder names to actual paths
    if (!IsFilePath(path))
        path := ResolveSpecialFolder(path)

    if (path = "" || !IsFilePath(path))
    {
        ToolTip, Unsupported folder: %path%
        SetTimer, HideQuickTip, -1500
        return
    }

    ; Close current tab
    if (!WinActive("ahk_id " originalHwnd))
    {
        WinActivate, ahk_id %originalHwnd%
        WinWaitActive, ahk_id %originalHwnd%,, 0.5
    }
    SendInput, ^w
    Sleep, 150

    ; If close dialog appears, dismiss it
    if WinExist("ahk_class #32770 ahk_exe explorer.exe")
    {
        WinActivate, ahk_class #32770 ahk_exe explorer.exe
        Sleep, 50
        SendInput, {Enter}
        Sleep, 100
    }

    ; Get list of existing Explorer windows before opening new one
    existingWindows := {}
    WinGet, preList, List, ahk_class CabinetWClass
    Loop, %preList%
        existingWindows[preList%A_Index%] := true

    ; Open new Explorer window at the saved path
    Run, explorer.exe "%path%"

    ; Wait for new window to appear
    newHwnd := ""
    Loop, 20  ; Try for up to 2 seconds
    {
        Sleep, 100
        WinGet, postList, List, ahk_class CabinetWClass
        Loop, %postList%
        {
            thisHwnd := postList%A_Index%
            if (!existingWindows[thisHwnd])
            {
                newHwnd := thisHwnd
                break 2
            }
        }
    }

    if (newHwnd)
    {
        WinActivate, ahk_id %newHwnd%
        WinWaitActive, ahk_id %newHwnd%,, 1
        Sleep, 1000  ; Wait for window to fully render

        ; Find another Explorer window to merge into (the pre-existing one)
        targetMergeHwnd := FindOtherExplorerWindow(newHwnd)
        if (targetMergeHwnd)
        {
            MergeExplorerWindows(newHwnd, targetMergeHwnd)
        }

        ToolTip, Explorer tab reset.
        SetTimer, HideQuickTip, -800
    }
    else
    {
        ToolTip, Failed to open new window.
        SetTimer, HideQuickTip, -1500
    }
return

ExplorerWindowExists()
{
    WinGet, list, List, ahk_class CabinetWClass
    if (list > 0)
        return true
    WinGet, list, List, ahk_class ExploreWClass
    return (list > 0)
}

IsExplorerActive()
{
    return WinActive("ahk_class CabinetWClass") || WinActive("ahk_class ExploreWClass")
}

ExplorerGetWindowByHwnd(hwnd)
{
    shell := ComObjCreate("Shell.Application")
    for shellWindow in shell.Windows
    {
        try shellHwnd := shellWindow.HWND
        catch
            continue
        if (shellHwnd = hwnd)
            return shellWindow
    }
    return ""
}

ExplorerWaitReady(hwnd, timeoutMs := 2000)
{
    start := A_TickCount
    while (A_TickCount - start < timeoutMs)
    {
        win := ExplorerGetWindowByHwnd(hwnd)
        if (!win)
            return false
        try
        {
            if (!win.Busy && win.ReadyState = 4)
                return true
        }
        catch
            return false
        Sleep, 50
    }
    return false
}

ExplorerWaitForPath(hwnd, targetPath, timeoutMs := 2000)
{
    targetLeaf := PathLeaf(targetPath)
    start := A_TickCount
    while (A_TickCount - start < timeoutMs)
    {
        capturedPath := ExplorerCapturePath(hwnd)
        if (capturedPath != "" && (capturedPath = targetPath || capturedPath = targetLeaf))
            return true
        barText := ExplorerGetAddressBarText(hwnd)
        if (barText != "" && (barText = targetPath || barText = targetLeaf))
            return true
        Sleep, 100
    }
    return false
}

ExplorerNavigateAddressBar(targetHwnd, targetPath, attempts := 3)
{
    navClipSaved := ClipboardAll
    Clipboard := targetPath
    Loop, %attempts%
    {
        WinActivate, ahk_id %targetHwnd%
        Sleep, 100
        Send, {Esc}
        Sleep, 50
        Send, !d
        Sleep, 150
        Send, ^a
        Sleep, 50
        Send, ^v
        Sleep, 100
        Send, {Enter}

        if (ExplorerWaitForPath(targetHwnd, targetPath, 2000))
        {
            Clipboard := navClipSaved
            return true
        }
        Sleep, 200
    }
    Clipboard := navClipSaved
    return false
}

ExplorerFindAddressBarControl(targetHwnd)
{
    ControlGet, controlList, List,,, ahk_id %targetHwnd%
    Loop, Parse, controlList, `n
    {
        controlName := A_LoopField
        if (SubStr(controlName, 1, 4) = "Edit")
            return controlName
    }
    return ""
}

ExplorerCapturePath(targetHwnd)
{
    captureClipSaved := ClipboardAll
    Clipboard := ""
    WinActivate, ahk_id %targetHwnd%
    Sleep, 100
    Send, !d
    Sleep, 150
    Send, ^c
    ClipWait, 0.5
    capturedPath := Clipboard
    Clipboard := captureClipSaved
    return capturedPath
}

ExplorerCapturePathSilent(targetHwnd)
{
    ; Capture address bar using clipboard
    captureClipSaved := ClipboardAll
    Clipboard := ""

    ; Ensure window is active
    if (!WinActive("ahk_id " targetHwnd))
    {
        WinActivate, ahk_id %targetHwnd%
        WinWaitActive, ahk_id %targetHwnd%,, 0.5
        if (ErrorLevel)
            return ""
    }

    ; Try to capture path, with retry if focus isn't right
    Loop, 2
    {
        SendInput, ^l
        Sleep, 200
        SendInput, ^c
        ClipWait, 0.8
        if (!ErrorLevel)
            break

        ; First attempt failed - click center of window to ensure proper focus
        if (A_Index = 1)
        {
            SendInput, {Esc}
            WinGetPos, winX, winY, winW, winH, ahk_id %targetHwnd%
            clickX := winX + (winW // 2)
            clickY := winY + (winH // 2)
            CoordMode, Mouse, Screen
            Click, %clickX%, %clickY%
            Sleep, 200
            Clipboard := ""
        }
    }

    if (ErrorLevel)
    {
        Clipboard := captureClipSaved
        SendInput, {Esc}
        return ""
    }

    capturedPath := Clipboard
    Clipboard := captureClipSaved

    ; Close address bar dropdown
    SendInput, {Esc}
    Sleep, 50

    return capturedPath
}

ResolveSpecialFolder(folderName)
{
    ; Handle common Windows special folder display names
    if (folderName = "Desktop")
        return A_Desktop
    if (folderName = "Documents" || folderName = "My Documents")
        return A_MyDocuments
    if (folderName = "Downloads")
    {
        EnvGet, userProfile, USERPROFILE
        return userProfile . "\Downloads"
    }
    if (folderName = "Pictures")
    {
        EnvGet, userProfile, USERPROFILE
        return userProfile . "\Pictures"
    }
    if (folderName = "Music")
    {
        EnvGet, userProfile, USERPROFILE
        return userProfile . "\Music"
    }
    if (folderName = "Videos")
    {
        EnvGet, userProfile, USERPROFILE
        return userProfile . "\Videos"
    }
    return ""
}

FindOtherExplorerWindow(excludeHwnd)
{
    ; Find another Explorer window that's not the one we're resetting
    WinGet, windowList, List, ahk_class CabinetWClass
    Loop, %windowList%
    {
        hwnd := windowList%A_Index%
        if (hwnd != excludeHwnd)
            return hwnd
    }
    return ""
}

MergeExplorerWindows(sourceHwnd, targetHwnd)
{
    ; sourceHwnd = new window (the one to drag FROM - has the reset tab)
    ; targetHwnd = existing window (the one to drag INTO)
    CoordMode, Mouse, Screen

    ; Get source window (new window) position - drag from its tab area
    WinGetPos, srcX, srcY, srcW, srcH, ahk_id %sourceHwnd%
    srcTabX := srcX + 60   ; First tab position
    srcTabY := srcY + 12   ; Tab bar vertical center

    ; Get target window (existing window) position
    WinGetPos, tgtX, tgtY, tgtW, tgtH, ahk_id %targetHwnd%
    tgtTabX := tgtX + (tgtW // 2)  ; Middle of target's tab bar
    tgtTabY := tgtY + 12

    ; Activate source window (new window) and move mouse to its tab
    WinActivate, ahk_id %sourceHwnd%
    WinWaitActive, ahk_id %sourceHwnd%,, 0.5
    Sleep, 200

    MouseMove, srcTabX, srcTabY, 0
    Sleep, 150

    ; Start dragging (hold mouse down)
    Click, Down, Left
    Sleep, 150

    ; Drag to target window's tab bar
    MouseMove, tgtTabX, tgtTabY, 10
    Sleep, 300

    ; Release mouse
    Click, Up, Left
    Sleep, 300

    ; Activate the merged window (the target/existing window)
    WinActivate, ahk_id %targetHwnd%
}

ExplorerCloseAddressDropdown(targetHwnd)
{
    WinActivate, ahk_id %targetHwnd%
    Sleep, 50
    Send, {Esc}
}

PathLeaf(path)
{
    cleanedPath := RegExReplace(path, "[\\/]$")
    return RegExReplace(cleanedPath, "^.*[\\/]")
}

ExplorerGetAddressBarText(targetHwnd)
{
    ; First try: Get path from COM interface for the specific window handle
    ; This should work correctly with tabs in Windows 11
    try
    {
        shell := ComObjCreate("Shell.Application")
        for window in shell.Windows
        {
            try
            {
                if (window.HWND = targetHwnd)
                {
                    ; Try to get the document path
                    try
                    {
                        docPath := window.Document.Folder.Self.Path
                        if (docPath != "")
                            return docPath
                    }
                    catch
                    {
                        ; Try LocationURL as fallback
                        locationUrl := window.LocationURL
                        if (InStr(locationUrl, "file:///"))
                        {
                            decodedPath := StrReplace(locationUrl, "file:///", "")
                            decodedPath := StrReplace(decodedPath, "/", "\")
                            decodedPath := UriDecode(decodedPath)
                            if (decodedPath != "")
                                return decodedPath
                        }
                    }
                }
            }
        }
    }

    ; Second try: ControlGetText on address bar
    ControlGetFocus, focusCtrlName, ahk_id %targetHwnd%
    if (focusCtrlName != "")
    {
        ControlGetText, focusText, %focusCtrlName%, ahk_id %targetHwnd%
        if (focusText != "")
            return focusText
    }
    addrCtrl := ExplorerFindAddressBarControl(targetHwnd)
    if (addrCtrl != "")
    {
        ControlGetText, addrText, %addrCtrl%, ahk_id %targetHwnd%
        return addrText
    }
    return ""
}

IsFilePath(value)
{
    return RegExMatch(value, "i)^(?:[a-z]:\\|\\\\)")
}

GetExplorerPath()
{
    try
    {
        for explorerWindow in ComObjCreate("Shell.Application").Windows
        {
            if (explorerWindow.HWND = WinExist("A"))
            {
                try
                {
                    resolvedPath := explorerWindow.Document.Folder.Self.Path
                    if (resolvedPath != "")
                        return resolvedPath
                }
                catch
                {
                    try
                    {
                        locationUrl := explorerWindow.LocationURL
                        if (InStr(locationUrl, "file:///"))
                        {
                            decodedPath := StrReplace(locationUrl, "file:///", "")
                            decodedPath := StrReplace(decodedPath, "/", "\")
                            decodedPath := UriDecode(decodedPath)
                            return decodedPath
                        }
                    }
                }
            }
        }
    }
    catch
    {
        return ""
    }
    return ""
}

UriDecode(str)
{
    Loop
    {
        if (!RegExMatch(str, "i)(%[0-9A-F]{2})", hex))
            break
        str := StrReplace(str, hex, Chr("0x" . SubStr(hex, 2)))
    }
    return str
}

HideQuickTip:
    ToolTip
return

ToggleAudioOutput()
{
    focusrite := "Speakers (Focusrite USB Audio)"
    hda := "Speakers (High Definition Audio Device)"
    current := GetDefaultPlaybackDeviceName()
    if (InStr(current, "Focusrite USB Audio", false))
        target := hda
    else if (InStr(current, "High Definition Audio Device", false))
        target := focusrite
    else
        target := focusrite

    if (!SetDefaultPlaybackDeviceByName(target))
    {
        ToolTip, Audio device not found.
        SetTimer, HideAudioTip, -1500
    }
    else
    {
        if (target = focusrite)
            ToolTip, Focusrite Speaker Active
        else
            ToolTip, High Def Audio Device Active
        SetTimer, HideAudioTip, -1000
    }
}

HideAudioTip:
    ToolTip
return

GetDefaultPlaybackDeviceName()
{
    deviceName := ""
    devEnum := 0
    dev := 0
    try
        devEnum := ComObjCreate("{BCDE0395-E52F-467C-8E3D-C4579291692E}"
            , "{A95664D2-9614-4F35-A746-DE8DB63617E6}")
    catch
        return ""
    if (!devEnum)
        return ""

    hr := DllCall(NumGet(NumGet(devEnum+0)+4*A_PtrSize), "ptr", devEnum
        , "int", 0, "int", 1, "ptr*", dev)
    if (hr != 0 || !dev)
    {
        if (devEnum)
            ObjRelease(devEnum)
        return ""
    }

    deviceName := GetDeviceFriendlyName(dev)
    if (dev)
        ObjRelease(dev)
    if (devEnum)
        ObjRelease(devEnum)
    return deviceName
}

SetDefaultPlaybackDeviceByName(name)
{
    deviceId := GetPlaybackDeviceIdByName(name)
    if (deviceId = "")
        return false
    policy := 0
    try
        policy := ComObjCreate("{870af99c-171d-4f9e-af0d-e63df40c2bc9}"
            , "{f8679f50-850a-41cf-9c72-430f290290c8}")
    catch
        return false
    if (!policy)
        return false
    DllCall(NumGet(NumGet(policy+0)+13*A_PtrSize), "ptr", policy
        , "wstr", deviceId, "uint", 0)
    DllCall(NumGet(NumGet(policy+0)+13*A_PtrSize), "ptr", policy
        , "wstr", deviceId, "uint", 1)
    DllCall(NumGet(NumGet(policy+0)+13*A_PtrSize), "ptr", policy
        , "wstr", deviceId, "uint", 2)
    ObjRelease(policy)
    return true
}

GetPlaybackDeviceIdByName(name)
{
    devEnum := 0
    collection := 0
    count := 0
    deviceId := ""
    try
        devEnum := ComObjCreate("{BCDE0395-E52F-467C-8E3D-C4579291692E}"
            , "{A95664D2-9614-4F35-A746-DE8DB63617E6}")
    catch
        return ""
    if (!devEnum)
        return ""

    hr := DllCall(NumGet(NumGet(devEnum+0)+3*A_PtrSize), "ptr", devEnum
        , "int", 0, "int", 1, "ptr*", collection)
    if (hr != 0 || !collection)
    {
        if (devEnum)
            ObjRelease(devEnum)
        return ""
    }

    hr := DllCall(NumGet(NumGet(collection+0)+3*A_PtrSize), "ptr", collection
        , "uint*", count)
    if (hr != 0)
    {
        ObjRelease(collection)
        ObjRelease(devEnum)
        return ""
    }

    Loop % count
    {
        idx := A_Index - 1
        dev := 0
        hr := DllCall(NumGet(NumGet(collection+0)+4*A_PtrSize), "ptr", collection
            , "uint", idx, "ptr*", dev)
        if (hr != 0 || !dev)
            continue
        nameFound := GetDeviceFriendlyName(dev)
        if (InStr(nameFound, name, false))
        {
            deviceId := GetDeviceId(dev)
            ObjRelease(dev)
            break
        }
        ObjRelease(dev)
    }

    ObjRelease(collection)
    ObjRelease(devEnum)
    return deviceId
}

GetDeviceFriendlyName(dev)
{
    store := 0
    name := ""
    hr := DllCall(NumGet(NumGet(dev+0)+4*A_PtrSize), "ptr", dev, "int", 0
        , "ptr*", store)
    if (hr != 0 || !store)
        return ""
    name := ReadFriendlyNameFromStore(store)
    ObjRelease(store)
    return name
}

GetDeviceId(dev)
{
    idPtr := 0
    deviceId := ""
    hr := DllCall(NumGet(NumGet(dev+0)+5*A_PtrSize), "ptr", dev, "ptr*", idPtr)
    if (hr = 0 && idPtr)
    {
        deviceId := StrGet(idPtr, "UTF-16")
        DllCall("ole32\CoTaskMemFree", "ptr", idPtr)
    }
    return deviceId
}

ReadFriendlyNameFromStore(store)
{
    value := ""
    VarSetCapacity(pkey, 20, 0)
    DllCall("ole32\CLSIDFromString", "wstr"
        , "{A45C254E-DF1C-4EFD-8020-67D146A850E0}", "ptr", &pkey)
    NumPut(14, pkey, 16, "UInt")  ; PKEY_Device_FriendlyName
    VarSetCapacity(prop, 16 + A_PtrSize, 0)
    hr := DllCall(NumGet(NumGet(store+0)+5*A_PtrSize), "ptr", store
        , "ptr", &pkey, "ptr", &prop)
    if (hr >= 0)
        value := StrGet(NumGet(prop, 8, "ptr"), "UTF-16")
    DllCall("ole32\PropVariantClear", "ptr", &prop)
    return value
}

; Multi-tab Explorer reset Ctrl+Shift+Alt+E
^!+e::
    if (!ExplorerWindowExists())
    {
        ToolTip, No Explorer window running.
        SetTimer, HideQuickTip, -1200
        return
    }
    if (!IsExplorerActive())
    {
        ToolTip, Focus Explorer window and try again.
        SetTimer, HideQuickTip, -1500
        return
    }

    InputBox, tabCount, Multi-Tab Reset, How many tabs to reset? (1-9), , 220, 130, , , , , 1
    if (ErrorLevel || tabCount = "")
        return
    if tabCount is not integer
    {
        ToolTip, Invalid number.
        SetTimer, HideQuickTip, -1200
        return
    }
    if (tabCount < 1 || tabCount > 9)
    {
        ToolTip, Enter a number between 1 and 9.
        SetTimer, HideQuickTip, -1500
        return
    }

    ToolTip, Resetting %tabCount% tab(s)...
    Sleep, 500

    Loop, %tabCount%
    {
        currentTab := A_Index
        ToolTip, Resetting tab %currentTab% of %tabCount%...

        ; Run the same logic as Ctrl+Alt+E for the current tab
        if (!ResetCurrentExplorerTab())
        {
            ToolTip, Failed on tab %currentTab%. Stopping.
            SetTimer, HideQuickTip, -2000
            return
        }

        ; If not the last tab, switch to next tab
        if (A_Index < tabCount)
        {
            Sleep, 350
            SendInput, ^{Tab}
            Sleep, 400
        }
    }

    ToolTip, Reset %tabCount% tab(s) complete.
    SetTimer, HideQuickTip, -1500
return

ResetCurrentExplorerTab()
{
    ; Returns true on success, false on failure
    local originalHwnd, path, existingWindows, preList, newHwnd, postList, thisHwnd, targetMergeHwnd

    if (!IsExplorerActive())
        return false

    originalHwnd := WinExist("A")

    ; Capture path from address bar
    path := ExplorerCapturePathSilent(originalHwnd)

    if (path = "")
        return false

    ; Resolve special folder names to actual paths
    if (!IsFilePath(path))
        path := ResolveSpecialFolder(path)

    if (path = "" || !IsFilePath(path))
        return false

    ; Close current tab
    if (!WinActive("ahk_id " originalHwnd))
    {
        WinActivate, ahk_id %originalHwnd%
        WinWaitActive, ahk_id %originalHwnd%,, 0.5
    }
    SendInput, ^w
    Sleep, 150

    ; If close dialog appears, dismiss it
    if WinExist("ahk_class #32770 ahk_exe explorer.exe")
    {
        WinActivate, ahk_class #32770 ahk_exe explorer.exe
        Sleep, 50
        SendInput, {Enter}
        Sleep, 100
    }

    ; Get list of existing Explorer windows before opening new one
    existingWindows := {}
    WinGet, preList, List, ahk_class CabinetWClass
    Loop, %preList%
        existingWindows[preList%A_Index%] := true

    ; Open new Explorer window at the saved path
    Run, explorer.exe "%path%"

    ; Wait for new window to appear
    newHwnd := ""
    Loop, 20  ; Try for up to 2 seconds
    {
        Sleep, 100
        WinGet, postList, List, ahk_class CabinetWClass
        Loop, %postList%
        {
            thisHwnd := postList%A_Index%
            if (!existingWindows[thisHwnd])
            {
                newHwnd := thisHwnd
                break 2
            }
        }
    }

    if (!newHwnd)
        return false

    WinActivate, ahk_id %newHwnd%
    WinWaitActive, ahk_id %newHwnd%,, 1
    Sleep, 1000  ; Wait for window to fully render

    ; Find another Explorer window to merge into (the pre-existing one)
    targetMergeHwnd := FindOtherExplorerWindow(newHwnd)
    if (targetMergeHwnd)
        MergeExplorerWindows(newHwnd, targetMergeHwnd)

    return true
}

; Modifier key reset Ctrl+Shift+Alt+R
^+!r::
    SendInput, {Ctrl Up}{Alt Up}{Shift Up}{LWin Up}{RWin Up}
    ToolTip, Modifier keys reset
    SetTimer, HideQuickTip, -1000
return
