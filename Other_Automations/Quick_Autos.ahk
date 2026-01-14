#Requires AutoHotkey v1
#NoEnv ; Prevents Unnecessary Environment Variable lookup
#Warn ; Warn All (All Warnings Enabled)
#SingleInstance, Force ; Removes script already open warning when reloading scripts
SendMode Input
SetWorkingDir, %A_ScriptDir%
SetKeyDelay, 25

^Esc::Reload

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

    path := ExplorerCapturePath(originalHwnd)
    if (path = "")
        path := GetExplorerPath()
    if (path = "")
    {
        ToolTip, Could not get current path.
        SetTimer, HideQuickTip, -1200
        return
    }

    Send, ^w
    Sleep, 200

    if (WinExist("ahk_id " originalHwnd))
    {
        Send, ^t
        Sleep, 600

        WinActivate, ahk_id %originalHwnd%
        ExplorerWaitReady(originalHwnd, 2000)

        success := ExplorerNavigateAddressBar(originalHwnd, path, 4)

        if (!success)
        {
            ClipSaved := ClipboardAll
            Clipboard := path
            Sleep, 50

            Loop, 3
            {
                Send, {Esc}
                Sleep, 50
                Send, ^l
                Sleep, 250
                Send, ^a
                Sleep, 50
                Send, ^v
                Sleep, 100
                Send, {Enter}
                Sleep, 400

                observedPath := ExplorerGetAddressBarText(originalHwnd)
                if (observedPath != "" && observedPath = path)
                {
                    success := true
                    break
                }
                Sleep, 200
            }

            Clipboard := ClipSaved
        }

        if (success)
            ExplorerCloseAddressDropdown(originalHwnd)
        else
        {
            ToolTip, Navigation failed. Target: %path%
            SetTimer, HideQuickTip, -3000
        }
    }
    else
    {
        Run, explorer.exe "%path%"
        WinWait, ahk_class CabinetWClass,, 3
        if (!ErrorLevel)
            WinActivate
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

; Next hotkey (not yet implemented)
