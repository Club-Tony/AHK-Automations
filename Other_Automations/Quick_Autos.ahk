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

; Explorer tab reset (scaling fix) Ctrl+Alt+E
#If (WinActive("ahk_class CabinetWClass") || WinActive("ahk_class ExploreWClass"))
^!e::
    ExplorerResetAbort := false
    Hotkey, Esc, ExplorerResetCancel, On
    ClipSaved := ClipboardAll
    path := GetExplorerPath()
    if (ExplorerResetAbort)
        goto ExplorerResetCleanup
    if (path = "")
        goto ExplorerResetCleanup
    originalPath := RTrim(path, "\")
    ; Move focus away from address bar after getting path
    Send, {F6}
    Sleep, 80
    Send, {F6}
    Sleep, 80
    if (ExplorerResetAbort)
        goto ExplorerResetCleanup
    Send, ^t
    Sleep, 200
    if (ExplorerResetAbort)
        goto ExplorerResetCleanup
    Send, ^l
    Sleep, 80
    if (ExplorerResetAbort)
        goto ExplorerResetCleanup
    SendInput, {Raw}%path%
    Sleep, 50
    if (ExplorerResetAbort)
        goto ExplorerResetCleanup
    Send, {Enter}
    Sleep, 400
    if (ExplorerResetAbort)
        goto ExplorerResetCleanup
    ; CRITICAL: Move focus away from address bar before cycling tabs
    Send, {F6}
    Sleep, 80
    Send, {F6}
    Sleep, 80
    if (ExplorerResetAbort)
        goto ExplorerResetCleanup
    ; New tab opens at the end; search backward to find the original tab
    tabsSearched := 0
    maxTabs := 10
    foundOriginal := false
    Loop, %maxTabs%
    {
        if (ExplorerResetAbort)
            goto ExplorerResetCleanup
        SendInput, {Ctrl down}{Shift down}{Tab}{Shift up}{Ctrl up}
        Sleep, 250
        if (ExplorerResetAbort)
            goto ExplorerResetCleanup
        tabsSearched++
        currPath := GetExplorerPath(2, 0.8)
        if (ExplorerResetAbort)
            goto ExplorerResetCleanup
        ; Move focus away from address bar after reading path
        Send, {F6}
        Sleep, 80
        Send, {F6}
        Sleep, 80
        if (ExplorerResetAbort)
            goto ExplorerResetCleanup
        if (currPath = "")
            continue
        currPath := RTrim(currPath, "\")
        if (currPath = originalPath)
        {
            Send, ^w
            Sleep, 200
            if (ExplorerResetAbort)
                goto ExplorerResetCleanup
            foundOriginal := true
            if (tabsSearched > 0)
            {
                Loop, % tabsSearched
                {
                    if (ExplorerResetAbort)
                        goto ExplorerResetCleanup
                    SendInput, {Ctrl down}{Tab}{Ctrl up}
                    Sleep, 150
                    if (ExplorerResetAbort)
                        goto ExplorerResetCleanup
                }
            }
            break
        }
    }
    goto ExplorerResetCleanup
#If

ExplorerResetCancel:
    ExplorerResetAbort := true
return

ExplorerResetCleanup:
    Hotkey, Esc, ExplorerResetCancel, Off
    Clipboard := ClipSaved
    ExplorerResetAbort := false
return

GetExplorerPath(attempts := 2, waitSec := 0.7)
{
    Loop, %attempts%
    {
        Clipboard := ""
        Send, ^l
        Sleep, 80
        Send, ^c
        ClipWait, %waitSec%
        if (!ErrorLevel && Clipboard != "")
            return Clipboard
        Sleep, 80
    }
    return ""
}

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
