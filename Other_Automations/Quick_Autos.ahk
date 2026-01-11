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

ToggleAudioOutput()
{
    current := GetDefaultPlaybackDevice()
    if (InStr(current, "Focusrite USB Audio", false))
        downCount := 1
    else if (InStr(current, "High Definition Audio Device", false))
        downCount := 2
    else
        downCount := 1
    ; Adjust counts if the Quick Settings output list order changes.

    OpenSoundOutputFlyout()
    Sleep 150
    SendInput, {Home}
    Sleep 50
    SendInput, {Down %downCount%}
    SendInput, {Enter}
    Sleep 500
    SendInput, {Esc}
}

GetDefaultPlaybackDevice()
{
    deviceName := ""
    devEnum := 0
    dev := 0
    store := 0
    try
        devEnum := ComObjCreate("{BCDE0395-E52F-467C-8E3D-C4579291692E}"
            , "{A95664D2-9614-4F35-A746-DE8DB63617E6}")
    catch
        return ""
    if (!devEnum)
        return ""

    DllCall(NumGet(NumGet(devEnum+0)+4*A_PtrSize), "ptr", devEnum
        , "int", 0, "int", 1, "ptr*", dev)
    if (!dev)
    {
        if (devEnum)
            ObjRelease(devEnum)
        return ""
    }

    DllCall(NumGet(NumGet(dev+0)+4*A_PtrSize), "ptr", dev, "int", 0
        , "ptr*", store)
    if (!store)
    {
        if (dev)
            ObjRelease(dev)
        if (devEnum)
            ObjRelease(devEnum)
        return ""
    }

    VarSetCapacity(pkey, 20, 0)
    DllCall("ole32\CLSIDFromString", "wstr"
        , "{A45C254E-DF1C-4EFD-8020-67D146A850E0}", "ptr", &pkey)
    NumPut(14, pkey, 16, "UInt")  ; PKEY_Device_FriendlyName
    VarSetCapacity(prop, 16 + A_PtrSize, 0)
    hr := DllCall(NumGet(NumGet(store+0)+5*A_PtrSize), "ptr", store
        , "ptr", &pkey, "ptr", &prop)
    if (hr >= 0)
        deviceName := StrGet(NumGet(prop, 8, "ptr"), "UTF-16")
    DllCall("ole32\PropVariantClear", "ptr", &prop)
    if (store)
        ObjRelease(store)
    if (dev)
        ObjRelease(dev)
    if (devEnum)
        ObjRelease(devEnum)
    return deviceName
}

OpenSoundOutputFlyout()
{
    if (IsSoundOutputFlyoutOpen())
    {
        SendInput, {Esc}
        Sleep 150
    }
    SendInput, #^v
    Sleep 250
}

IsSoundOutputFlyoutOpen()
{
    DetectHiddenWindows, On
    if (WinActive("ahk_class XamlExplorerHostIslandWindow ahk_exe ShellExperienceHost.exe"))
    {
        DetectHiddenWindows, Off
        return true
    }
    if (WinActive("ahk_class Windows.UI.Core.CoreWindow ahk_exe ShellExperienceHost.exe"))
    {
        DetectHiddenWindows, Off
        return true
    }
    DetectHiddenWindows, Off
    return false
}

; Next hotkey (not yet implemented)
