#Requires AutoHotkey v1
#NoEnv ; Prevents Unnecessary Environment Variable lookup
#Warn ; Warn All (All Warnings Enabled)
#SingleInstance, Force ; Removes script already open warning when reloading scripts
SendMode Input
SetWorkingDir, %A_ScriptDir%

; Script for common macros (fast clicking, fast repeated key presses, etc)

menuActive := false
slashMacroOn := false
macroTipVisible := false
macroTipIdleBaseline := 0
macroTipEarlyHide := true
holdMacroReady := false
holdMacroOn := false
holdMacroKey := ""
holdMacroBoundKey := ""
holdMacroRepeatMs := 40
holdHoldReady := false
holdHoldOn := false
holdHoldKey := ""
holdHoldBoundKey := ""
autoClickReady := false
autoClickOn := false
autoClickInterval := 1000
recorderActive := false
recorderPlaying := false
recorderEvents := []
recorderStart := 0
recorderLast := 0
recorderMouseSampleMs := 40
sendMode := "Input"
sendModeTip := "SendMode: Input (toggle: Ctrl+Alt+P)"
prevSendMode := "Input"

return

^Esc::Reload

^!p::
    ToggleSendMode()
return

; Ctrl+Shift+Alt+Z shows a temporary menu for staged actions.
^+!z::
    if (menuActive)
        return
    DeactivateSlashMacro(true)
    DeactivateHoldMacro(true)
    DeactivatePureHold(true)
    DeactivateAutoclicker(true)
    StopRecorder(true)
    ClearMacroTips()
    menuActive := true
    ToolTip, % MenuTooltipText()
    SetTimer, MenuTimeout, -15000
return

#If (menuActive)
Esc::
    CloseMenu("timeout")
return

F1::
    CloseMenu()
    ActivateSlashMacro()
    ShowMacroToggledTip("Macro Toggled - Slash => Left click")
return

F2::
    CloseMenu()
    StartAutoclickerSetup()
return

F3::
    CloseMenu()
    StartHoldMacroSetup()
return

F4::
    CloseMenu()
    StartPureHoldSetup()
return

F5::
    CloseMenu()
    StartRecorder()
return
#If

#If (slashMacroOn)
$/::
    MouseClick, Left
return

Esc::
F1::
    DeactivateSlashMacro()
return
#If

#If (autoClickReady)
$/::
    ToggleAutoclicker()
return

F2::
Esc::
    DeactivateAutoclicker()
return
#If

#If (recorderActive && !recorderPlaying)
Esc::
F5::
    FinalizeRecording()
return
#If

#If (recorderEvents.MaxIndex() != "" && !recorderActive)
$/::ToggleRecorderPlayback()
#If

#If (!recorderActive && !recorderPlaying && recorderEvents.MaxIndex() != "")
Esc::ClearRecorder()
#If

ToggleRecorderPlayback()
{
    global recorderPlaying
    if (recorderPlaying)
        StopPlayback()
    else
        StartPlayback()
}

MenuTimeout:
    CloseMenu("timeout")
return

CloseMenu(reason := "")
{
    global menuActive
    SetTimer, MenuTimeout, Off
    ToolTip
    menuActive := false
    if (reason = "timeout")
    {
        ToolTip, Timed Out
        SetTimer, HideTimeoutTip, -3000
    }
}

HideTimeoutTip:
    ToolTip
return

ActivateSlashMacro()
{
    global slashMacroOn
    if (slashMacroOn)
        return
    slashMacroOn := true
    ShowMacroToggledTip("Macro Toggled - Slash => Left click")
}

DeactivateSlashMacro(silent := false)
{
    global slashMacroOn
    if (!slashMacroOn)
        return
    slashMacroOn := false
    if (!silent)
        ShowMacroToggledTip("Macro Toggled Off")
}

ShowMacroToggledTip(text := "Macro Toggled", durationMs := 3000, earlyHide := true)
{
    global macroTipVisible, macroTipIdleBaseline, macroTipEarlyHide, sendMode
    if InStr(text, "Macro Toggled Off")
        text := "Macro Toggled Off (SendMode: " sendMode ") - Esc to exit"
    else
        text := text " (SendMode: " sendMode ")"
    macroTipVisible := true
    macroTipEarlyHide := earlyHide
    macroTipIdleBaseline := earlyHide ? (A_TimeIdlePhysical + 1) : 0
    ToolTip, %text%
    SetTimer, HideMacroTipTimeout, % -durationMs
    if (earlyHide)
        SetTimer, HideMacroTipOnInput, 50
    else
        SetTimer, HideMacroTipOnInput, Off
}

HideMacroTipTimeout:
    HideMacroTip()
return

HideMacroTipOnInput:
    global macroTipVisible, macroTipIdleBaseline, macroTipEarlyHide
    if (!macroTipVisible || !macroTipEarlyHide)
        return
    if (A_TimeIdlePhysical < macroTipIdleBaseline)
        HideMacroTip()
return

HideMacroTip()
{
    global macroTipVisible
    if (!macroTipVisible)
        return
    macroTipVisible := false
    SetTimer, HideMacroTipTimeout, Off
    SetTimer, HideMacroTipOnInput, Off
    ToolTip
}

ClearMacroTips()
{
    global macroTipVisible
    macroTipVisible := false
    SetTimer, HideMacroTipTimeout, Off
    SetTimer, HideMacroTipOnInput, Off
    ToolTip
}

MenuTooltipText()
{
    global sendMode
    return "F1 - Stage left click with ""/"" key`n"
        . "F2 - Stage Autoclicker`n"
        . "F3 - Stage turbo keyhold`n"
        . "F4 - Stage pure key hold`n"
        . "F5 - Stage Record Macro`n"
        . "^!P - Toggle send mode (" sendMode ")"
}

StartAutoclickerSetup()
{
    global autoClickInterval, autoClickReady, autoClickOn
    tooltipText := "Type click frequency in ms and press Enter (15s timeout)."
    ToolTip, %tooltipText%
    SetTimer, HideTempTip, -15000
    InputBox, newInterval, Autoclicker, % "Enter click interval in ms (e.g., 100).", , , , , , , , 15
    SetTimer, HideTempTip, Off
    ToolTip
    if (ErrorLevel)
    {
        ShowMacroToggledTip("Autoclicker canceled")
        return
    }
    if newInterval is not integer
    {
        ShowMacroToggledTip("Autoclicker canceled (invalid number)")
        return
    }
    if (newInterval < 1)
        newInterval := 1
    autoClickInterval := newInterval
    autoClickOn := false
    autoClickReady := true
    ShowMacroToggledTip("Macro ready - / toggles autoclicker (" autoClickInterval " ms)")
}

ActivateAutoclicker()
{
    global autoClickOn, autoClickInterval
    if (autoClickOn)
        return
    autoClickOn := true
    SetTimer, AutoClickTick, %autoClickInterval%
    ShowMacroToggledTip("Macro Toggled - Autoclicker " autoClickInterval " ms")
}

ToggleAutoclicker()
{
    global autoClickOn
    if (autoClickOn)
    {
        StopAutoclickerKeepReady()
    }
    else
    {
        ActivateAutoclicker()
    }
}

StopAutoclickerKeepReady()
{
    global autoClickOn
    if (!autoClickOn)
        return
    autoClickOn := false
    SetTimer, AutoClickTick, Off
    ShowMacroToggledTip("Macro Toggled Off")
}

DeactivateAutoclicker(silent := false)
{
    global autoClickOn, autoClickReady
    if (!autoClickOn && !autoClickReady)
        return
    autoClickOn := false
    autoClickReady := false
    SetTimer, AutoClickTick, Off
    if (!silent)
        ShowMacroToggledTip("Macro Toggled Off")
}

AutoClickTick:
    MouseClick, Left
return

HideTempTip:
    ToolTip
return

#If (holdMacroReady)
Esc::
F3::
    DeactivateHoldMacro()
return
#If

#If (holdHoldReady)
Esc::
F4::
    DeactivatePureHold()
return
#If

#If (recorderPlaying)
$/::
    StopPlayback()
return

F5::
Esc::
    StopRecorder()
return
#If

StartHoldMacroSetup()
{
    global holdMacroReady, holdMacroOn, holdMacroKey, holdMacroBoundKey
    ; reset any existing hold macro
    DeactivateHoldMacro(true)
    tooltipText := "Input key to hold down (15s timeout)."
    ToolTip, %tooltipText%
    SetTimer, HideTempTip, -15000
    ; Use InputHook so non-text keys (F-keys, arrows, etc.) are captured via EndKey.
    ih := InputHook("L1 T15 V")
    ih.Start()
    ih.Wait()
    SetTimer, HideTempTip, Off
    ToolTip
    holdKey := ih.Input
    if (holdKey = "")
        holdKey := ih.EndKey
    holdKey := GetKeyName(holdKey)
    if (holdKey = "")
    {
        ShowMacroToggledTip("Keyhold canceled (invalid key)")
        return
    }
    holdMacroKey := holdKey
    BindHoldHotkey(holdMacroKey, "On")
    holdMacroReady := true
    holdMacroOn := false
    Gosub, HoldMacroToggle  ; first press activates hold
}

BindHoldHotkey(key, mode := "On")
{
    global holdMacroBoundKey
    if (holdMacroBoundKey != "")
        Hotkey, *%holdMacroBoundKey%, Off
    if (key = "")
    {
        holdMacroBoundKey := ""
        return
    }
    holdMacroBoundKey := key
    Hotkey, *%key%, HoldMacroToggle, %mode%
}

HoldMacroToggle:
    global holdMacroReady, holdMacroOn, holdMacroKey, holdMacroRepeatMs
    if (!holdMacroReady || holdMacroKey = "")
        return
    if (holdMacroOn)
    {
        Send, {%holdMacroKey% up}
        holdMacroOn := false
        SetTimer, HoldKeyRepeat, Off
        ShowMacroToggledTip("Keyup", 1000, false)
    }
    else
    {
        Send, {%holdMacroKey% down}
        holdMacroOn := true
        SetTimer, HoldKeyRepeat, % holdMacroRepeatMs
        ShowMacroToggledTip("Keydown", 1000, false)
    }
return

DeactivateHoldMacro(silent := false)
{
    global holdMacroReady, holdMacroOn, holdMacroKey, holdMacroBoundKey
    if (!holdMacroReady && !holdMacroOn)
        return
    if (holdMacroOn && holdMacroKey != "")
        Send, {%holdMacroKey% up}
    SetTimer, HoldKeyRepeat, Off
    holdMacroOn := false
    holdMacroReady := false
    holdMacroKey := ""
    BindHoldHotkey("", "Off")
    if (!silent)
        ShowMacroToggledTip("Macro Toggled Off")
}

HoldKeyRepeat:
    global holdMacroOn, holdMacroKey
    if (!holdMacroOn || holdMacroKey = "")
        return
    Send, {%holdMacroKey%}
return

StartRecorder()
{
    global recorderActive, recorderPlaying, recorderEvents, recorderStart, recorderLast, recorderMouseSampleMs
    StopRecorder(true)
    recorderActive := true
    recorderPlaying := false
    recorderEvents := []
    recorderStart := A_TickCount
    recorderLast := recorderStart
    SetTimer, RecorderSampleMouse, % recorderMouseSampleMs
    ShowMacroToggledTip("Recording macro... F5 to stop", 3000, true)
    ; install low-level hooks for keys via hotkey prefix below
}

StopRecorder(silent := false)
{
    global recorderActive, recorderPlaying
    if (!recorderActive && !recorderPlaying)
        return
    StopPlayback(true)
    recorderActive := false
    recorderPlaying := false
    SetTimer, RecorderSampleMouse, Off
    if (!silent)
        ShowMacroToggledTip("Macro Toggled Off")
}

StartPlayback()
{
    global recorderEvents, recorderPlaying
    if (recorderPlaying)
        return
    if (recorderEvents.MaxIndex() = "")
    {
        ShowMacroToggledTip("No recording to play")
        return
    }
    recorderPlaying := true
    ShowMacroToggledTip("Playing recorded macro (/ to stop)")
    SetTimer, RecorderPlayNext, -1
}

StopPlayback(silent := false)
{
    global recorderPlaying
    if (!recorderPlaying)
        return
    recorderPlaying := false
    SetTimer, RecorderPlayNext, Off
    if (!silent)
        ShowMacroToggledTip("Macro Toggled Off")
}

ClearRecorder()
{
    global recorderEvents, recorderStart, recorderLast
    recorderEvents := []
    recorderStart := 0
    recorderLast := 0
    ShowMacroToggledTip("Macro Toggled Off")
}

RecorderPlayNext:
    global recorderEvents, recorderPlaying, sendMode
    if (!recorderPlaying)
        return
    Loop % recorderEvents.MaxIndex()
    {
        if (!recorderPlaying)
            break
        evt := recorderEvents[A_Index]
        Sleep, % evt.delay
        if (evt.type = "key")
        {
            SendEventOrInput("{" evt.code " " evt.state "}", evt.state)
        }
        else if (evt.type = "mousebtn")
        {
            SendEventOrInput("{" evt.code " " evt.state "}")
        }
        else if (evt.type = "mousemove")
        {
            MouseMove, % evt.x, % evt.y, 0
        }
    }
    if (recorderPlaying)
        SetTimer, RecorderPlayNext, -1
return

SendEventOrInput(seq, state := "")
{
    ; Use a stable send path for recorder playback to avoid unexpected window drags.
    SendInput, %seq%
}

ToggleSendMode()
{
    global sendMode, menuActive
    if (sendMode = "Input")
        sendMode := "Play"
    else
        sendMode := "Input"
    ApplySendMode()
    ShowMacroToggledTip("SendMode: " sendMode)
    if (menuActive)
        ToolTip, % MenuTooltipText()
}

ApplySendMode()
{
    global sendMode
    SendMode %sendMode%
    SetKeyDelay, -1, -1
    SetMouseDelay, -1
}

RecorderSampleMouse:
    global recorderActive
    if (!recorderActive)
        return
    MouseGetPos, mx, my
    RecorderAddEvent("mousemove", "", "", mx, my)
return

; Global hook for keyboard/mouse while recording
#If (recorderActive)
~*LButton::
    RecorderAddEvent("mousebtn", "LButton", "Down")
return

~*LButton Up::
    RecorderAddEvent("mousebtn", "LButton", "Up")
return

~*RButton::
    RecorderAddEvent("mousebtn", "RButton", "Down")
return

~*RButton Up::
    RecorderAddEvent("mousebtn", "RButton", "Up")
return

~*MButton::
    RecorderAddEvent("mousebtn", "MButton", "Down")
return

~*MButton Up::
    RecorderAddEvent("mousebtn", "MButton", "Up")
return

~*WheelUp::
    RecorderAddEvent("mousebtn", "WheelUp", "")
return

~*WheelDown::
    RecorderAddEvent("mousebtn", "WheelDown", "")
return

~*WheelLeft::
    RecorderAddEvent("mousebtn", "WheelLeft", "")
return

~*WheelRight::
    RecorderAddEvent("mousebtn", "WheelRight", "")
return

~*a::RecorderAddEvent("key", "a", "Down")
~*a up::RecorderAddEvent("key", "a", "Up")
~*b::RecorderAddEvent("key", "b", "Down")
~*b up::RecorderAddEvent("key", "b", "Up")
~*c::RecorderAddEvent("key", "c", "Down")
~*c up::RecorderAddEvent("key", "c", "Up")
~*d::RecorderAddEvent("key", "d", "Down")
~*d up::RecorderAddEvent("key", "d", "Up")
~*e::RecorderAddEvent("key", "e", "Down")
~*e up::RecorderAddEvent("key", "e", "Up")
~*f::RecorderAddEvent("key", "f", "Down")
~*f up::RecorderAddEvent("key", "f", "Up")
~*g::RecorderAddEvent("key", "g", "Down")
~*g up::RecorderAddEvent("key", "g", "Up")
~*h::RecorderAddEvent("key", "h", "Down")
~*h up::RecorderAddEvent("key", "h", "Up")
~*i::RecorderAddEvent("key", "i", "Down")
~*i up::RecorderAddEvent("key", "i", "Up")
~*j::RecorderAddEvent("key", "j", "Down")
~*j up::RecorderAddEvent("key", "j", "Up")
~*k::RecorderAddEvent("key", "k", "Down")
~*k up::RecorderAddEvent("key", "k", "Up")
~*l::RecorderAddEvent("key", "l", "Down")
~*l up::RecorderAddEvent("key", "l", "Up")
~*m::RecorderAddEvent("key", "m", "Down")
~*m up::RecorderAddEvent("key", "m", "Up")
~*n::RecorderAddEvent("key", "n", "Down")
~*n up::RecorderAddEvent("key", "n", "Up")
~*o::RecorderAddEvent("key", "o", "Down")
~*o up::RecorderAddEvent("key", "o", "Up")
~*p::RecorderAddEvent("key", "p", "Down")
~*p up::RecorderAddEvent("key", "p", "Up")
~*q::RecorderAddEvent("key", "q", "Down")
~*q up::RecorderAddEvent("key", "q", "Up")
~*r::RecorderAddEvent("key", "r", "Down")
~*r up::RecorderAddEvent("key", "r", "Up")
~*s::RecorderAddEvent("key", "s", "Down")
~*s up::RecorderAddEvent("key", "s", "Up")
~*t::RecorderAddEvent("key", "t", "Down")
~*t up::RecorderAddEvent("key", "t", "Up")
~*u::RecorderAddEvent("key", "u", "Down")
~*u up::RecorderAddEvent("key", "u", "Up")
~*v::RecorderAddEvent("key", "v", "Down")
~*v up::RecorderAddEvent("key", "v", "Up")
~*w::RecorderAddEvent("key", "w", "Down")
~*w up::RecorderAddEvent("key", "w", "Up")
~*x::RecorderAddEvent("key", "x", "Down")
~*x up::RecorderAddEvent("key", "x", "Up")
~*y::RecorderAddEvent("key", "y", "Down")
~*y up::RecorderAddEvent("key", "y", "Up")
~*z::RecorderAddEvent("key", "z", "Down")
~*z up::RecorderAddEvent("key", "z", "Up")
~*1::RecorderAddEvent("key", "1", "Down")
~*1 up::RecorderAddEvent("key", "1", "Up")
~*2::RecorderAddEvent("key", "2", "Down")
~*2 up::RecorderAddEvent("key", "2", "Up")
~*3::RecorderAddEvent("key", "3", "Down")
~*3 up::RecorderAddEvent("key", "3", "Up")
~*4::RecorderAddEvent("key", "4", "Down")
~*4 up::RecorderAddEvent("key", "4", "Up")
~*5::RecorderAddEvent("key", "5", "Down")
~*5 up::RecorderAddEvent("key", "5", "Up")
~*6::RecorderAddEvent("key", "6", "Down")
~*6 up::RecorderAddEvent("key", "6", "Up")
~*7::RecorderAddEvent("key", "7", "Down")
~*7 up::RecorderAddEvent("key", "7", "Up")
~*8::RecorderAddEvent("key", "8", "Down")
~*8 up::RecorderAddEvent("key", "8", "Up")
~*9::RecorderAddEvent("key", "9", "Down")
~*9 up::RecorderAddEvent("key", "9", "Up")
~*0::RecorderAddEvent("key", "0", "Down")
~*0 up::RecorderAddEvent("key", "0", "Up")
~*Space::RecorderAddEvent("key", "Space", "Down")
~*Space up::RecorderAddEvent("key", "Space", "Up")
~*Enter::RecorderAddEvent("key", "Enter", "Down")
~*Enter up::RecorderAddEvent("key", "Enter", "Up")
~*Tab::RecorderAddEvent("key", "Tab", "Down")
~*Tab up::RecorderAddEvent("key", "Tab", "Up")
~*Esc::RecorderAddEvent("key", "Esc", "Down")
~*Esc up::RecorderAddEvent("key", "Esc", "Up")
~*Shift::RecorderAddEvent("key", "Shift", "Down")
~*Shift up::RecorderAddEvent("key", "Shift", "Up")
~*Ctrl::RecorderAddEvent("key", "Ctrl", "Down")
~*Ctrl up::RecorderAddEvent("key", "Ctrl", "Up")
~*Alt::RecorderAddEvent("key", "Alt", "Down")
~*Alt up::RecorderAddEvent("key", "Alt", "Up")
~*LShift::RecorderAddEvent("key", "LShift", "Down")
~*LShift up::RecorderAddEvent("key", "LShift", "Up")
~*RShift::RecorderAddEvent("key", "RShift", "Down")
~*RShift up::RecorderAddEvent("key", "RShift", "Up")
~*LControl::RecorderAddEvent("key", "LControl", "Down")
~*LControl up::RecorderAddEvent("key", "LControl", "Up")
~*RControl::RecorderAddEvent("key", "RControl", "Down")
~*RControl up::RecorderAddEvent("key", "RControl", "Up")
~*LAlt::RecorderAddEvent("key", "LAlt", "Down")
~*LAlt up::RecorderAddEvent("key", "LAlt", "Up")
~*RAlt::RecorderAddEvent("key", "RAlt", "Down")
~*RAlt up::RecorderAddEvent("key", "RAlt", "Up")
~*Up::RecorderAddEvent("key", "Up", "Down")
~*Up up::RecorderAddEvent("key", "Up", "Up")
~*Down::RecorderAddEvent("key", "Down", "Down")
~*Down up::RecorderAddEvent("key", "Down", "Up")
~*Left::RecorderAddEvent("key", "Left", "Down")
~*Left up::RecorderAddEvent("key", "Left", "Up")
~*Right::RecorderAddEvent("key", "Right", "Down")
~*Right up::RecorderAddEvent("key", "Right", "Up")
~*F1::RecorderAddEvent("key", "F1", "Down")
~*F1 up::RecorderAddEvent("key", "F1", "Up")
~*F2::RecorderAddEvent("key", "F2", "Down")
~*F2 up::RecorderAddEvent("key", "F2", "Up")
~*F3::RecorderAddEvent("key", "F3", "Down")
~*F3 up::RecorderAddEvent("key", "F3", "Up")
~*F4::RecorderAddEvent("key", "F4", "Down")
~*F4 up::RecorderAddEvent("key", "F4", "Up")
~*F5::RecorderAddEvent("key", "F5", "Down")
~*F5 up::RecorderAddEvent("key", "F5", "Up")
~*F6::RecorderAddEvent("key", "F6", "Down")
~*F6 up::RecorderAddEvent("key", "F6", "Up")
~*F7::RecorderAddEvent("key", "F7", "Down")
~*F7 up::RecorderAddEvent("key", "F7", "Up")
~*F8::RecorderAddEvent("key", "F8", "Down")
~*F8 up::RecorderAddEvent("key", "F8", "Up")
~*F9::RecorderAddEvent("key", "F9", "Down")
~*F9 up::RecorderAddEvent("key", "F9", "Up")
~*F10::RecorderAddEvent("key", "F10", "Down")
~*F10 up::RecorderAddEvent("key", "F10", "Up")
~*F11::RecorderAddEvent("key", "F11", "Down")
~*F11 up::RecorderAddEvent("key", "F11", "Up")
~*F12::RecorderAddEvent("key", "F12", "Down")
~*F12 up::RecorderAddEvent("key", "F12", "Up")
#If

FinalizeRecording()
{
    global recorderActive, recorderEvents
    if (!recorderActive)
        return
    recorderActive := false
    SetTimer, RecorderSampleMouse, Off
    ShowMacroToggledTip("Recording stopped - / toggles playback")
}

RecorderAddEvent(type, code := "", state := "", x := "", y := "")
{
    global recorderEvents, recorderLast
    now := A_TickCount
    delay := now - recorderLast
    recorderLast := now
    recEvt := {}
    recEvt.type := type
    recEvt.delay := delay
    if (type = "key" || type = "mousebtn")
    {
        recEvt.code := code
        recEvt.state := state
    }
    else if (type = "mousemove")
    {
        recEvt.x := x
        recEvt.y := y
    }
    recorderEvents.Push(recEvt)
}

StartPureHoldSetup()
{
    global holdHoldReady, holdHoldOn, holdHoldKey, holdHoldBoundKey
    DeactivatePureHold(true)
    tooltipText := "Input key to hold down (15s timeout)."
    ToolTip, %tooltipText%
    SetTimer, HideTempTip, -15000
    holdKey := ""
    Input, holdKey, L1 T15 V
    SetTimer, HideTempTip, Off
    ToolTip
    if (holdKey = "")
    {
        ShowMacroToggledTip("Keyhold canceled (no key)")
        return
    }
    holdHoldKey := holdKey
    BindPureHoldHotkey(holdHoldKey, "On")
    holdHoldReady := true
    holdHoldOn := false
    Gosub, PureHoldToggle
}

BindPureHoldHotkey(key, mode := "On")
{
    global holdHoldBoundKey
    if (holdHoldBoundKey != "")
        Hotkey, *%holdHoldBoundKey%, Off
    if (key = "")
    {
        holdHoldBoundKey := ""
        return
    }
    holdHoldBoundKey := key
    Hotkey, *%key%, PureHoldToggle, %mode%
}

PureHoldToggle:
    global holdHoldReady, holdHoldOn, holdHoldKey
    if (!holdHoldReady || holdHoldKey = "")
        return
    if (holdHoldOn)
    {
        Send, {%holdHoldKey% up}
        holdHoldOn := false
        ShowMacroToggledTip("Keyup", 1000, false)
    }
    else
    {
        Send, {%holdHoldKey% down}
        holdHoldOn := true
        ShowMacroToggledTip("Keydown", 1000, false)
    }
return

DeactivatePureHold(silent := false)
{
    global holdHoldReady, holdHoldOn, holdHoldKey, holdHoldBoundKey
    if (!holdHoldReady && !holdHoldOn)
        return
    if (holdHoldOn && holdHoldKey != "")
        Send, {%holdHoldKey% up}
    holdHoldOn := false
    holdHoldReady := false
    holdHoldKey := ""
    BindPureHoldHotkey("", "Off")
    if (!silent)
        ShowMacroToggledTip("Macro Toggled Off")
}
