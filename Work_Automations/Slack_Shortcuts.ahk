#Requires AutoHotkey v1
#NoEnv
#Warn
#SingleInstance, Force
SendMode Input
SetKeyDelay, 25, 25
SetWorkingDir %A_ScriptDir%
SetTitleMatchMode, 2

; Built-in Slack “jump” is Ctrl+K (or Cmd+K on macOS). These hotkeys wrap it.

slackExe := "slack.exe"
slackWindow := "ahk_exe " slackExe
#IfWinActive, ahk_exe slack.exe
~Alt Up::return

^Esc::Reload

; Jump to specific channels (edit the names to your workspace).
!0::JumpToSlackChannel("daveyuan")
!1::JumpToSlackChannel("leona-array")
!2::JumpToSlackChannel("sps-byod")
!3::JumpToSlackChannel("sea124_ouroboros")
!4::JumpToSlackChannel("acp_bsc_comms")
!5::JumpToSlackChannel("spssea124lostnfound")
!s::JumpToSlackChannel("felsusad grovfred")
!a::JumpToSlackChannel("tstepama grovfred")
!j::JumpToSlackChannel("@jssjens")
!l::JumpToSlackChannel("@leobanks")
!r::JumpToSlackChannel("@grovfred")
#If

JumpToSlackChannel(channel)
{
    if (!FocusSlack())
        return
    OpenQuickSwitcher()
    SendInput, %channel%
    Sleep 175
    if (!IsSafeSearchFocus())
        return
    Send, {Enter}
}

OpenQuickSwitcher()
{
    ; Click near the search bar (top-center), then re-open and clear the switcher.
    CoordMode, Mouse, Screen
    WinGetPos, wx, wy, ww, wh, ahk_exe slack.exe
    if (ww && wh)
    {
        searchX := wx + ww//2
        searchY := wy + 40
        MouseClick, left, %searchX%, %searchY%
        Sleep 145
    }
    SendInput, ^k
    Sleep 225
    SendInput, ^a
    Sleep 105
    SendInput, {Backspace}
    Sleep 105
}

FocusSlack()
{
    global slackWindow
    if WinExist(slackWindow)
    {
        WinActivate
        WinWaitActive, %slackWindow%,, 1
        return !ErrorLevel
    }
    return false
}

IsSafeSearchFocus()
{
    ; Basic guard: ensure we're still in Slack before sending Enter.
    return WinActive("ahk_exe slack.exe")
}
