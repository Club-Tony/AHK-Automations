#Requires AutoHotkey v1
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
#SingleInstance, Force  ; Reload without prompt when Esc is pressed.
SendMode Input ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

SetTitleMatchMode, 2 ; allows for partial window title matches
; Two scoped Alt+Space handlers: Item Details vs Search Results windows.

^Esc::Reload

#IfWinActive, Item Details
!space::
    KeyWait, space
    KeyWait, Alt
    SendInput, !{space}
    Sleep 100
    SendInput, s
    Sleep 100
    SendInput, {Right}
    Sleep 100
    SendInput, {Down}
    MouseMove, 1820, 955
    Sleep 100
    SendInput, {Enter}
    Sleep 100
    MouseMove, 200, 500
Return
#IfWinActive

#IfWinActive, Search Results:
!space::
    KeyWait, space
    KeyWait, Alt
    SendInput, !{space}
    Sleep 100
    SendInput, s
    Sleep 100
    SendInput, {Right}
    Sleep 100
    SendInput, {Down}
    MouseMove, 2050, 1025
    Sleep 100
    SendInput, {Enter}
    Sleep 100
    MouseMove, 115, 90
    Sleep 100
    Send, {Down}
Return
#IfWinActive
