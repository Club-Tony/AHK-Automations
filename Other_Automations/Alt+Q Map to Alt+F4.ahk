#NoEnv
#Warn
SendMode Input
SetWorkingDir %A_ScriptDir%

; Block Alt+Q->Alt+F4 when specific Firefox tabs are active
#If !( WinActive("Intra: Shipping Request Form ahk_exe firefox.exe")
    || WinActive("Intra: Interoffice Request ahk_exe firefox.exe")
    || WinActive("Intra: Search ahk_exe firefox.exe") )
!q::Send !{F4}
#If
