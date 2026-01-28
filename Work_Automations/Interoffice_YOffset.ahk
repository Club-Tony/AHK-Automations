; Shared Y-offset toggle + helpers for Intra: Interoffice Request scripts.

coordToggleIni := A_ScriptDir "\Interoffice_Coord_Toggle.ini"
coordToggleSection := "Interoffice"
coordToggleKey := "YOffsetEnabled"
coordYOffsetUp := -77
coordYOffsetDown := 2
coordYOffsetBuild := 13

GetInterofficeYOffset(mode := "up")
{
    global coordToggleIni, coordToggleSection, coordToggleKey
    global coordYOffsetUp, coordYOffsetDown, coordYOffsetBuild
    IniRead, enabled, %coordToggleIni%, %coordToggleSection%, %coordToggleKey%, 0
    if (enabled = 1)
    {
        if (mode = "down")
            IniRead, offset, %coordToggleIni%, %coordToggleSection%, YOffsetDown, %coordYOffsetDown%
        else if (mode = "build")
            IniRead, offset, %coordToggleIni%, %coordToggleSection%, YOffsetBuild, %coordYOffsetBuild%
        else
            IniRead, offset, %coordToggleIni%, %coordToggleSection%, YOffsetUp, %coordYOffsetUp%
        return offset
    }
    return 0
}

IsInterofficeYOffsetEnabled()
{
    global coordToggleIni, coordToggleSection, coordToggleKey
    IniRead, enabled, %coordToggleIni%, %coordToggleSection%, %coordToggleKey%, 0
    return (enabled = 1)
}

IOY(y, mode := "up")
{
    return y + GetInterofficeYOffset(mode)
}

