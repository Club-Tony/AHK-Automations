#Requires AutoHotkey v1
#NoEnv
#SingleInstance, Force
SendMode Input
SetWorkingDir %A_ScriptDir%

; =============================================================================
; Cookie_Extractor.ahk
; Extracts Intra session cookies from Firefox's cookies.sqlite database
;
; Requires: sqlite3.exe in parent directory (DSRF-Export root)
; Hotkey:   Ctrl+Alt+K - Extract cookies from Firefox
; =============================================================================

; Parent directory (DSRF-Export root) for shared files
dsrfExportDir := A_ScriptDir . "\.."
cookiesFile := dsrfExportDir . "\cookies.txt"
intraDomain := "amazonmailservices.us.spsprod.net"

return ; end of auto-execute

; =============================================================================
; Ctrl+Alt+K: Extract cookies from Firefox SQLite database
; =============================================================================
^!k::
    ToolTip, Extracting cookies from Firefox...

    ; Find Firefox profile
    firefoxProfiles := A_AppData . "\Mozilla\Firefox\Profiles"
    if (!FileExist(firefoxProfiles))
    {
        ToolTip
        MsgBox, 48, Error, Firefox profiles folder not found.`n`nExpected: %firefoxProfiles%
        return
    }

    ; Find the default profile from profiles.ini
    profileDir := ""
    cookiesDb := ""
    profilesIni := A_AppData . "\Mozilla\Firefox\profiles.ini"

    if (FileExist(profilesIni))
    {
        ; Read profiles.ini to find the default profile
        FileRead, iniContent, %profilesIni%

        ; Parse sections to find Default=1
        currentPath := ""
        currentIsRelative := ""
        foundDefault := false

        Loop, Parse, iniContent, `n, `r
        {
            line := A_LoopField

            ; Check for new section
            if (RegExMatch(line, "^\[Profile\d+\]"))
            {
                ; If we had a default in previous section, use it
                if (foundDefault && currentPath != "")
                {
                    break
                }
                currentPath := ""
                currentIsRelative := ""
                foundDefault := false
            }
            else if (RegExMatch(line, "^Path=(.+)$", match))
            {
                currentPath := match1
            }
            else if (RegExMatch(line, "^IsRelative=(\d)$", match))
            {
                currentIsRelative := match1
            }
            else if (line = "Default=1")
            {
                foundDefault := true
            }
        }

        ; Build full path if we found a default
        if (foundDefault && currentPath != "")
        {
            if (currentIsRelative = "1")
                profileDir := A_AppData . "\Mozilla\Firefox\" . currentPath
            else
                profileDir := currentPath

            cookiesDb := profileDir . "\cookies.sqlite"
            if (!FileExist(cookiesDb))
            {
                profileDir := ""
                cookiesDb := ""
            }
        }
    }

    ; Fallback: scan for profiles with cookies.sqlite
    ; Priority: 1) .default-release  2) .default  3) any profile
    if (profileDir = "")
    {
        Loop, Files, %firefoxProfiles%\*, D
        {
            if (InStr(A_LoopFileName, ".default-release"))
            {
                testDb := A_LoopFileFullPath . "\cookies.sqlite"
                if (FileExist(testDb))
                {
                    profileDir := A_LoopFileFullPath
                    cookiesDb := testDb
                    break
                }
            }
        }
    }

    if (profileDir = "")
    {
        Loop, Files, %firefoxProfiles%\*, D
        {
            testDb := A_LoopFileFullPath . "\cookies.sqlite"
            if (FileExist(testDb))
            {
                profileDir := A_LoopFileFullPath
                cookiesDb := testDb
                break
            }
        }
    }

    if (profileDir = "" || cookiesDb = "")
    {
        ToolTip
        MsgBox, 48, Error, No Firefox profile with cookies.sqlite found.`n`nProfiles folder: %firefoxProfiles%
        return
    }

    ; Copy database to temp (Firefox locks it while running)
    tempDb := A_Temp . "\firefox_cookies_temp.sqlite"
    FileCopy, %cookiesDb%, %tempDb%, 1
    if (ErrorLevel)
    {
        ToolTip
        MsgBox, 48, Error, Could not copy cookies database.`n`nFirefox may have it locked.`nTry closing Firefox and retry.
        return
    }

    ; Find sqlite3.exe
    sqlite3 := dsrfExportDir . "\sqlite3.exe"
    if (!FileExist(sqlite3))
    {
        ; Check other common paths
        sqlite3Paths := ["C:\Program Files\SQLite\sqlite3.exe"
                       , "C:\Program Files (x86)\SQLite\sqlite3.exe"
                       , "C:\ProgramData\chocolatey\bin\sqlite3.exe"]

        for idx, testPath in sqlite3Paths
        {
            if (FileExist(testPath))
            {
                sqlite3 := testPath
                break
            }
        }
    }

    if (!FileExist(sqlite3))
    {
        FileDelete, %tempDb%
        ToolTip
        MsgBox, 48, Error, sqlite3.exe not found.`n`nRun DSRF-ExportSetup.bat or place sqlite3.exe in:`n%dsrfExportDir%
        return
    }

    ; Query cookies for our domain
    ; Write SQL to temp file to avoid command-line quoting issues
    sqlFile := A_Temp . "\cookie_query.sql"
    outputFile := A_Temp . "\cookie_output.txt"

    FileDelete, %sqlFile%
    FileDelete, %outputFile%

    sqlQuery := "SELECT name || '=' || value FROM moz_cookies WHERE host LIKE '%" . intraDomain . "%' ORDER BY name;"
    FileAppend, %sqlQuery%, %sqlFile%

    ; Run sqlite3 with SQL file as input
    RunWait, %ComSpec% /c ""%sqlite3%" "%tempDb%" < "%sqlFile%" > "%outputFile%"",, Hide
    queryResult := ErrorLevel

    FileDelete, %tempDb%
    FileDelete, %sqlFile%

    if (queryResult != 0 || !FileExist(outputFile))
    {
        ToolTip
        MsgBox, 48, Error, SQLite query failed (exit code: %queryResult%).
        return
    }

    FileRead, cookieLines, %outputFile%
    FileDelete, %outputFile%

    if (cookieLines = "")
    {
        ToolTip
        MsgBox, 48, Error, No cookies found for domain:`n%intraDomain%`n`nMake sure you're logged into Intra in Firefox.
        return
    }

    ; Join lines with "; "
    cookieValue := ""
    Loop, Parse, cookieLines, `n, `r
    {
        if (A_LoopField != "")
        {
            if (cookieValue != "")
                cookieValue .= "; "
            cookieValue .= A_LoopField
        }
    }

    cookieLen := StrLen(cookieValue)
    if (cookieLen < 100)
    {
        ToolTip
        MsgBox, 48, Error, Extracted cookies seem too short (%cookieLen% chars).`n`nMake sure you're logged into Intra.
        return
    }

    ; Check for expected markers
    hasValidMarkers := false
    if (InStr(cookieValue, "AWSALB") || InStr(cookieValue, "ASP.NET_SessionId") || InStr(cookieValue, "Saml2"))
        hasValidMarkers := true

    if (!hasValidMarkers)
    {
        cookiePreview := SubStr(cookieValue, 1, 100)
        MsgBox, 52, Validation Warning, Cookies may be incomplete (missing expected session markers).`n`nLength: %cookieLen% chars`nPreview: %cookiePreview%...`n`nSave anyway?
        IfMsgBox, No
            return
    }

    ; Backup existing file
    if FileExist(cookiesFile)
    {
        backupFile := cookiesFile . ".bak"
        FileCopy, %cookiesFile%, %backupFile%, 1
    }

    ; Save cookies - with retry logic for locked files
    saveSuccess := false
    maxRetries := 3

    Loop, %maxRetries%
    {
        ; Delete existing file first (ensures complete overwrite)
        if FileExist(cookiesFile)
        {
            FileDelete, %cookiesFile%
            if (ErrorLevel)
            {
                ; File might be locked - wait and retry
                Sleep 500
                continue
            }
        }

        ; Write new content
        FileAppend, %cookieValue%, %cookiesFile%
        if (!ErrorLevel)
        {
            saveSuccess := true
            break
        }

        ; Failed - wait and retry
        Sleep 500
    }

    if (!saveSuccess)
    {
        ToolTip
        MsgBox, 48, Error, Failed to write cookies to:`n%cookiesFile%`n`nThe file may be open in another application.`nClose cookies.txt and try again.
        return
    }

    ToolTip, Cookies saved!`nLength: %cookieLen% chars
    SetTimer, HideTooltip, -3000
return

HideTooltip:
    ToolTip
return
