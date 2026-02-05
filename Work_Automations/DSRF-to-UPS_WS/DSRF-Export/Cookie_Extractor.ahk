#Requires AutoHotkey v1
#NoEnv
#Warn
#SingleInstance, Force
SendMode Input
SetWorkingDir %A_ScriptDir%
SetKeyDelay, 25

SetTitleMatchMode, 2
SetDefaultMouseSpeed, 0
CoordMode, Mouse, Window

; =============================================================================
; Cookie_Extractor.ahk
; Extracts session cookies from Firefox DevTools and saves to cookies.txt
; Hotkey: Ctrl+Alt+K (when Intra: Shipping Request Form window exists)
; =============================================================================

; Firefox DevTools coordinates (window-relative)
; These are based on default Firefox DevTools panel layout
; Adjust if your DevTools is docked differently or at different size

devToolsCoords := {}
devToolsCoords.NetworkFilter    := {x: 200, y: 150}   ; Network filter input area
devToolsCoords.HeadersTab       := {x: 1032, y: 860}  ; Headers tab in right panel
devToolsCoords.HeadersScrollArea := {x: 1250, y: 1250} ; Area to scroll for cookie header
devToolsCoords.RequestListArea  := {x: 400, y: 400}   ; Network request list area

cookiesFile := A_ScriptDir . "\cookies.txt"
extractionRunning := false  ; Flag to track if extraction is in progress

return ; end of auto-execute

; =============================================================================
; Esc: Cancel and reload script (only during extraction)
; =============================================================================
#If (extractionRunning)
Esc::
    extractionRunning := false
    ToolTip, Extraction cancelled.
    SetTimer, HideTooltip, -1500
    Reload
return
#If

; =============================================================================
; Ctrl+Alt+K: Extract cookies from Firefox DevTools
; =============================================================================
#If WinExist("Intra: Shipping Request Form ahk_exe firefox.exe")
^!k::
    global extractionRunning
    extractionRunning := true
    ToolTip, Starting cookie extraction... (Press Esc to cancel)

    ; Find and activate the Intra window
    dsrfTitle := "Intra: Shipping Request Form"

    if !WinExist(dsrfTitle " ahk_exe firefox.exe")
    {
        ToolTip
        MsgBox, 48, Error, Could not find Intra: Shipping Request Form in Firefox.
        extractionRunning := false
        return
    }

    WinActivate, %dsrfTitle% ahk_exe firefox.exe
    WinWaitActive, %dsrfTitle% ahk_exe firefox.exe,, 3
    if (ErrorLevel)
    {
        ToolTip
        MsgBox, 48, Error, Could not activate Firefox window.
        extractionRunning := false
        return
    }

    Sleep 300

    ; Step 1: Open DevTools with F12
    ToolTip, Opening DevTools...
    SendInput, {F12}
    Sleep 1500  ; Wait for DevTools to fully open

    ; Step 2: Switch to Network tab - click directly (Ctrl+Shift+E unreliable)
    ; TODO: Get Network tab coordinates from user if needed
    ToolTip, Switching to Network tab...
    ; Skip keyboard shortcut for now - DevTools should remember last tab
    Sleep 300

    ; Step 3: Reload page to capture fresh network traffic
    ; Click the browser reload button directly (keyboard shortcuts conflict with DevTools)
    ToolTip, Reloading page...
    MouseClick, left, 182, 871  ; Reload button coordinates
    Sleep 4000  ; Wait for page to fully reload and network requests to appear

    ; Step 4: Filter network requests to find executeQuery
    ToolTip, Searching for executeQuery request...

    ; Click in the network filter area and search
    SendInput, ^+f  ; Firefox DevTools: Open search in Network panel
    Sleep 300
    SendInput, executeQuery
    Sleep 300
    SendInput, {Enter}
    Sleep 500

    ; Step 5: Click on the first matching request (should be POST executeQuery)
    ; Use Down arrow to select from filter results, then Enter to focus
    SendInput, {Down}
    Sleep 200
    SendInput, {Enter}
    Sleep 500

    ; Step 6: Navigate to Headers tab
    ; The Headers tab should be visible in the right panel after selecting a request
    ToolTip, Opening Headers panel...

    ; Click Headers tab area (user-provided coordinates)
    global devToolsCoords
    MouseClick, left, % devToolsCoords.HeadersTab.x, % devToolsCoords.HeadersTab.y
    Sleep 400

    ; Step 7: Search for Cookie header within the Headers panel
    ToolTip, Finding Cookie header...

    ; Use Ctrl+F to search within the panel
    SendInput, ^f
    Sleep 300
    SendInput, Cookie:
    Sleep 300
    SendInput, {Enter}
    Sleep 300
    SendInput, {Escape}  ; Close search box
    Sleep 200

    ; Step 8: Scroll to Cookie header area and select the value
    ; Move to scroll area and scroll down to find cookie
    MouseMove, % devToolsCoords.HeadersScrollArea.x, % devToolsCoords.HeadersScrollArea.y
    Sleep 200

    ; Scroll down to reveal Cookie header (user mentioned 8 wheel downs)
    Loop, 8
    {
        SendInput, {WheelDown}
        Sleep 100
    }
    Sleep 300

    ; Triple-click to select the Cookie line
    ToolTip, Selecting cookie value...
    MouseClick, left, % devToolsCoords.HeadersScrollArea.x, % devToolsCoords.HeadersScrollArea.y, 3
    Sleep 300

    ; Copy the selection
    ClipSaved := ClipboardAll
    Clipboard := ""
    SendInput, ^c
    ClipWait, 2

    if (ErrorLevel)
    {
        Clipboard := ClipSaved
        ToolTip
        MsgBox, 48, Error, Could not copy cookie value.`nTry manual extraction:`n1. In Headers panel, scroll to "Cookie:" under Request Headers`n2. Triple-click to select the cookie value`n3. Copy (Ctrl+C) and paste into cookies.txt
        ; Close DevTools
        SendInput, {F12}
        extractionRunning := false
        return
    }

    cookieValue := Clipboard
    Clipboard := ClipSaved
    ClipSaved := ""

    ; Step 9: Clean up the cookie value
    ; Remove "Cookie: " prefix if present
    cookieValue := Trim(cookieValue, " `t`r`n")
    cookieValue := RegExReplace(cookieValue, "^Cookie:\s*", "")
    cookieValue := Trim(cookieValue, " `t`r`n")

    ; Step 10: Validate the cookie
    ToolTip, Validating cookie...

    cookieLen := StrLen(cookieValue)
    if (cookieLen < 100)
    {
        ToolTip
        MsgBox, 48, Validation Failed, Cookie value seems too short (%cookieLen% chars).`n`nExpected a long string with session tokens.`n`nTry manual extraction instead.
        SendInput, {F12}
        extractionRunning := false
        return
    }

    ; Check for expected cookie markers (session IDs, etc.)
    hasValidMarkers := false
    if (InStr(cookieValue, "JSESSIONID") || InStr(cookieValue, "AWSALB") || InStr(cookieValue, "awsalb") || InStr(cookieValue, "session"))
        hasValidMarkers := true

    if (!hasValidMarkers)
    {
        ToolTip
        cookiePreview := SubStr(cookieValue, 1, 100)
        MsgBox, 52, Validation Warning, Cookie value doesn't contain expected session markers.`n`nLength: %cookieLen% chars`nPreview: %cookiePreview%...`n`nSave anyway?
        IfMsgBox, No
        {
            SendInput, {F12}
            extractionRunning := false
            return
        }
    }

    ; Step 11: Save to cookies.txt
    ToolTip, Saving cookies...

    global cookiesFile

    ; Backup existing cookies file if it exists
    if FileExist(cookiesFile)
    {
        backupFile := cookiesFile . ".bak"
        FileCopy, %cookiesFile%, %backupFile%, 1
    }

    ; Write new cookies
    FileDelete, %cookiesFile%
    FileAppend, %cookieValue%, %cookiesFile%

    if (ErrorLevel)
    {
        ToolTip
        MsgBox, 48, Error, Failed to write cookies to:`n%cookiesFile%
        SendInput, {F12}
        extractionRunning := false
        return
    }

    ; Step 12: Close DevTools and show success
    SendInput, {F12}
    Sleep 300

    extractionRunning := false
    ToolTip, Cookies saved successfully!`nLength: %cookieLen% chars`nFile: cookies.txt
    SetTimer, HideTooltip, -4000
return
#If

; =============================================================================
; Alternative: Manual Cookie Save (Ctrl+Alt+Shift+K)
; Saves clipboard content directly to cookies.txt
; Use after manually copying the cookie value from DevTools
; =============================================================================
^!+k::
    global cookiesFile

    cookieValue := Trim(Clipboard, " `t`r`n")

    ; Remove "Cookie: " prefix if present
    cookieValue := RegExReplace(cookieValue, "^Cookie:\s*", "")
    cookieValue := Trim(cookieValue, " `t`r`n")

    if (cookieValue = "")
    {
        MsgBox, 48, Error, Clipboard is empty.`n`nCopy the Cookie header value from DevTools first.
        return
    }

    cookieLen := StrLen(cookieValue)
    if (cookieLen < 100)
    {
        MsgBox, 48, Warning, Clipboard content seems too short (%cookieLen% chars) to be a valid cookie.`n`nMake sure you copied the full Cookie: header value.
        return
    }

    ; Backup existing cookies file
    if FileExist(cookiesFile)
    {
        backupFile := cookiesFile . ".bak"
        FileCopy, %cookiesFile%, %backupFile%, 1
    }

    ; Write new cookies
    FileDelete, %cookiesFile%
    FileAppend, %cookieValue%, %cookiesFile%

    if (ErrorLevel)
    {
        MsgBox, 48, Error, Failed to write cookies to:`n%cookiesFile%
        return
    }

    ToolTip, Cookies saved from clipboard!`nLength: %cookieLen% chars
    SetTimer, HideTooltip, -3000
return

HideTooltip:
    ToolTip
return
