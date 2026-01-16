#Requires AutoHotkey v1
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
#SingleInstance, Force  ; Reload without prompt when Esc is pressed.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetTitleMatchMode, 2  ; Allow partial matches for Intra window title.
CoordMode, Mouse, Window  ; Window-relative coordinates from Window Spy.
SetKeyDelay, 50, 50
exportRunning := false
abortHotkey := false

; Window Coordinates (Intra Desktop Client - Assign Recip):
; Window Position: x: -7 y: 0 w: 1322 h: 1339
; 200, 245 = Scan field
; 130, 850 = Name field
; 1035, 185 = enable item var lookup button
; 1060, 185 = applying to all item vars button
; 1100, 365 = Package type
; 1100, 535 = Pieces field
; 1100, 650 = BSC Location (Destination)
; 500, 1025 = Select Preset
; 350, 900 = Docksided items preset select
; 725, 190 = Enable Parent Item #
; 500, 120 = Print Label button
; 40, 1300 = All Button
ScanField := {x: 200, y: 245}

^Esc::Reload

!+e::  ; paste list into scan field (normal speed)
    RunExport()
return

^+!e::  ; paste list into scan field (fast speed)
    RunExport(true)
return

RunExport(isFast := false)
{
    global exportRunning, ScanField
    if (exportRunning)
    {
        ShowTimedTooltip("Export already running!", 2000)
        return
    }
    exportRunning := true
    ResetAbort()

    ; Validate window first
    if (!FocusAssignRecipWindow())
    {
        ShowTimedTooltip("Could not find or activate Assign Recip window!", 3000)
        exportRunning := false
        return
    }
    if (!WinActive("Intra Desktop Client - Assign Recip"))
    {
        ShowTimedTooltip("Assign Recip window not active!", 3000)
        exportRunning := false
        return
    }

    ; Determine which file to use
    trackingFileTxt := A_ScriptDir . "\tracking_numbers.txt"
    trackingFileCsv := A_ScriptDir . "\tracking_numbers.csv"
    fileType := SelectTrackingFile()

    if (fileType = "")
    {
        ShowTimedTooltip("No tracking files found or both are empty!", 4000)
        exportRunning := false
        return
    }

    ; Read file contents based on type
    if (fileType = "txt")
    {
        FileRead, ItemList, %trackingFileTxt%
        if (ErrorLevel)
        {
            ShowTimedTooltip("Error reading TXT file!", 3000)
            exportRunning := false
            return
        }
        formatLabel := "TXT"
    }
    else  ; fileType = "csv"
    {
        ItemList := ReadCSVColumn(trackingFileCsv)
        if (ItemList = "")
        {
            ShowTimedTooltip("Error reading CSV or no valid data!", 3000)
            exportRunning := false
            return
        }
        formatLabel := "CSV"
    }

    ; Trim whitespace and validate content
    ItemList := Trim(ItemList)
    if (ItemList = "")
    {
        ShowTimedTooltip("Tracking file is empty!", 3000)
        exportRunning := false
        return
    }

    ; Count items before starting
    totalItems := 0
    Loop, Parse, ItemList, `n, `r
    {
        if (A_LoopField != "")
            totalItems++
    }

    if (totalItems = 0)
    {
        ShowTimedTooltip("No valid tracking numbers found in file!", 3000)
        exportRunning := false
        return
    }

    ; Set timing parameters based on speed mode
    if (isFast)
    {
        initialSleep := 200
        clickSleep := 75
        typeSleep := 40
        enterSleep := 150
        progressEvery := 100
        modeLabel := "FAST"
    }
    else
    {
        initialSleep := 500
        clickSleep := 200
        typeSleep := 100
        enterSleep := 400
        progressEvery := 50
        modeLabel := "NORMAL"
    }

    ; Show start message
    ShowTimedTooltip("Starting " modeLabel " export (" formatLabel "): " totalItems " items (Esc to abort)", 2000)
    Sleep, %initialSleep%

    itemCount := 0
    skippedCount := 0
    aborted := false
    startTime := A_TickCount

    ; Process each line
    Loop, Parse, ItemList, `n, `r
    {
        ; Check for abort request
        if (AbortRequested())
        {
            ShowTimedTooltip("Aborted at item " itemCount " of " totalItems, 4000)
            aborted := true
            break
        }

        ; Skip empty lines
        currentItem := Trim(A_LoopField)
        if (currentItem = "")
        {
            skippedCount++
            continue
        }

        itemCount++

        ; Click scan field
        MouseClick, left, % ScanField.x, % ScanField.y, 2
        Sleep, %clickSleep%

        ; Check abort again before typing
        if (AbortRequested())
        {
            ShowTimedTooltip("Aborted at item " itemCount " of " totalItems, 4000)
            aborted := true
            break
        }

        ; Type the tracking number
        SendInput, % "{Raw}" currentItem
        Sleep, %typeSleep%
        SendInput, {Enter}
        Sleep, %enterSleep%

        ; Show progress tooltip periodically
        if (progressEvery && Mod(itemCount, progressEvery) = 0)
        {
            percentDone := Round((itemCount / totalItems) * 100)
            ToolTip, % modeLabel " mode: " itemCount " of " totalItems " (" percentDone "%) - Esc to abort"
        }
    }

    ; Calculate elapsed time
    elapsedMs := A_TickCount - startTime
    elapsedSec := Round(elapsedMs / 1000, 1)

    ; Clear progress tooltip
    ToolTip

    ; Show completion message
    if (!aborted)
    {
        completionMsg := "Export complete!`n`n"
            . "Processed: " itemCount " items`n"
        if (skippedCount > 0)
            completionMsg .= "Skipped: " skippedCount " empty lines`n"
        completionMsg .= "Time: " elapsedSec "s`n"
            . "Mode: " modeLabel
        MsgBox, % completionMsg
    }

    exportRunning := false
}

#If ( WinActive("Intra Desktop Client - Assign Recip") && exportRunning )
Esc::
    abortHotkey := true
    ToolTip, Aborting...
return
#If

ResetAbort()
{
    global abortHotkey
    abortHotkey := false
}

AbortRequested()
{
    global abortHotkey
    return abortHotkey
}

FocusAssignRecipWindow()
{
    ; Bring forward Assign Recip if it's open before running the flow.
    assignTitle := "Intra Desktop Client - Assign Recip"
    if (!WinExist(assignTitle))
        return false
    WinActivate, %assignTitle%
    WinWaitActive, %assignTitle%,, 1
    return !ErrorLevel
}

ShowTimedTooltip(msg, duration := 3000)
{
    ToolTip, %msg%
    SetTimer, HideExportTooltip, -%duration%
}

HideExportTooltip:
    ToolTip
return

SelectTrackingFile()
{
    global trackingFileTxt, trackingFileCsv
    ; Returns: "txt", "csv", or "" (empty if no valid file)

    ; Simple check: does each file exist and have content?
    txtOK := false
    csvOK := false

    ; Check TXT file
    txtExists := FileExist(trackingFileTxt)
    if (txtExists != "")
    {
        FileRead, txtContent, %trackingFileTxt%
        if (!ErrorLevel)
        {
            txtContent := Trim(txtContent)
            if (txtContent != "")
                txtOK := true
        }
    }

    ; Check CSV file
    csvExists := FileExist(trackingFileCsv)
    if (csvExists != "")
    {
        FileRead, csvContent, %trackingFileCsv%
        if (!ErrorLevel)
        {
            csvContent := Trim(csvContent)
            if (csvContent != "")
                csvOK := true
        }
    }

    ; If only one exists, use it
    if (txtOK && !csvOK)
        return "txt"
    if (csvOK && !txtOK)
        return "csv"
    if (!txtOK && !csvOK)
        return ""

    ; Both exist - ask user via tooltip
    ToolTip, Both files found!`nPress 1 for TXT, 2 for CSV (5s timeout)
    startTime := A_TickCount
    while (A_TickCount - startTime < 5000)
    {
        if GetKeyState("1", "P")
        {
            ToolTip
            return "txt"
        }
        if GetKeyState("2", "P")
        {
            ToolTip
            return "csv"
        }
        Sleep 50
    }
    ToolTip
    return "txt"  ; Default to TXT on timeout
}

FileHasContent(filePath)
{
    FileRead, content, %filePath%
    if (ErrorLevel)
        return false
    content := Trim(content)
    if (content = "")
        return false
    return true
}

ReadCSVColumn(filePath)
{
    ; Returns: String with one tracking number per line
    FileRead, csvContent, %filePath%
    if (ErrorLevel)
        return ""

    result := ""
    firstRow := true
    Loop, Parse, csvContent, `n, `r
    {
        line := Trim(A_LoopField)
        if (line = "")
            continue

        ; Extract first column (before first comma)
        commaPos := InStr(line, ",")
        if (commaPos > 0)
            cell := Trim(SubStr(line, 1, commaPos - 1))
        else
            cell := line  ; No comma = single column

        ; Auto-detect header: Skip if first row looks like header text
        if (firstRow)
        {
            firstRow := false
            if (IsHeaderRow(cell))
                continue
        }

        if (cell != "")
            result .= cell "`n"
    }
    return RTrim(result, "`n")
}

IsHeaderRow(cell)
{
    ; Check if cell looks like a header (contains common header keywords)
    StringLower, cellLower, cell

    ; Check for common header keywords
    if InStr(cellLower, "tracking")
        return true
    if InStr(cellLower, "number")
        return true
    if InStr(cellLower, "id")
        return true
    if InStr(cellLower, "barcode")
        return true
    if InStr(cellLower, "reference")
        return true
    if InStr(cellLower, "item")
        return true

    ; Also check if entirely non-alphanumeric (unlikely for tracking number)
    if RegExMatch(cell, "^[^a-zA-Z0-9]+$")
        return true

    return false
}
