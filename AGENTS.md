# AGENTS.md

AutoHotkey v1 automation scripts for work-focused mailroom/logistics workflows plus general utilities.

## Structure
- Work_Automations/ for workflow-specific scripts.
- Other_Automations/ for general utilities (ToolTips, Quick_Autos, Window_Switch).
- Lib/ for shared libraries.

## Key patterns
- #Warn is enabled; declare locals in functions.
- Ctrl+Esc reloads scripts; Ctrl+Alt+T shows contextual tooltips.

## When adding or changing hotkeys
- Update Other_Automations/ToolTips.ahk.
- Update README.md.
- Add tooltip passthrough in the #If (TooltipActive) block when needed.
