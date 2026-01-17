# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

AutoHotkey v1 automation scripts primarily for work-focused mailroom/logistics software, UPS WorldShip, and desktop client workflows. Also includes general-purpose utility scripts in `Other_Automations/`.

## Architecture

**Work_Automations/** - Scripts for specific work workflows (Intra desktop client, UPS WorldShip, etc.)
**Other_Automations/** - General-purpose utility scripts:
- `Quick_Autos.ahk` - Misc hotkeys (audio toggle, Explorer tab reset)
- `ToolTips.ahk` - Displays hotkey help tooltips based on active window
- `Window_Switch.ahk` - Quick window focus/minimize shortcuts

**Lib/** - Shared libraries

## Key Patterns

- All scripts use AutoHotkey v1 syntax
- `Ctrl+Esc` reloads scripts
- `Ctrl+Alt+T` shows context-sensitive tooltip help
- `#Warn` directive is enabled - declare local variables explicitly in functions

## When Adding/Modifying Hotkeys

1. **Update ToolTips.ahk** - Add new hotkey to the tooltip display text
2. **Update README.md** - Document new hotkey in the appropriate section
3. **Add tooltip passthrough** - If the hotkey should dismiss the tooltip when pressed, add a passthrough line in the `#If (TooltipActive)` section:
   ```autohotkey
   ~^!+x::Gosub HideTooltips
   ```

## File Locations

- **ToolTips**: `Other_Automations/ToolTips.ahk`
- **Quick utilities**: `Other_Automations/Quick_Autos.ahk`
- **Launcher/keybinds**: `Work_Automations/Launcher Script (Keybinds).ahk`
