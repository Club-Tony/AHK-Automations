# Other_Automations-specific README/Instructions
## Recommended: add shortcuts for the following files (if needed) into Startup folder [Run Dialog (Win+R - Shell:Startup)], as they are meant to always be running and don't launch through Launcher Script:

- **Quick_Autos.ahk** - lightweight always-on helper (Ctrl+Esc reload).
- **Macros.ahk** - staged macro menu for fast click, turbo key hold, pure hold, and macro recorder playback.
- **ToolTips.ahk** - shows a temporary tooltip of global hotkeys when requested.
- **Window_Switch.ahk** - quick focus/minimize/open for various windows/programs.
- **Alt+Q Map to Alt+F4.ahk** - remaps Alt+Q to Alt+F4 except on specified Firefox tabs.
- **Ctrl+Home key = Sleep.ahk** - puts the machine to sleep with Ctrl+Home.
- **Launcher_Script_Other.ahk** - hub to launch helper tools (RS focus, coordinate capture, coord.txt viewer).

---

## Essential Hotkeys

- `Ctrl+Esc` - reloads most scripts in this folder.
- `Esc` - exits certain launched scripts early.
- `Ctrl+Alt+T` - show/hide global hotkey tooltip (ToolTips.ahk).
- `Alt+Q` - acts as Alt+F4.
- `Ctrl+Home` - sleep the machine.

---

## Launcher_Script_Other.ahk

- `Ctrl+Alt+R` — launch RS focus helper (Focus_RS_Window.ahk).
- `Ctrl+Shift+Alt+C` — launch Coordinate Capture Helper (Alt+C starts capture, Alt+I shows more info).
- `Ctrl+Shift+Alt+O` — open the saved `coord.txt` capture file.

---

## Macros.ahk

- `Ctrl+Shift+Alt+Z` — open the macro menu overlay; use `Esc` to cancel/timeout.
- While menu is open: `F1` stage `/` => left-click toggle; `F2` stage autoclicker; `F3` stage turbo key hold; `F4` stage pure key hold; `F5` start recording.
- To toggle off F1-F5 functions - `Esc` or corresponding FKey.
- Controller combos (vJoy + XInput): `L1+L2+R1+R2+B` starts turbo hold; `L1+L2+R1+R2+Y` starts pure hold; `L1+L2+R1+R2+X` is the kill switch.
- When F5 (Macro Recording) is staged/active: `F5` toggles start/end macro recording, `F12` toggles playback after recording. If using controller: `R1+R2+L1+L2+A` toggles recording macro or playback; `R1+R2+L1+L2+X` toggles off Macro recording function.
- Tooltip will warn if controller support is unavailable (vJoy/XInput missing).
- `Ctrl+Alt+P` — toggle SendMode (Input/Play) used by the macros. Useful as a switch to SendPlay sends if game doesn't allow for default (SendInput).
