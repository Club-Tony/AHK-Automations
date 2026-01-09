## Recommended: add shortcuts for the following files (if needed) into Startup folder [Run Dialog (Win+R - Shell:Startup)], as they are meant to always be running and don't launch through Launcher Script:

- **Launcher_Script_Other.ahk** — hub to launch helper tools (RuneScape focus, coordinate capture, coord.txt viewer).
- **Macros.ahk** — staged macro menu for fast click, turbo key hold, pure hold, and macro recorder playback.
- **Window_Switch.ahk** — quick focus/minimize/open for various windows/programs.
- **ToolTips.ahk** — shows a temporary tooltip of global hotkeys when requested.
- **Coordinate Capture Helper/Coord_Capture.ahk** — on-screen overlay to click-capture coordinates into `coord.txt`.
- **Alt+S = Quick Action Keybinds.ahk** — toggles `/` to send a single left click.
- **Alt+Q Map to Alt+F4.ahk** — remaps Alt+Q to Alt+F4 except on specified Firefox tabs.
- **Alt+A-WIN_Key_Backup.ahk** — backup Start/Run workflow (Alt+A then Alt+R) when the Win key is blocked.
- **RK71_Key_Fixes.ahk** — RK71 keyboard-friendly Start/Run helpers (Win+R alias plus Alt+A/Alt+R pairing).
- **Ctrl+Home key = Sleep.ahk** — puts the machine to sleep with Ctrl+Home.
- **Focus_RS_Window.ahk** — Alt+Z focuses the RuneScape client.

---

## Essential Hotkeys

- `Ctrl+Esc` — reloads most scripts in this folder.
- `Esc` — exits certain launched scripts early
- `Ctrl+Alt+T` — show/hide global hotkey tooltip (ToolTips.ahk).
- `Alt+Q` — acts as Alt+F4
- `Ctrl+Home` — sleep the machine.
- Start/Run helpers: `Alt+A` then `Alt+R` within 5s (Alt+A-WIN_Key_Backup); `Win+R` or `Alt+A` then `Alt+R` (RK71_Key_Fixes, also `Win+R` via `#r`).

---

## Launcher_Script_Other.ahk

- `Ctrl+Alt+R` — launch RuneScape focus helper (Focus_RS_Window.ahk).
- `Ctrl+Shift+Alt+C` — launch Coordinate Capture Helper (Alt+C starts capture, Alt+I shows more info).
- `Ctrl+Shift+Alt+O` — open the saved `coord.txt` capture file.

---

## Macros.ahk

- `Ctrl+Shift+Alt+Z` — open the macro menu overlay; use `Esc` to cancel/timeout.
- While menu is open: `F1` stage `/` => left-click toggle; `F2` stage autoclicker; `F3` stage turbo key hold; `F4` stage pure key hold; `F5` start recording.
- When staged/active: `/` toggles slash macro or playback; `F2`/`Esc` stop autoclicker; `F3`/`Esc` stop turbo hold; `F4`/`Esc` stop pure hold; `F5`/`Esc` stop recorder; `/` toggles playback if a recording exists.
- `Ctrl+Alt+P` — toggle SendMode (Input/Play) used by the macros.
