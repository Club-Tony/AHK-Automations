# Other_Automations-specific README/Instructions
## Recommended: add shortcuts for the following files (if needed) into Startup folder [Run Dialog (Win+R - Shell:Startup)], as they are meant to always be running and don't launch through Launcher Script:

- **Quick_Autos.ahk** - lightweight always-on helper (Ctrl+Esc reload).
- **ToolTips.ahk** - shows a temporary tooltip of global hotkeys when requested.
- **Window_Switch.ahk** - quick focus/minimize/open for various windows/programs.
- **Alt+Q=Alt+F4.ahk** - remaps Alt+Q to Alt+F4 except on specified Firefox tabs.
- **Ctrl+Home=Sleep.ahk** - puts the machine to sleep with Ctrl+Home.
- **Launcher_Script.ahk** - hub to launch helper tools (RS focus, coordinate capture, coord.txt viewer).

---

## Essential Hotkeys

- `Ctrl+Esc` - reloads most scripts in this folder.
- `Esc` - exits certain launched scripts early.
- `Ctrl+Alt+T` - show/hide global hotkey tooltip (ToolTips.ahk).
- `Ctrl+Alt+E` - reset Explorer tab scaling (Quick_Autos.ahk).
- `Ctrl+Shift+Alt+R` - reset stuck modifier keys (Quick_Autos.ahk).
- `Ctrl+Shift+Alt+I` - run installers (new device setup, VS Code).
- `Alt+Q` - acts as Alt+F4.
- `Ctrl+Home` - sleep the machine.

---

## Window_Switch.ahk

- `Win+E` - Focus/minimize/cycle Explorer windows.
- `Win+Alt+E` - Open new Explorer window.
- `Win+F` - Focus/minimize/launch Firefox.
- `Win+Alt+V` - Launch/focus/minimize VS Code.
- `Ctrl+Win+Alt+V` - Focus/minimize/position VLC.
- `Win+Alt+R` - Launch RS focus helper (Alt+Z to use).

---

## Launcher_Script.ahk

- `Ctrl+Shift+Alt+C` — launch Coordinate Capture Helper (Alt+C starts capture, Alt+I shows more info).
- `Ctrl+Shift+Alt+O` — open the saved `coord.txt` capture file.

---

## Macros.ahk (Moved)

Macros.ahk has been moved to its own repository: [Macros-Script](https://github.com/club-tony/Macros-Script)
