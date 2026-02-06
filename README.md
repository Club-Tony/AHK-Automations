# Note: Not recommended for use if unfamiliar with functionality - Test and QC hotkey actions before using for any work tasks
### -----Additionally, due to lack of automatic coordinate scaling functionality based on window size/monitor size, the following requirements are necessary to ensure all scripts function properly-----
## Requirements:
1. #### 3440 x 1440p (21:9 Aspect Ratio) Ultra-Wide monitor [As main monitor]
2. #### 2 separate desktop instances for logistics software, one for Assign Recip, and one for Update [Add a third for Pickup if using desk signature pad for customer signatures]
3. #### Assign Recip should be hooked to left side of screen while active (`WinKey+Left`), and Update hooked to the right side (`WinKey+Right`) [If using Pickup as well, it should be maximized]
4. #### Firefox*, Chrome, or Edge Browser *(Firefox preferred)
5. #### Please note: Multiple Scripts/hotkeys below are focused for only 1 particular mailroom's process and wouldn't apply to or be usable by most others
---
# Setup:
## Recommended: add shortcuts for the following files (if needed) into Startup folder [Run Dialog (Win+R - Shell:Startup)], as they are meant to always be running and don't launch through Launcher Script:
- **Launcher Script (Keybinds)** - loads global hotkeys and launcher bindings
- **Alt+Q Map to Alt+F4** - remaps `Alt+Q` to behave like `Alt+F4`
- **Alt+R_Post_Parent_Normalize** - restores Assign Recip window buttons/settings back to default state after Parent ticket automations
- **Bulk_Tracking_Export_to_Intra** - bulk tracking number export hotkeys (Alt+Shift+E, Ctrl+Shift+Alt+E)
- **Intra_Buttons** - adds quick-access hotkeys for common Interoffice Form actions
- **Intra_Desktop_Search_Shortcuts** - hotkeys for desktop client search and navigation
- **Intra_Focus_Fields** - jump-to-field shortcuts for commonly used desktop client fields
- **Intra_L&F** - Lost and Found Assign Recip helper flow
- **IO_Autos.ahk** - Interoffice Request Form Full Automations
- **Intra_Pickup_Shortcuts** - shortcuts for desktop client's Pickup tab navigation/actions
- **Intra_Posters_Full_Auto.ahk** - automates the full "poster program" workflow beginning-to-end
- **Intra_Update_Tab_Shortcuts** - shortcuts for desktop client's Update tab navigation/actions
- **IntraWinArrange** - arranges and snaps desktop client's windows to preset layouts
- **Resize_Intra_Search_Window** - auto-resizes the desktop client's search window results for better visibility
- **Right Click** - helper script for context/right-click actions (Reprinting labels for searched items/sending to manifest) [Note: "Use the Print screen key to open screen capture" must be toggled off in Windows settings]
- **Slack_Shortcuts** - custom Slack navigation and messaging hotkeys [Focused script - replace send strings with your own]
- **ToolTips** - displays Shortcuts/Hotkeys tooltips related to currently active window
- **UPS_WS_Shortcuts** - UPS workspace task shortcuts and macros
- **Window_Switch** - quick window switching/focusing/minimizing for various windows and programs
- **Yellow_Pouch** - shortcuts/macros for Yellow Pouch workflows
---
## Essential Hotkeys:
#### -`Ctrl+Alt+T` = ToolTips (Displays relevant Hotkey Menus depending on currently active window)
#### -`Esc` key = Exit certain launched scripts + Stop automations early
#### -`Ctrl+Esc` = Reloads all scripts + Stop automations early
#### -`Ctrl+Shift+Alt+I` = Run installers (VS Code) - sets up PowerShell `cd + Tab` + `git + Tab` shortcuts
## Run 'Launcher Script (Keybinds)' to launch certain scripts with Hotkeys:
#### -`Ctrl+Alt+C` - Launch DSRF-to-UPS_WS Script for shipping request form to UPS WorldShip field transfer automation
#### -`Ctrl+Alt+I` - Launch SSJ desktop client scripts bundle
#### -`Ctrl+Alt+F` - Launch/reload Search shortcuts script and show search hotkey tooltip
#### -`Ctrl+Alt+E` - Toggle tracking file (press 1 for TXT, 2 for CSV)
#### -`Ctrl+Alt+L` - Launch Daily Audit + Daily Smartsheet scripts
#### -`Ctrl+Shift+Alt+L` - Auto-run Daily Audit then Daily Smartsheet
#### -`Ctrl+Shift+Alt+D` - Run Daily Audit flow
#### -`Ctrl+Shift+Alt+S` - Run Daily Smartsheet flow
#### -`Ctrl+Shift+Alt+W` - Toggle Window Spy
#### -`Ctrl+Shift+Alt+O` - Toggle coord.txt open/close (Coord Capture helper)
## Bulk Tracking Export to Intra
#### -`Alt+Shift+E` - Paste list into Assign Recip scan field (normal speed)
#### -`Ctrl+Shift+Alt+E` - Paste list into Assign Recip scan field (fast speed)
#### -`Ctrl+Alt+E` - Toggle tracking file (press 1 for TXT, 2 for CSV)
#### -`Alt+Shift+C` - Toggle tracking_numbers.csv (direct access)
#### -`Ctrl+Shift+Alt+Delete` - Clear both tracking files

**Formats Supported:**
- **TXT**: One tracking number per line in tracking_numbers.txt
- **CSV**: Paste Excel/CSV column into tracking_numbers.csv (first column extracted)

If both files have content when running export, you'll be prompted to choose which to use (press 1 for TXT, 2 for CSV).
CSV headers are auto-detected and skipped automatically.
## Other Hotkeys:
#### -`Ctrl+Shift+Alt+C` - Launch Window Coordinate Capture Helper (Alt+C starts capture, Alt+I shows more info)
#### -`Alt+5` - Focus Intra Custom Carrier (Assign Recip)
#### -`Ctrl+Alt+O` - Intra Custom Carrier -> Other (Assign Recip)
