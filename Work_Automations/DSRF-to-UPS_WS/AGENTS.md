# AGENTS.md

AutoHotkey v1 automations that transfer data from "Intra: Shipping Request Form" into UPS WorldShip.
All operational scripts in this folder are working successfully in the current workflow.
`DSRF-to-UPS_WS.ahk` is the most stable template and the preferred baseline for future changes.

## Primary entry point (launcher integration)
- Use `Repositories/AHK-Automations/Work_Automations/Launcher Script (Keybinds).ahk` to launch these scripts.
- `Ctrl+Alt+C` launches the normal script in Firefox or the Legacy script in Chrome/Edge.
- `Ctrl+Alt+U` launches the Super-Speed script in Firefox or the Legacy script in Chrome/Edge.
- The launcher closes any existing DSRF-to-UPS scripts first, focuses Intra, sizes the window, and shows hotkey tooltips.

## Script inventory
- `DSRF-to-UPS_WS.ahk`: best, most stable template. Collects all form fields into a payload, then pastes into WorldShip.
- `DSRF-to-UPS_WS(Super-Speed).ahk`: speed-tuned variant with shorter sleeps and address-fill detection (more timing-sensitive).
- `DSRF-to-UPS_WS-Legacy.ahk`: Chrome/Edge-compatible flow with legacy coordinates and extra scrolling.
- `DSRF-to-UPS_WS-NewScale(Firefox+Legacy).ahk`: older step-by-step sequence; kept as a reference for alternate field offsets.
- `DSRF-to-UPS_WS-New_Approach.ahk`: concept stub for a future .bat-based flow; not active.

## How the scripts work (shared mechanics)
- Window targeting: activates Intra ("Intra: Shipping Request Form") and UPS WorldShip, then moves Intra to a fixed size.
- Coordinate-driven data entry: `CoordMode, Mouse, Window` with fixed pixel targets for each field.
- Field capture: uses clipboard copy for each field, then restores the clipboard to avoid user disruption.
- WorldShip flow: navigates between tabs (Ship From/Ship To/Service/Options), pastes into fields, and handles QVN recipients.
- Data hygiene: clears Postal Code before paste to avoid address-book autofill collisions.
- Scale lag: clicks a fixed point once to disable the electronic scale in WorldShip.

## DSRF-to-UPS_WS.ahk (stable template specifics)
- Uses Firefox zoom 60% for capture and returns to 100% after completion.
- Captures Cost Center for Business forms only; Personal forms use a separate field map.
- Uses a neutral click + `^{Home}` to normalize scroll before copying.
- Includes a declared-value read by tabbing from Postal Code instead of a raw coordinate hit.

## Editing guidelines
- Keep window geometry consistent with the 1530x1399 Intra sizing and the fixed UPS WorldShip coordinates.
- When changing hotkeys or adding variants, update:
  - `Repositories/AHK-Automations/Work_Automations/Launcher Script (Keybinds).ahk`
  - `Repositories/AHK-Automations/Other_Automations/ToolTips.ahk`
  - `Repositories/AHK-Automations/README.md`
- Prefer making changes in `DSRF-to-UPS_WS.ahk` first, then port to other variants if needed.
