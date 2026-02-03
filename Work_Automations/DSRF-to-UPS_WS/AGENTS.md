# AGENTS.md

AutoHotkey v1 automations that transfer data from "Intra: Shipping Request Form" into UPS WorldShip.

## Scope
- This folder contains coordinate-based DSRF -> WorldShip scripts plus an API-based package in `DSRF-Export/`.
- `DSRF-to-UPS_WS.ahk` is the primary baseline and most stable script for behavior changes.
- Legacy variants are stored in `Legacy/`.

## Primary entry point (launcher integration)
- Use `Repositories/AHK-Automations/Work_Automations/Launcher_Script(WorkAutos).ahk` to launch these scripts.
- `Ctrl+Alt+C` launches the normal script in Firefox or the Legacy script in Chrome/Edge.
- `Ctrl+Alt+U` launches the Super-Speed script in Firefox or the Legacy script in Chrome/Edge.
- Launcher flow: close existing DSRF-to-UPS scripts, focus Intra, resize to 1530x1399, scroll to top, then launch.

## Script inventory
- `DSRF-to-UPS_WS.ahk`: stable "collect then paste" flow with Business (`Ctrl+Alt+B`) and Personal (`Ctrl+Alt+P`) hotkeys.
- `DSRF-to-UPS_WS(Super-Speed).ahk`: faster variant using `FastSleep`, retry-aware copy helpers, and address-fill polling.
- `Legacy/DSRF-to-UPS_WS-Legacy.ahk`: Chrome/Edge fallback flow with legacy coordinates.
- `Legacy/DSRF-to-UPS_WS-NewScale(Firefox+Legacy).ahk`: older reference implementation.
- `DSRF-Export/`: API-based tooling (batch export + AHK paste helper) with its own `AGENTS.md`.

## Shared mechanics (coordinate scripts)
- Normalize Intra window to 1530x1399 and use fixed window-relative coordinates for field targeting.
- Use neutral click + `Ctrl+Home` before capture to stabilize scroll/focus state.
- Capture all Intra fields first, then switch to WorldShip and paste in sequence.
- Use Firefox zoom 60% for capture and return to 100% when done.
- Clear WorldShip postal code before entry; postal update can trigger the "State/Province/County..." popup.
- Read Declared Value by tabbing from Postal Code and validating numeric format instead of direct coordinate clicks.
- Disable WorldShip electronic scale once per run to reduce lag (`scaleOffClick`).

## Editing guidelines
- Keep coordinate changes synchronized between normal and super-speed scripts unless intentionally diverging.
- Preserve clipboard save/restore patterns in copy/paste helpers.
- When changing hotkeys or launcher behavior, update:
  - `Repositories/AHK-Automations/Work_Automations/Launcher_Script(WorkAutos).ahk`
  - `Repositories/AHK-Automations/Other_Automations/ToolTips.ahk`
  - `Repositories/AHK-Automations/README.md`
- Implement in `DSRF-to-UPS_WS.ahk` first, then port to Super-Speed and Legacy variants as needed.
