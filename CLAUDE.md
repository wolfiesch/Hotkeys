# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is an AutoHotkey v2.0 project that provides layered hotkey functionality for Excel and PowerPoint automation. The main script uses CapsLock as a modifier key to create an extensive set of shortcuts for data manipulation, formatting, and navigation tasks in Excel.

## Architecture

### Core Components

- **ExcelDatabookLayers.ahk** - Main AutoHotkey script containing all hotkey definitions and automation functions
- **ExcelDatabookLayers_Hotkeys.csv** - Structured hotkey reference data
- **Python utilities** - Documentation generation scripts for Excel and PDF reference materials

### Key Design Patterns

1. **Layer System**: Uses CapsLock as primary modifier with optional Ctrl combinations for extended functionality
2. **Application Detection**: `IsExcel()` and `WinActive()` functions to ensure hotkeys only work in target applications  
3. **Ribbon Commands**: Prioritizes Excel ribbon automation over COM for reliability
4. **Error Handling**: `Do(fn, operation)` wrapper provides consistent error handling with user feedback
5. **Timing Control**: `Timing` class constants ensure consistent delays for ribbon operations

### State Management

- **No persistent state** - The script uses held-down key detection rather than toggle states
- **HUD System** - Visual feedback through tooltips and overlay GUI for user confirmation
- **Conditional Modifiers** - Ctrl detection within CapsLock layer for dual-purpose hotkeys

## Common Development Tasks

### Testing AutoHotkey Changes

The script requires AutoHotkey v2.0. To test changes:

1. Ensure AutoHotkey v2.0 is installed
2. Close any running instances of the script
3. Run the script: `ExcelDatabookLayers.ahk`
4. Test in Excel with CapsLock + key combinations

### Adding New Hotkeys

1. Add function definition in the appropriate section (PASTE, FORMAT, etc.)
2. Add hotkey binding in the `#HotIf (IsExcel() && GetKeyState("SC03A","P"))` block
3. Update documentation files (CSV, mapping.md)
4. Follow the `Do(() => FunctionName(), "Description")` pattern for consistency

### Documentation Generation

Run Python utilities to regenerate reference materials:
- `python create_styled_hotkeys.py` - Updates Excel reference file
- `python create_styled_pdf.py` - Generates PDF documentation

### Key Constants and Helpers

- `Timing.RIBBON_DELAY` (120ms) - Initial Alt key press timing
- `Timing.DIALOG_DELAY` (250ms) - Dialog interaction timing  
- `Timing.NAV_DELAY` (50ms) - Ribbon navigation timing
- `ShowHUD(msg, ms)` - Display temporary status messages
- `ShowOSD(title, body, ms)` - Rich overlay display

## Excel Integration Details

The script exclusively uses Excel's ribbon interface commands rather than COM automation for maximum compatibility. Key patterns:

- `Send("!h")` + `Wait(Timing.RIBBON_DELAY)` - Access Home ribbon
- Sequential navigation with appropriate delays between keystrokes
- Error recovery through consistent escape sequences

## File Organization

- `archive/` - Deprecated script versions and migration notes
- `docs/` - Comprehensive documentation including HOTKEYS.md
- Root directory contains main script and reference materials
- Python utilities for documentation generation