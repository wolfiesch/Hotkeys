# Hotkey Mapping Documentation

**File**: ExcelDatabookLayers.ahk  
**Last Updated**: 2025-10-29 (Formatting layer reorganization)

## Table of Contents

- [Overview](#overview)
- [PowerPoint Hotkeys](#powerpoint-hotkeys)
- [Excel Layer Overview](#excel-layer-overview)
  - [CapsLock Base Layer – Paste & Navigation](#capslock-base-layer--paste--navigation)
    - [Paste & Clipboard](#paste--clipboard)
    - [Filters & Operations](#filters--operations)
    - [Navigation & Selection](#navigation--selection)
    - [Data Cleanup & Clears](#data-cleanup--clears)
    - [Help Overlay](#help-overlay)
  - [CapsLock+Ctrl Layout Layer](#capslockctrl-layout-layer)
  - [CapsLock+Ctrl+Alt Formatting Layer](#capslockctrlalt-formatting-layer)
- [Special Modifier Combinations](#special-modifier-combinations)
- [Changelog](#changelog)

## Overview

This AutoHotkey script provides Excel and PowerPoint automation through layered hotkeys. The primary layer uses CapsLock as a modifier key when Excel is focused, with additional Ctrl combinations for extended functionality.

## PowerPoint Hotkeys

*Active when PowerPoint (POWERPNT.EXE) is focused*

| Hotkey | Action | Description |
|--------|--------|-------------|
| `Ctrl+Alt+Shift+S` | Format Sequence | Sends: Tab, Tab, 70, Enter, Tab×4, 4.57, Tab×2, 21.47 |
| `Ctrl+Alt+Shift+F` | Format Object Pane | Alt+4 → Ctrl+A → "70" → Enter |

## Excel Layer Overview

*Active when Excel (EXCEL.EXE) is focused and CapsLock is held down*

### CapsLock Base Layer – Paste & Navigation

#### Paste & Clipboard

| Hotkey | Action | Description |
|--------|--------|-------------|
| `CapsLock+V` | Paste Values | Alt+H → V → V |
| `CapsLock+F` | Paste Formulas | Alt+H → V → F |
| `CapsLock+T` | Paste Formats | Alt+H → V → R |
| `CapsLock+W` | Paste Column Widths | Ctrl+Alt+V → W → Enter |
| `CapsLock+S` | Paste Formulas + Format | Alt+H → V → S → R → Enter |
| `CapsLock+X` | Paste Values (Transpose) | Paste values with transpose option |
| `CapsLock+L` | Paste Link | Ctrl+Alt+V → Alt+L → Enter |
| `CapsLock+P` | Paste Special Dialog | Ctrl+Alt+V |

#### Filters & Operations

| Hotkey | Action | Description |
|--------|--------|-------------|
| `CapsLock+/` | Delete Column | Alt+H → D → C |
| `CapsLock+Numpad/` | Toggle Filter | Alt+H → S → F |
| `CapsLock+Numpad*` | Clear Filter | Alt+H → S → C |
| `CapsLock+Numpad+` | Paste Add | Paste Special with Add operation |
| `CapsLock+Numpad-` | Paste Subtract | Paste Special with Subtract operation |

#### Navigation & Selection

| Hotkey | Action | Description |
|--------|--------|-------------|
| `CapsLock+[` | Jump to Previous Divider | Ctrl+Left |
| `CapsLock+]` | Jump to Next Divider | Ctrl+Right |
| `CapsLock+=` | Jump to First Block Edge | Ctrl+Home → Ctrl+Right |
| `CapsLock+-` | Jump to Last Block Edge | Ctrl+End |
| `CapsLock+,` | Previous Sheet | Ctrl+PgUp |
| `CapsLock+.` | Next Sheet | Ctrl+PgDn |
| `CapsLock+Right` | Move Right 12 cells | `{Right 12}` |
| `CapsLock+Left` | Move Left 12 cells | `{Left 12}` |
| `CapsLock+Shift+Right` | Select 11 cells to the right | Shift+Right×11 |
| `CapsLock+Shift+Left` | Select 11 cells to the left | Shift+Left×11 |
| `CapsLock+8` | Select Current Region | Ctrl+Shift+8 |
| `CapsLock+G` | Go To Dialog | Ctrl+G |
| `CapsLock+Numpad8` | Ctrl+Up | Navigate to top of data block |
| `CapsLock+Numpad2` | Ctrl+Down | Navigate to bottom of data block |
| `CapsLock+Numpad4` | Ctrl+Left | Navigate to left edge |
| `CapsLock+Numpad6` | Ctrl+Right | Navigate to right edge |
| `CapsLock+Numpad7` | Ctrl+Home | Navigate to A1 |
| `CapsLock+Numpad9` | Ctrl+End | Navigate to last used cell |

#### Data Cleanup & Clears

| Hotkey | Action | Description |
|--------|--------|-------------|
| `CapsLock+U` | Trim In Place | Removes leading/trailing spaces |
| `CapsLock+F8` | Clean In Place | Removes line breaks |
| `CapsLock+N` | Convert to Number | Text to Columns to force number conversion |
| `CapsLock+E` | Text to Columns | Alt+D → E |
| `CapsLock+F7` | Toggle AutoFilter | Ctrl+Shift+L |
| `CapsLock+H` | Highlight Cell | Ctrl+Alt+Shift+H macro |
| `CapsLock+Z` | Clear Formats | Alt+H → E → F |
| `CapsLock+Backspace` | Clear Contents | Delete key |
| `CapsLock+Delete` | Clear All | Alt+H → E → A |

#### Help Overlay

| Hotkey | Action | Description |
|--------|--------|-------------|
| `CapsLock+Space` | Show Help Overlay | Displays on-screen reference for all layers |

### CapsLock+Ctrl Layout Layer

| Hotkey | Action | Description |
|--------|--------|-------------|
| `CapsLock+Ctrl+F` | Freeze Panes | Alt+W → F → F |
| `CapsLock+Ctrl+R` | AutoFit Row Height | Alt+H → O → A |
| `CapsLock+Ctrl+C` | AutoFit Column Width | Alt+H → O → I |
| `CapsLock+Ctrl+Q` | Set Row Height | Alt+H → O → H → 5 → Enter |
| `CapsLock+Ctrl+Shift+Q` | Set Column Width | Alt+H → O → W → 0.5 |
| `CapsLock+Ctrl+/` | Delete Row | Alt+H → D → R |
| `CapsLock+Ctrl+G` | Group & Collapse | Alt+A → G → G, then Alt+A → H |
| `CapsLock+Ctrl+Right` | Next Sheet | Ctrl+PgDn |
| `CapsLock+Ctrl+Left` | Previous Sheet | Ctrl+PgUp |

### CapsLock+Ctrl+Alt Formatting Layer

#### Number Formats

| Hotkey | Action | Description |
|--------|--------|-------------|
| `CapsLock+Ctrl+Alt+1` | General Format | Alt+H → N → G |
| `CapsLock+Ctrl+Alt+2` | Accounting Format | Alt+H → N → A |
| `CapsLock+Ctrl+Alt+3` | Thousands Format | Alt+H → 0 |
| `CapsLock+Ctrl+Alt+4` | Percent Format | Alt+H → P |
| `CapsLock+Ctrl+Alt+5` | Date Format | Alt+H → N → D |
| `CapsLock+Ctrl+Alt+6` | Month Format | Alt+H → N → M |

#### Row Styling Macros

| Hotkey | Action | Description |
|--------|--------|-------------|
| `CapsLock+Ctrl+Alt+A` | Subtotal Format | Applies subtotal styling macro |
| `CapsLock+Ctrl+Alt+S` | Major Total Format | Applies major total styling |
| `CapsLock+Ctrl+Alt+D` | Grand Total Format | Applies grand total styling |
| `CapsLock+Ctrl+Alt+F` | Custom Font Color | Ctrl+Alt+Shift+1 macro |

#### Border Toolkit

| Hotkey | Action | Description |
|--------|--------|-------------|
| `CapsLock+Ctrl+Alt+Q` | Outline Borders | Alt+H → B → O |
| `CapsLock+Ctrl+Alt+W` | Inside Borders | Alt+H → B → I |
| `CapsLock+Ctrl+Alt+E` | Clear Borders | Alt+H → B → N |
| `CapsLock+Ctrl+Alt+B` | Bottom Border | Ctrl+Shift+B macro |
| `CapsLock+Ctrl+Alt+R` | Right Border | Ctrl+Shift+R macro |
| `CapsLock+Ctrl+Alt+T` | Top Double Border | Alt+H → B → T |
| `CapsLock+Ctrl+Alt+Y` | Left Thick Border | Alt+H → B → L |

#### Alignment & Indent

| Hotkey | Action | Description |
|--------|--------|-------------|
| `CapsLock+Ctrl+Alt+F1` | Align Left | Alt+H → A → L |
| `CapsLock+Ctrl+Alt+F2` | Align Center | Alt+H → A → C |
| `CapsLock+Ctrl+Alt+F3` | Align Right | Alt+H → A → R |
| `CapsLock+Ctrl+Alt+F4` | Toggle Wrap Text | Alt+H → W |
| `CapsLock+Ctrl+Alt+F11` | Increase Indent | Alt+H → 6 |
| `CapsLock+Ctrl+Alt+F12` | Decrease Indent | Alt+H → 5 |

#### Decimal Precision

| Hotkey | Action | Description |
|--------|--------|-------------|
| `CapsLock+Ctrl+Alt+Numpad.` | Add Decimal Place | Alt+H → 00 |
| `CapsLock+Ctrl+Alt+Numpad0` | Remove Decimal Place | Alt+H → 9 |

## Special Modifier Combinations

The Excel layer now uses a tiered approach:

- **CapsLock** alone focuses on paste, filtering, navigation, and cleanup.
- **CapsLock + Ctrl** handles sheet layout utilities (freeze panes, AutoFit, sizing, and grouping).
- **CapsLock + Ctrl + Alt** groups all formatting and styling shortcuts, keeping number formats, row styling macros, borders, alignment, and decimal precision together.

## Changelog

### 2025-10-29
- **Added**: CapsLock+Ctrl+Alt formatting layer with grouped number formats, row styles, borders, alignment, and decimal tools.
- **Updated**: Base CapsLock layer to focus on paste, navigation, and cleanup operations only.
- **Updated**: CapsLock+Ctrl layout layer with a dedicated column-width shortcut (CapsLock+Ctrl+Shift+Q).
- **Docs**: Reorganized tables to reflect the three-tier modifier structure.

### 2025-01-09
- **Added**: CapsLock+Ctrl+Q hotkey for Set Row Height 5pt
- **Added**: CapsLock+Ctrl+R and CapsLock+Ctrl+C hotkeys for AutoFit operations
- **Modified**: Integrated Ctrl modifier detection into existing F, R, and C hotkeys
- **Modified**: Changed CapsLock+A from Accounting Format to Set Row Height 5pt
- **Created**: Initial comprehensive documentation of all hotkeys

### Future Updates
*This section will be updated as hotkeys are added, modified, or removed*