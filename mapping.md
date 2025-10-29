# Hotkey Mapping Documentation

**File**: ExcelDatabookLayers.ahk
**Last Updated**: 2025-02-15

## Table of Contents

- [Overview](#overview)
- [PowerPoint Hotkeys](#powerpoint-hotkeys)
- [Excel CapsLock Layer](#excel-capslock-layer)
  - [Clipboard & Paste](#clipboard--paste)
  - [Formatting & Styling](#formatting--styling)
  - [Alignment & Borders](#alignment--borders)
  - [Layout & Sizing](#layout--sizing)
  - [Navigation & Grouping](#navigation--grouping)
  - [Navigation — Numpad Movement](#navigation--numpad-movement)
  - [Data Tools & Cleanup](#data-tools--cleanup)
  - [Clearing & Utility](#clearing--utility)
  - [Extended Layer — CapsLock + Ctrl](#extended-layer--capslock--ctrl)
  - [Extended Layer — CapsLock + Ctrl + Alt](#extended-layer--capslock--ctrl--alt)
- [Changelog](#changelog)

## Overview

This AutoHotkey automation suite layers Excel shortcuts behind CapsLock. The base layer focuses on high-frequency formatting and
navigation tasks. A secondary layer (CapsLock + Ctrl) exposes structural edits, and a new tertiary layer (CapsLock + Ctrl + Alt)
provides workbook management actions such as table creation and connection refreshes. PowerPoint sequencing macros remain
available globally.

## PowerPoint Hotkeys

*Active when PowerPoint (`POWERPNT.EXE`) is focused*

| Hotkey | Action | Description |
|--------|--------|-------------|
| `Ctrl+Alt+Shift+S` | Format Sequence | Sends: Tab, Tab, 70, Enter, Tab×4, 4.57, Tab×2, 21.47 |
| `Ctrl+Alt+Shift+F` | Format Object Pane | Alt+4 → Ctrl+A → "70" → Enter |

## Excel CapsLock Layer

*Active when Excel (`EXCEL.EXE`) is focused and CapsLock is held down*

### Clipboard & Paste

| Hotkey | Action | Description |
|--------|--------|-------------|
| `CapsLock+V` | Paste Values | Alt+H → V → V |
| `CapsLock+F` | Paste Formulas | Alt+H → V → F *(Ctrl: Freeze Panes, Ctrl+Alt: Freeze Top Row)* |
| `CapsLock+T` | Paste Formats | Alt+H → V → R |
| `CapsLock+W` | Paste Column Widths | Ctrl+Alt+V → W → Enter |
| `CapsLock+S` | Paste Formulas + Format | Alt+H → V → S → R → Enter |
| `CapsLock+X` | Paste Values (Transpose) | Paste values with transpose option |
| `CapsLock+L` | Paste Link | Ctrl+Alt+V → Alt+L → Enter |
| `CapsLock+P` | Paste Special Dialog | Ctrl+Alt+V |
| `CapsLock+Numpad/` | Toggle Filter | Alt+H → S → F |
| `CapsLock+Numpad*` | Clear Filter | Alt+H → S → C |
| `CapsLock+Numpad+` | Paste Add | Paste Special with Add operation |
| `CapsLock+Numpad-` | Paste Subtract | Paste Special with Subtract operation |

### Formatting & Styling

| Hotkey | Action | Description |
|--------|--------|-------------|
| `CapsLock+1` | Custom Font Color | Ctrl+Alt+Shift+1 (macro) |
| `CapsLock+2` | Subtotal Format | Bold + top border |
| `CapsLock+3` | Major Total Format | Bold + top border |
| `CapsLock+4` | Grand Total Format | Bold + top double border |
| `CapsLock+5` | Percent Format | Alt+H → P |
| `CapsLock+6` | Set Row Height | Set row height to 5pt |
| `CapsLock+9` | General Format | Alt+H → N → G |
| `CapsLock+A` | Set Row Height | Set row height to 5pt |
| `CapsLock+K` | Custom Number Format | Ctrl+Alt+Shift+K (macro) |
| `CapsLock+M` | Month Format | Alt+H → N → M |
| `CapsLock+D` | Date Format | Alt+H → N → D |

### Alignment & Borders

| Hotkey | Action | Description |
|--------|--------|-------------|
| `CapsLock+F1` | Align Left | Alt+H → A → L |
| `CapsLock+F2` | Align Center | Alt+H → A → C |
| `CapsLock+F3` | Align Right | Alt+H → A → R |
| `CapsLock+F4` | Toggle Wrap Text | Alt+H → W |
| `CapsLock+R` | Apply Right Border | Applies right border *(Ctrl: AutoFit Rows, Ctrl+Alt: Refresh Connections)* |
| `CapsLock+B` | Apply Bottom Border | Ctrl+Shift+B (macro) |
| `CapsLock+O` | Outline Borders | Alt+H → B → O |
| `CapsLock+I` | Inside Borders | Alt+H → B → I |
| `CapsLock+C` | Clear Borders | Alt+H → B → N *(Ctrl: AutoFit Columns, Ctrl+Alt: Create Table)* |
| `CapsLock+Y` | Top Double Border | Alt+H → B → T |
| `CapsLock+J` | Left Thick Border | Alt+H → B → L |
| `CapsLock+;` | Right Thick Border | Alt+H → B → R |

### Layout & Sizing

| Hotkey | Action | Description |
|--------|--------|-------------|
| `CapsLock+/` | Delete Column | Alt+H → D → C *(Ctrl: Delete Row, Ctrl+Alt: Insert Row Above)* |
| `CapsLock+Q` | Set Column Width Preset | Column width 0.5 *(Ctrl: Row Height 5pt, Ctrl+Alt: Column Width 15)* |
| `CapsLock+F5` | AutoFit Columns | Alt+H → O → I |
| `CapsLock+F6` | AutoFit Rows | Alt+H → O → A |
| `CapsLock+F11` | Increase Indent | Alt+H → 6 |
| `CapsLock+F12` | Decrease Indent | Alt+H → 5 |
| `CapsLock+Numpad.` | Add Decimal Place | Alt+H → 0 |
| `CapsLock+Numpad0` | Remove Decimal Place | Alt+H → 9 |

### Navigation & Grouping

| Hotkey | Action | Description |
|--------|--------|-------------|
| `CapsLock+[` | Previous Divider | Jump to previous data boundary |
| `CapsLock+]` | Next Divider | Jump to next data boundary |
| `CapsLock+=` | First Block Edge | Ctrl+Home → Ctrl+Right |
| `CapsLock+-` | Last Block Edge | Ctrl+End |
| `CapsLock+Comma` | Previous Sheet | Ctrl+PgUp |
| `CapsLock+Period` | Next Sheet | Ctrl+PgDn |
| `CapsLock+G` | Go To Dialog | Ctrl+G *(Ctrl: Group & Collapse, Ctrl+Alt: Ungroup)* |
| `CapsLock+H` | Highlight Cell | Ctrl+Alt+Shift+H |
| `CapsLock+8` | Select Current Region | Ctrl+Shift+8 |
| `CapsLock+Right` | Move Right 12 | Step 12 cells right *(Ctrl: Next Sheet)* |
| `CapsLock+Left` | Move Left 12 | Step 12 cells left *(Ctrl: Previous Sheet)* |
| `CapsLock+Ctrl+Right` | Next Sheet | Ctrl+PgDn |
| `CapsLock+Ctrl+Left` | Previous Sheet | Ctrl+PgUp |

### Navigation — Numpad Movement

| Hotkey | Action | Description |
|--------|--------|-------------|
| `CapsLock+Numpad8` | Ctrl+Up | Move to top of data block |
| `CapsLock+Numpad2` | Ctrl+Down | Move to bottom of data block |
| `CapsLock+Numpad4` | Ctrl+Left | Move to left edge |
| `CapsLock+Numpad6` | Ctrl+Right | Move to right edge |
| `CapsLock+Numpad7` | Ctrl+Home | Jump to A1 |
| `CapsLock+Numpad9` | Ctrl+End | Jump to last used cell |

### Data Tools & Cleanup

| Hotkey | Action | Description |
|--------|--------|-------------|
| `CapsLock+U` | Trim In Place | Removes leading/trailing spaces |
| `CapsLock+F8` | Clean In Place | Strips line breaks |
| `CapsLock+N` | Convert to Number | Text-to-columns number coercion |
| `CapsLock+E` | Text to Columns | Alt+D → E wizard |
| `CapsLock+F7` | Toggle AutoFilter | Ctrl+Shift+L |
| `CapsLock+H` | Highlight Cell | Custom macro (Ctrl+Alt+Shift+H) |

### Clearing & Utility

| Hotkey | Action | Description |
|--------|--------|-------------|
| `CapsLock+Z` | Clear Formats | Alt+H → E → F |
| `CapsLock+Backspace` | Clear Contents | Delete selected contents |
| `CapsLock+Delete` | Clear All | Alt+H → E → A |
| `CapsLock+Space` | Show Help Overlay | Displays on-screen reference |
| `CapsLock+;` | Keep On Top | Toggle Excel always-on-top |
| `CapsLock+Enter` | Terminal Here | Open terminal in workbook directory |
| `CapsLock+Down` | Next Tab | Browser-style tab navigation |
| `CapsLock+Up` | Previous Tab | Browser-style tab navigation |

### Extended Layer — CapsLock + Ctrl

| Hotkey | Action | Description |
|--------|--------|-------------|
| `CapsLock+Ctrl+F` | Freeze Panes | Alt+W → F → F |
| `CapsLock+Ctrl+/` | Delete Row | Alt+H → D → R |
| `CapsLock+Ctrl+R` | AutoFit Row Height | Alt+H → O → A |
| `CapsLock+Ctrl+C` | AutoFit Column Width | Alt+H → O → I |
| `CapsLock+Ctrl+Q` | Set Row Height 5pt | Alt+H → O → H → 5 → Enter |
| `CapsLock+Ctrl+G` | Group & Collapse | Alt+A → G → G followed by Alt+A → H |

### Extended Layer — CapsLock + Ctrl + Alt

| Hotkey | Action | Description |
|--------|--------|-------------|
| `CapsLock+Ctrl+Alt+F` | Freeze Top Row | Alt+W → F → R |
| `CapsLock+Ctrl+Alt+/` | Insert Row Above | Alt+H → I → R |
| `CapsLock+Ctrl+Alt+R` | Refresh All Connections | Alt+A → R → A |
| `CapsLock+Ctrl+Alt+C` | Create Table | Ctrl+T → Enter |
| `CapsLock+Ctrl+Alt+Q` | Column Width 15 | Alt+H → O → W → 15 |
| `CapsLock+Ctrl+Alt+G` | Ungroup Selection | Alt+A → U → U |

## Changelog

### 2025-02-15
- **Added**: CapsLock + Ctrl + Alt tertiary layer for freeze, table, refresh, sizing, and outline management shortcuts.
- **Refactored**: All Excel hotkeys reorganized into clipboard, styling, layout, navigation, cleanup, and utility groupings.
- **Documented**: Updated CSV/export mapping and inline descriptions to highlight layered behaviors per key family.

### 2025-01-09
- **Added**: CapsLock+Ctrl+Q hotkey for Set Row Height 5pt
- **Added**: CapsLock+Ctrl+R and CapsLock+Ctrl+C hotkeys for AutoFit operations
- **Modified**: Integrated Ctrl modifier detection into existing F, R, and C hotkeys
- **Modified**: Changed CapsLock+A from Accounting Format to Set Row Height 5pt
- **Created**: Initial comprehensive documentation of all hotkeys

### Future Updates
*This section will be updated as hotkeys are added, modified, or removed*
