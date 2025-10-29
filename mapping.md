# Hotkey Mapping Documentation

**File**: ExcelDatabookLayers.ahk  
**Last Updated**: 2025-10-29 (Three-key layer reorganization)

## Table of Contents

- [Overview](#overview)
- [PowerPoint Hotkeys](#powerpoint-hotkeys)
- [Excel CapsLock Layer](#excel-capslock-layer)
  - [PASTE Operations](#paste-operations)
  - [FORMAT Operations](#format-operations)
  - [ALIGNMENT](#alignment)
  - [BORDERS/MACROS](#bordersmacros)
  - [SIZING](#sizing)
  - [NAVIGATION](#navigation)
  - [DATA/CLEANUP](#datacleanup)
  - [HELP](#help)
- [Three-Key Layer (Ctrl+Alt+CapsLock)](#three-key-layer-ctrlaltcapslock)
  - [Data Tools](#data-tools)
  - [Layout & Structure](#layout--structure)
  - [Cleanup & Clears](#cleanup--clears)
- [Special Modifier Combinations](#special-modifier-combinations)
- [Changelog](#changelog)

## Overview

This AutoHotkey script provides Excel and PowerPoint automation through layered hotkeys. The primary layer uses CapsLock as a modifier key when Excel is focused, while metadata-driven Ctrl and Ctrl+Alt combinations extend the map with grouped data, layout, and cleanup workflows.

## PowerPoint Hotkeys

*Active when PowerPoint (POWERPNT.EXE) is focused*

| Hotkey | Action | Description |
|--------|--------|-------------|
| `Ctrl+Alt+Shift+S` | Format Sequence | Sends: Tab, Tab, 70, Enter, Tab×4, 4.57, Tab×2, 21.47 |
| `Ctrl+Alt+Shift+F` | Format Object Pane | Alt+4 → Ctrl+A → "70" → Enter |

## Excel CapsLock Layer

*Active when Excel (EXCEL.EXE) is focused and CapsLock is held down*

### PASTE Operations

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
| `CapsLock+Numpad/` | Toggle Filter | Alt+H → S → F |
| `CapsLock+Numpad*` | Clear Filter | Alt+H → S → C |
| `CapsLock+Numpad+` | Paste Add | Paste Special with Add operation |
| `CapsLock+Numpad-` | Paste Subtract | Paste Special with Subtract operation |

### FORMAT Operations

| Hotkey | Action | Description |
|--------|--------|-------------|
| `CapsLock+1` | Custom Font Color | Ctrl+Alt+Shift+1 (macro) |
| `CapsLock+2` | Subtotal Format | Bold + top border |
| `CapsLock+3` | Major Total Format | Bold + top border |
| `CapsLock+4` | Grand Total Format | Bold + top double border |
| `CapsLock+6` | Set Row Height | Set row height to 5pt |
| `CapsLock+9` | General Format | Alt+H → N → G |
| `CapsLock+A` | Set Row Height | Set row height to 5pt |
| `CapsLock+K` | Custom Number Format | Ctrl+Alt+Shift+K (macro) |
| `CapsLock+5` | Percent Format | Alt+H → P |
| `CapsLock+M` | Month Format | Alt+H → N → M |
| `CapsLock+D` | Date Format | Alt+H → N → D |

### ALIGNMENT

| Hotkey | Action | Description |
|--------|--------|-------------|
| `CapsLock+F1` | Align Left | Alt+H → AL |
| `CapsLock+F2` | Align Center | Alt+H → AC |
| `CapsLock+F3` | Align Right | Alt+H → AR |
| `CapsLock+F4` | Toggle Wrap Text | Alt+H → W |

### BORDERS/MACROS

| Hotkey | Action | Description |
|--------|--------|-------------|
| `CapsLock+R` | Apply Right Border | Ctrl+Shift+R (macro) |
| `CapsLock+B` | Apply Bottom Border | Ctrl+Shift+B (macro) |
| `CapsLock+O` | Outline Borders | Alt+H → B → O |
| `CapsLock+I` | Inside Borders | Alt+H → B → I |
| `CapsLock+C` | Clear Borders | Alt+H → B → N |
| `CapsLock+Y` | Top Double Border | Alt+H → B → T |
| `CapsLock+J` | Left Thick Border | Alt+H → B → L |
| `CapsLock+;` | Right Thick Border | Alt+H → B → R |

### SIZING

| Hotkey | Action | Description |
|--------|--------|-------------|
| `CapsLock+Q` | Set Column Width | Set column width to 0.5 |
| `CapsLock+Ctrl+Q` | Set Row Height | Set row height to 5pt |
| `CapsLock+F5` | AutoFit Columns | Alt+H → O → I |
| `CapsLock+F6` | AutoFit Rows | Alt+H → O → A |
| `CapsLock+F11` | Increase Indent | Alt+H → 6 |
| `CapsLock+F12` | Decrease Indent | Alt+H → 5 |
| `CapsLock+Numpad.` | Add Decimal Place | Alt+H → 00 |
| `CapsLock+Numpad0` | Remove Decimal Place | Alt+H → 9 |

### NAVIGATION

| Hotkey | Action | Description |
|--------|--------|-------------|
| `CapsLock+[` | Jump to Previous Divider | Ctrl+Left |
| `CapsLock+]` | Jump to Next Divider | Ctrl+Right |
| `CapsLock+=` | Jump to First Block Edge | Ctrl+Home → Ctrl+Right |
| `CapsLock+-` | Jump to Last Block Edge | Ctrl+End |
| `CapsLock+,` | Previous Sheet | Ctrl+PgUp |
| `CapsLock+.` | Next Sheet | Ctrl+PgDn |
| `CapsLock+G` | Go To Dialog | Ctrl+G |
| `CapsLock+8` | Select Current Region | Ctrl+Shift+8 |
| `CapsLock+H` | Custom Macro | Ctrl+Alt+Shift+H |
| `CapsLock+Numpad8` | Ctrl+Up | Navigate to edge of data block up |
| `CapsLock+Numpad2` | Ctrl+Down | Navigate to edge of data block down |
| `CapsLock+Numpad4` | Ctrl+Left | Navigate to edge of data block left |
| `CapsLock+Numpad6` | Ctrl+Right | Navigate to edge of data block right |
| `CapsLock+Numpad7` | Ctrl+Home | Navigate to A1 |
| `CapsLock+Numpad9` | Ctrl+End | Navigate to last used cell |

### DATA/CLEANUP

| Hotkey | Action | Description |
|--------|--------|-------------|
| `CapsLock+U` | Trim In Place | Find & Replace to remove leading spaces |
| `CapsLock+F8` | Clean In Place | Find & Replace to clean line breaks |
| `CapsLock+N` | Convert to Number | Text to Columns to force number conversion |
| `CapsLock+E` | Text to Columns | Alt+D → E |
| `CapsLock+F7` | Toggle AutoFilter | Ctrl+Shift+L |
| `CapsLock+Z` | Clear Formats | Alt+H → E → F |
| `CapsLock+Backspace` | Clear Contents | Delete key |
| `CapsLock+Delete` | Clear All | Alt+H → E → A |

> **Tip:** Every data, layout, and cleanup command in this table is mirrored on the new Ctrl+Alt+CapsLock layer so that related tools can be executed from a dedicated three-modifier posture.

### HELP

| Hotkey | Action | Description |
|--------|--------|-------------|
| `CapsLock+Space` | Show Help Overlay | Display on-screen help with all hotkeys |

## Three-Key Layer (Ctrl+Alt+CapsLock)

Hold **CapsLock+Ctrl+Alt** to access a reorganized "power" layer that groups related Excel tooling. The tables below mirror the code-backed metadata so overlay descriptions stay in sync with these combinations.

### Data Tools

| Hotkey | Action | Description |
|--------|--------|-------------|
| `CapsLock+Ctrl+Alt+T` | Trim In Place | Remove leading/trailing spaces via find & replace |
| `CapsLock+Ctrl+Alt+C` | Clean In Place | Strip non-printable characters |
| `CapsLock+Ctrl+Alt+N` | Convert to Number | Force numeric coercion using Text-to-Columns |
| `CapsLock+Ctrl+Alt+E` | Text to Columns | Launch delimiter wizard (Alt+D → E) |
| `CapsLock+Ctrl+Alt+F7` | Toggle AutoFilter | Ctrl+Shift+L |

### Layout & Structure

| Hotkey | Action | Description |
|--------|--------|-------------|
| `CapsLock+Ctrl+Alt+F` | Freeze Panes | Alt+W → F → F |
| `CapsLock+Ctrl+Alt+A` | AutoFit Columns | AutoFit selected columns |
| `CapsLock+Ctrl+Alt+R` | AutoFit Rows | AutoFit selected rows |
| `CapsLock+Ctrl+Alt+W` | Set Column Width | Set width to 0.5 |
| `CapsLock+Ctrl+Alt+H` | Set Row Height | Set height to 5pt |
| `CapsLock+Ctrl+Alt+G` | Group & Collapse | Outline the selection and collapse |
| `CapsLock+Ctrl+Alt+/` | Delete Row | Remove the active sheet row |

### Cleanup & Clears

| Hotkey | Action | Description |
|--------|--------|-------------|
| `CapsLock+Ctrl+Alt+Z` | Clear Formats | Alt+H → E → F |
| `CapsLock+Ctrl+Alt+Backspace` | Clear Contents | Delete cell contents |
| `CapsLock+Ctrl+Alt+Delete` | Clear All | Alt+H → E → A |

## Special Modifier Combinations

*These hotkeys require CapsLock plus additional modifiers*

### CapsLock + Ctrl (Legacy quick actions)

| Hotkey | Action | Description |
|--------|--------|-------------|
| `CapsLock+Ctrl+F` | Freeze Panes | Alt+W → F → F |
| `CapsLock+Ctrl+R` | AutoFit Row Height | Alt+H → O → A |
| `CapsLock+Ctrl+C` | AutoFit Column Width | Alt+H → O → I |
| `CapsLock+Ctrl+Q` | Set Row Height | Alt+H → O → H → 5 → Enter |
| `CapsLock+Ctrl+G` | Group & Collapse | Outline selection |
| `CapsLock+Ctrl+/` | Delete Row | Remove active row |
| `CapsLock+Ctrl+Right` | Next Sheet | Ctrl+PgDn |
| `CapsLock+Ctrl+Left` | Previous Sheet | Ctrl+PgUp |

These remain available for muscle memory compatibility but now coexist with the richer [three-key layer](#three-key-layer-ctrlaltcapslock), which groups the same operations alongside related cleanup tools.

## Changelog

### 2025-10-29
- **Added**: CapsLock+Ctrl+Alt three-key layer grouping data tools, layout utilities, and cleanup actions
- **Added**: Metadata-driven registration for CapsLock+Ctrl and CapsLock+Ctrl+Alt combinations with overlay descriptions
- **Updated**: Base CapsLock layer now ignores Ctrl/Alt modifiers so stacked layers trigger reliably
- **Documented**: New tables highlighting grouped workflows and refreshed quick reference content

### 2025-01-09
- **Added**: CapsLock+Ctrl+Q hotkey for Set Row Height 5pt
- **Added**: CapsLock+Ctrl+R and CapsLock+Ctrl+C hotkeys for AutoFit operations
- **Modified**: Integrated Ctrl modifier detection into existing F, R, and C hotkeys
- **Modified**: Changed CapsLock+A from Accounting Format to Set Row Height 5pt
- **Created**: Initial comprehensive documentation of all hotkeys

### Future Updates
*This section will be updated as hotkeys are added, modified, or removed*