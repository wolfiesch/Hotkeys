# Excel Databook Layers - Hotkey Reference

## How Layers Work

Layers are activated by **holding down** CapsLock combinations:

- **CapsLock alone**: Hold CapsLock + key for placeholder functions
- **Ctrl+CapsLock**: Hold both keys + action key for PASTE operations
- **Shift+CapsLock**: Hold both keys + action key for FORMAT operations  
- **Alt+CapsLock**: Hold both keys + action key for NAV operations
- **Win+CapsLock**: Hold both keys + action key for DATA operations
- **CapsLock is disabled** - no accidental caps lock activation

## CapsLock Layer - CapsLock held down

| Key | Action |
|-----|--------|
| `Any letter` | Shows "Caps + [letter] (placeholder)" |

## PASTE Layer - Ctrl+CapsLock held down

| Key | Action |
|-----|--------|
| `b` or `v` | Paste Values |
| `f` | Paste Formulas |
| `t` | Paste Formats |
| `w` | Paste Column widths |
| `s` | Values Skip blanks |
| `x` | Values + Transpose |
| `l` | Paste Link |
| `p` | Paste Special dialog (fallback) |
| `NumpadMult` | Multiply operation |
| `NumpadDiv` | Divide operation |
| `NumpadAdd` | Add operation |
| `NumpadSub` | Subtract operation |
| `Down Arrow` | Propogat paste (Ctrl+Shift+Right, Ctrl+Shift+Down, Shift+Up, Paste Formulas + Format) |

## FORMAT Layer - Shift+CapsLock held down

### Row Classification
| Key | Action |
|-----|--------|
| `1` | Section header (Bold + bottom thin + gray fill + 5pt spacer above) |
| `2` | Subtotal (Bold + top thin) |
| `3` | Major total (Bold + top thin) |
| `4` | Grand total (Bold + top double + bottom medium) |
| `6` | Spacer row (height 5pt, clear borders) |

### Number Formats
| Key | Action | Format Mask |
|-----|--------|-------------|
| `n` | General | `General` |
| `a` | Accounting | `_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)` |
| `k` | Thousands | `#,##0` |
| `5` | Percent | `0.0%` |
| `m` | Month | `[$-409]mmm-yy;@` |
| `d` | Date | `yyyy-mm-dd` |

### Alignment
| Key | Action |
|-----|--------|
| `l` | Left align |
| `e` | Center align |
| `y` | Right align |
| `w` | Toggle wrap text |

### Borders
| Key | Action | Notes |
|-----|--------|-------|
| `r` | Right thin | Calls `RightBorder` macro if available, else COM |
| `b` | Bottom thin | Calls `CDS_BottomThinBorder` macro if available, else COM |
| `o` | Outline thin |
| `i` | Inside thin |
| `g` | Clear all borders |
| `t` | Top double |
| `d` | Bottom double |
| `j` | Left thick |
| `;` | Right thick |
| `Shift+o` | Outline medium |
| `Shift+i` | Inside medium |

### Dividers & Sizing
| Key | Action |
|-----|--------|
| `q` | Set column width to 0.5 |
| `Shift+q` | Insert divider column (width 0.5) |
| `v` | Set row height to 5pt |
| `Shift+v` | Insert spacer row (height 5pt) |
| `f` | AutoFit columns |
| `Shift+f` | AutoFit rows |

### Numpad Tweaks
| Key | Action |
|-----|--------|
| `NumpadAdd` | Increase indent |
| `NumpadSub` | Decrease indent |
| `NumpadDot` | Add decimal place |
| `Numpad0` | Remove decimal place |

## NAV Layer - Alt+CapsLock held down

### Divider Navigation
| Key | Action |
|-----|--------|
| `[` | Jump to previous divider column (width ≤ 1.0) |
| `]` | Jump to next divider column (width ≤ 1.0) |
| `=` | Jump to first data column in current block |
| `-` | Jump to last data column in current block |

### Numpad Navigation
| Key | Action |
|-----|--------|
| `Numpad8` | Ctrl+Up (jump to edge) |
| `Numpad2` | Ctrl+Down (jump to edge) |
| `Numpad4` | Ctrl+Left (jump to edge) |
| `Numpad6` | Ctrl+Right (jump to edge) |
| `Numpad7` | Jump to A1 |
| `Numpad9` | Jump to last cell (Ctrl+End) |
| `Numpad5` | Select current region |

### Sheet & Selection
| Key | Action |
|-----|--------|
| `,` | Previous sheet |
| `.` | Next sheet |
| `g` | Go To dialog |
| `s` | Select current region |
| `t` | Select table/data block |
| `h` | Jump to A1 |

## DATA Layer - Win+CapsLock held down

| Key | Action |
|-----|--------|
| `t` | TRIM in place |
| `c` | CLEAN in place |
| `n` | Coerce text to number |
| `e` | Text to Columns |
| `f` | Toggle AutoFilter |
| `k` | Freeze panes at active cell |
| `z` | Clear formats |
| `x` | Clear contents |
| `a` | Clear all |

## Notes
- `FORMAT r` calls `RightBorder` if available; `FORMAT b` calls `CDS_BottomThinBorder` if available. If a macro is not present or fails, the script applies the same effect via COM as a fallback.

## Acceptance Checks

1. **Right/Bottom border**: Arm FORMAT layer, press `r` then `b`. Borders apply via macro if loaded, else via COM. HUD indicates which path was used.

2. **Divider workflow**: FORMAT `Shift+q` inserts a 0.5-width divider column. NAV `]` jumps to it. `=` and `-` snap to first/last data column in the block.

3. **Classify row**: FORMAT `1`, then `6` creates a header row with styling and ensures a 5pt spacer above without duplicating if already present.

4. **Paste safety**: PASTE `s` pastes values with skip blanks enabled, leaving existing filled cells intact.
