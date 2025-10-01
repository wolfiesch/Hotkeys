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
- `docs/autohotkey-v2/` - Official AutoHotkey v2.0.19 documentation (CHM format)
- Root directory contains main script and reference materials
- Python utilities for documentation generation

## AutoHotkey v2 Documentation

### Local Documentation Access

The project includes the official AutoHotkey v2.0.19 documentation for offline reference:

- **Primary Reference**: `docs/autohotkey-v2/AutoHotkey.chm` - Complete AutoHotkey v2 documentation
- **License**: `docs/autohotkey-v2/AutoHotkey-license.txt` - AutoHotkey licensing terms

### Development Workflow

When working with AutoHotkey v2 code:

1. **Check local docs first**: Open `docs/autohotkey-v2/AutoHotkey.chm` for syntax, functions, and examples
2. **Common sections to reference**:
   - Language syntax and operators
   - Built-in functions and methods
   - COM object automation
   - Windows API integration
   - Error handling patterns
3. **Only search web** if local documentation doesn't cover the specific use case

### Key AutoHotkey v2 Reference Areas

- **Syntax Changes**: v2 uses different syntax from v1 (function calls require parentheses, etc.)
- **COM Automation**: Used for File Explorer integration and Windows shell operations
- **Window Detection**: `WinActive()` and window class detection patterns
- **Error Handling**: `try/catch` blocks and error object structure
- **Timing Functions**: `Sleep()`, timing constants, and delay management

## Critical Debugging Lessons

### ALWAYS Use /ErrorStdOut for Testing

**Never test AutoHotkey scripts without the `/ErrorStdOut` switch!** This provides exact error messages instead of popup dialogs:

```bash
"C:\Users\wschoenberger\AppData\Local\Programs\AutoHotkey\v2\AutoHotkey64.exe" /ErrorStdOut "ExcelDatabookLayers.ahk"
```

### Common AutoHotkey v2 Parsing Issues

1. **File Encoding**: AutoHotkey v2 prefers ANSI/ASCII encoding over UTF-8. Convert if seeing strange parsing errors.
2. **Line Endings**: Must be Windows CRLF (`\r\n`), not Unix LF (`\n`). Use `unix2dos` or similar tools.
3. **Backticks in Strings**: Backticks (`` ` ``) in strings can cause "A ':' is missing its '?'" errors as they're interpreted as incomplete ternary operators. Solutions:
   - Use `Chr(10)` instead of `` `n`` for newlines
   - Use double backticks (```` `` ````) to escape them
   - Break long strings into multiple concatenated parts
4. **Invalid Hotkey Combinations**:
   - Cannot combine modifiers with scan code syntax: `^!SC03A & h` is INVALID
   - Use simple combinations: `^!h` is VALID
   - Scan codes must be uppercase: `SC03A` not `sc03a`
5. **Semicolon in Hotkeys**: The semicolon (`;`) is problematic in hotkey definitions:
   - `` CapsLock & `;`` causes parsing errors
   - Use scan code instead: `SC03A & SC027`
   - Or use a different key entirely

### Error Diagnosis Process

1. **Run with /ErrorStdOut** to get exact line numbers and error messages
2. **Check the specific line** mentioned in the error
3. **Look for backticks, semicolons, and special characters** in that line and nearby lines
4. **Verify file encoding and line endings** if errors persist
5. **Test incrementally** by commenting out sections to isolate issues

### Testing Commands

```bash
# Check file encoding and line endings
file ExcelDatabookLayers.ahk

# Convert to Windows line endings
unix2dos ExcelDatabookLayers.ahk

# Convert UTF-8 to ANSI
powershell -Command "Get-Content -Path 'ExcelDatabookLayers.ahk' -Encoding UTF8 | Set-Content -Path 'ExcelDatabookLayers.ahk' -Encoding Default"

# Test with error output
powershell -Command "& 'C:\Users\wschoenberger\AppData\Local\Programs\AutoHotkey\v2\AutoHotkey64.exe' '/ErrorStdOut' 'ExcelDatabookLayers.ahk'"
```