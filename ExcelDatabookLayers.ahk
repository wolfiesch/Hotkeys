#Requires AutoHotkey v2.0
#SingleInstance Force
#Warn All, MsgBox
#UseHook

; =============================================================================
; MAIN FILE - ExcelDatabookLayers.ahk
; =============================================================================
; This is the primary AutoHotkey file for Excel automation.
; Other files have been moved to archive/ subdirectory and are deprecated.
; =============================================================================

; Keep CapsLock off
SetCapsLockState("AlwaysOff")

; Treat CapsLock as a global modifier
SC03A::Return
SC03A up::Return

; Set consistent key timing for better Excel compatibility
SetKeyDelay(50, 50)  ; 50ms press duration, 50ms release delay

; Multi-application support: Excel (CapsLock layers) + PowerPoint (Ctrl+Alt+Shift+S)

; Active when Excel or PowerPoint is focused

; -----------------------------------------------------------------------------
; PowerPoint CapsLock Integration
; -----------------------------------------------------------------------------
; Only active when PowerPoint is focused
#HotIf WinActive("ahk_exe POWERPNT.EXE")

; Ctrl+Alt+Shift+S → Send specific sequence: Tab, Tab, 70, Enter, Tab x4, 4.57, Tab x2, 21.47
^!+s::
{
    ; ensure PowerPoint is focused
    WinActivate("ahk_exe POWERPNT.EXE")
    Sleep(Timing.DIALOG_DELAY)

    ; sequence: Tab, Tab, 70, Enter, Tab x4, 4.57, Tab x2, 21.47
    Send("{Tab}{Tab}")
    Sleep(60)
    Send("70{Enter}")
    Sleep(60)
    Send("{Tab 4}")
    Sleep(60)
    Send("4.57")
    Sleep(60)
    Send("{Tab 2}")
    Sleep(60)
    Send("21.47")
    return
}

; Ctrl+Alt+Shift+F → Format Object Pane Macro (Alt+4, wait, Ctrl+A, wait, "70", Enter)
^!+f::
{
    ; ensure PowerPoint is focused
    WinActivate("ahk_exe POWERPNT.EXE")
    Sleep(Timing.DIALOG_DELAY)

    ; Open format object pane with Alt+4
    Send("!4")
    Sleep(Timing.SLOW_DELAY)

    ; Select all with Ctrl+A
    Send("^a")
    Sleep(Timing.SLOW_DELAY)

    ; Send "70" and Enter
    Send("70")
    Send("{Enter}")
    return
}

#HotIf

; -----------------------------------------------------------------------------
; Chrome CapsLock Integration
; -----------------------------------------------------------------------------
; Only active when Chrome is focused and CapsLock is held
#HotIf (IsChrome() && GetKeyState("SC03A","P"))

; CapsLock + Left Arrow → Previous tab (Ctrl+Shift+Tab)
SC03A & Left::Send("^+{Tab}")

; CapsLock + Right Arrow → Next tab (Ctrl+Tab)
SC03A & Right::Send("^{Tab}")

#HotIf

; -----------------------------------------------------------------------------
; Cursor CapsLock Integration
; -----------------------------------------------------------------------------
; Only active when Cursor is focused and CapsLock is held
#HotIf (IsCursor() && GetKeyState("SC03A","P"))

; CapsLock + M → Insert commit message template and press Enter
SC03A & m::
{
    Send("^k")  ; Ctrl+K to open AI command palette
    Send("Generate a commit message summarizing recent changes")
    Send("{Enter}")
}

#HotIf

; -----------------------------------------------------------------------------
; Excel base overrides
; -----------------------------------------------------------------------------
#HotIf IsExcel()

F1::Send("{F2}")    ; Use F2 edit key while Excel is focused

#HotIf

; -----------------------------------------------------------------------------
; Single CapsLock Hold Layer Map
; -----------------------------------------------------------------------------
; PASTE: v,f,t,w,s,x,l,p + Numpad+ -
; FILTER: Numpad/ + Numpad*
; DELETE: / (Ctrl=Row)
; FORMAT: 1,2,3,4,6 + 9,a,k,5,m,d + F1-F4 + r,b,o,i,c,y,j,; + q,F5,F6,F11,F12 + Numpad.,0
; NAV: [,],=,-,.,,g,8,h,Right,Left + Numpad8,2,4,6,7,9
; DATA: u,F8,n,e,F7,F9 + z,Backspace,Delete
; -----------------------------------------------------------------------------

; Performance optimization
KeyDelay := 40
Wait(ms := KeyDelay) => Sleep(ms)

; Timing constants for consistent delays
class Timing {
    static RIBBON_DELAY := 120    ; Initial Alt key press (recommended 100-200ms)
    static DIALOG_DELAY := 250    ; Dialog interactions (recommended 250-500ms)
    static FAST_DELAY := 40       ; Quick operations
    static SLOW_DELAY := 500      ; Slow operations
    static NAV_DELAY := 50        ; Subsequent ribbon navigation (30-75ms)
}

; -----------------------------------------------------------------------------
; State - No longer needed for held-down system
; -----------------------------------------------------------------------------

; -----------------------------------------------------------------------------
; Helpers
; -----------------------------------------------------------------------------
IsExcel() => WinActive("ahk_exe EXCEL.EXE")
IsChrome() => WinActive("ahk_exe chrome.exe")
IsCursor() => WinActive("ahk_exe Cursor.exe")

; COM functions removed - using ribbon commands only

; --- Short HUD (ToolTip) ---
ShowHUD(msg, ms := 1200, x := "", y := "", slot := 20) {
    if (x = "" || y = "") {
        ToolTip(msg, , , slot)
    } else {
        ToolTip(msg, x, y, slot)
    }
    SetTimer(() => ToolTip("", , , slot), -ms)
}

; --- Rich HUD (OSD GUI) ---
global HUD := { gui: 0, timer: 0 }

ShowOSD(title := "", body := "", ms := 1600, anchor := "top-center", width := 560) {
    try HUD.gui.Destroy()
    HUD.gui := Gui("+AlwaysOnTop -Caption +ToolWindow +E0x20 -DPIScale", "")
    HUD.gui.BackColor := "0x111111"
    HUD.gui.MarginX := 12, HUD.gui.MarginY := 10
    HUD.gui.SetFont("s10", "Consolas")
    HUD.gui.AddText("xm ym cWhite", title)
    HUD.gui.AddText("xm cSilver w" width, body)
    WinSetTransparent(220, HUD.gui.Hwnd)  ; 0..255: 255=opaque

    pos :=
    (anchor = "cursor")      ? CursorPosString()
    : (anchor = "top-center")? "xCenter y20"
    : (anchor = "center")    ? "xCenter yCenter"
    :                          "x20 y20"

    HUD.gui.Show("NoActivate " . pos)

    if (HUD.timer)
        SetTimer(HUD.timer, 0)
    HUD.timer := (*) => (HUD.gui.Hide())
    SetTimer(HUD.timer, -ms)
}

CursorPosString() {
    MouseGetPos &mx, &my
    return "x" (mx+16) " y" (my+24)
}

; -----------------------------------------------------------------------------
; Layer System - Now using held-down keys
; -----------------------------------------------------------------------------
Do(fn, operation := "Unknown Operation") {
    try {
        fn()
    } catch as err {
        ShowHUD("Error in " . operation . ": " . err.Message, 3000)
    }
}



; -----------------------------------------------------------------------------
; Paste Operations
; -----------------------------------------------------------------------------
PasteSpecial(kind, opts := Map()) {
    ; Use ribbon commands - most reliable method
    if kind = "values" {
        Send("!h")      ; Home ribbon
        Wait(Timing.RIBBON_DELAY)
        Send("v")       ; Paste dropdown
        Wait(Timing.NAV_DELAY)
        Send("v")       ; Values
        ShowHUD("Paste Values", 800)
    } else if kind = "formulas" {
        Send("!h")      ; Home ribbon  
        Wait(Timing.RIBBON_DELAY)
        Send("v")       ; Paste dropdown
        Wait(Timing.NAV_DELAY)
        Send("f")       ; Formulas
        ShowHUD("Paste Formulas", 800)
    } else if kind = "formats" {
        Send("!h")      ; Home ribbon
        Wait(Timing.RIBBON_DELAY)
        Send("v")       ; Paste dropdown
        Wait(Timing.NAV_DELAY)
        Send("r")       ; Formatting
        ShowHUD("Paste Formats", 800)
    } else if kind = "colwidths" {
        Send("^!v")     ; Fallback to dialog for column widths
        Wait(100)
        Send("w{Enter}")
        ShowHUD("Paste Column Widths", 800)
    } else {
        Send("^v")      ; Regular paste
        ShowHUD("Paste All", 800)
    }
}

PasteOperation(op) {
    ; Use Paste Special dialog for operations
    Send("^!v")     ; Paste Special dialog
    Wait(Timing.DIALOG_DELAY)
    Send("v")       ; Values
    Wait(50)
    
    ; Select operation
    switch op {
        case "add":
            Send("!a")      ; Add
        case "subtract":
            Send("!s")      ; Subtract
        case "multiply":
            Send("!m")      ; Multiply
        case "divide":
            Send("!d")      ; Divide
        default:
            Send("{Escape}")
            return
    }
    Wait(50)
    Send("{Enter}")
    ShowHUD("Paste " . op, 800)
}

PasteLink() {
    ; Use Paste Special dialog for paste link
    Send("^!v")     ; Paste Special dialog
    Wait(Timing.DIALOG_DELAY)
    Send("!l")      ; Paste Link
    Wait(50)
    Send("{Enter}")
    ShowHUD("Paste Link", 800)
}

PasteFormulasWithFormat() {
    ; Alt + H → V → S → R → Enter
    Send("!h")      ; Home ribbon
    Wait(Timing.RIBBON_DELAY)
    Send("v")       ; Paste dropdown
    Wait(Timing.NAV_DELAY)
    Send("s")       ; Paste Special
    Wait(Timing.DIALOG_DELAY)
    Send("r")       ; Formats
    Wait(50)
    Send("{Enter}")
    ShowHUD("Paste Formulas + Format", 800)
}

; Filter operations
ToggleFilter() {
    ; Alt + H → S → F (Toggle Filter)
    Send("!h")      ; Home ribbon
    Wait(Timing.RIBBON_DELAY)
    Send("s")       ; Sort & Filter
    Wait(Timing.NAV_DELAY)
    Send("f")       ; Filter
    ShowHUD("Toggle Filter", 800)
}

ClearFilter() {
    ; Alt + H → S → C (Clear Filter)
    Send("!h")      ; Home ribbon
    Wait(Timing.RIBBON_DELAY)
    Send("s")       ; Sort & Filter
    Wait(Timing.NAV_DELAY)
    Send("c")       ; Clear
    ShowHUD("Clear Filter", 800)
}

DeleteSheetColumn() {
    ; Alt + H, D, C (Delete Sheet Columns)
    Send("!h")      ; Home ribbon
    Wait(Timing.RIBBON_DELAY)
    Send("d")       ; Delete menu
    Wait(Timing.NAV_DELAY)
    Send("c")       ; Delete sheet columns
    ShowHUD("Delete Column", 800)
}

DeleteSheetRow() {
    ; Alt + H, D, R (Delete Sheet Rows)
    Send("!h")      ; Home ribbon
    Wait(Timing.RIBBON_DELAY)
    Send("d")       ; Delete menu
    Wait(Timing.NAV_DELAY)
    Send("r")       ; Delete sheet rows
    ShowHUD("Delete Row", 800)
}

GroupAndCollapseSelection() {
    ; Alt + A, G, G then Alt + A, H to collapse group
    Send("!agg")
    Wait(Timing.DIALOG_DELAY)
    Send("!ah")
}

; -----------------------------------------------------------------------------
; Number Formats
; -----------------------------------------------------------------------------
SetNumberFormat(fmt) {
    ; Use ribbon commands for number formatting
    switch fmt {
        case "general":
            Send("!h")      ; Home ribbon
            Wait(Timing.RIBBON_DELAY)
            Send("n")       ; Number format dropdown
            Wait(Timing.NAV_DELAY)
            Send("g")       ; General
            ShowHUD("General Format", 800)
        case "accounting":
            Send("!h")      ; Home ribbon
            Wait(Timing.RIBBON_DELAY)
            Send("n")       ; Number format dropdown
            Wait(Timing.NAV_DELAY)
            Send("a")       ; Accounting
            ShowHUD("Accounting Format", 800)
        case "thousands":
            Send("!h")      ; Home ribbon
            Wait(Timing.RIBBON_DELAY)
            Send("0")       ; Add comma (thousands separator)
            ShowHUD("Thousands Format", 800)
        case "percent":
            Send("!h")      ; Home ribbon
            Wait(Timing.RIBBON_DELAY)
            Send("p")       ; Percent
            ShowHUD("Percent Format", 800)
        case "date":
            Send("!h")      ; Home ribbon
            Wait(Timing.RIBBON_DELAY)
            Send("n")       ; Number format dropdown
            Wait(Timing.NAV_DELAY)
            Send("d")       ; Date
            ShowHUD("Date Format", 800)
        case "month":
            Send("!h")      ; Home ribbon
            Wait(Timing.RIBBON_DELAY)
            Send("n")       ; Number format dropdown
            Wait(Timing.NAV_DELAY)
            Send("m")       ; Month
            ShowHUD("Month (mmm-yy)", 800)
        default:
            ShowHUD("Format not supported: " . fmt, 1500)
    }
}

; -----------------------------------------------------------------------------
; Borders
; -----------------------------------------------------------------------------
SetBorders(kind) {
    ; Use ribbon commands for borders
    if kind = "outline" {
        Send("!h")      ; Home ribbon
        Wait(Timing.RIBBON_DELAY)
        Send("b")       ; Borders dropdown
        Wait(Timing.NAV_DELAY)
        Send("o")       ; Outline
        ShowHUD("Outline Borders", 800)
    } else if kind = "inside" {
        Send("!h")      ; Home ribbon
        Wait(Timing.RIBBON_DELAY)
        Send("b")       ; Borders dropdown
        Wait(Timing.NAV_DELAY)
        Send("i")       ; Inside borders
        ShowHUD("Inside Borders", 800)
    }
}

SetBorderLine(side, style) {
    ; Use ribbon commands for specific border lines
    Send("!h")      ; Home ribbon
    Wait(Timing.RIBBON_DELAY)
    Send("b")       ; Borders dropdown
    Wait(Timing.NAV_DELAY)
    
    ; Map side and style to ribbon commands
    if side = "top" && style = "double" {
        Send("t")   ; Top border (will use current style)
        ShowHUD("Top Double Border", 800)
    } else if side = "bottom" && style = "double" {
        Send("b")   ; Bottom border
        ShowHUD("Bottom Double Border", 800)
    } else if side = "bottom" && style = "medium" {
        Send("b")   ; Bottom border
        ShowHUD("Bottom Medium Border", 800)
    } else if side = "left" && style = "thick" {
        Send("l")   ; Left border
        ShowHUD("Left Thick Border", 800)
    } else if side = "right" && style = "thick" {
        Send("r")   ; Right border
        ShowHUD("Right Thick Border", 800)
    } else {
        Send("{Escape}")  ; Cancel if not supported
        ShowHUD("Border style not supported", 1000)
    }
}

ClearBorders() {
    ; Use ribbon command to clear borders
    Send("!h")      ; Home ribbon
    Wait(Timing.RIBBON_DELAY)
    Send("b")       ; Borders dropdown
    Wait(Timing.NAV_DELAY)
    Send("n")       ; No borders
    ShowHUD("Clear Borders", 800)
}

; -----------------------------------------------------------------------------
; Row/Column operations
; -----------------------------------------------------------------------------
SetRowHeight(pts) {
    ; Use ribbon command for row height
    Send("!h")      ; Home ribbon
    Wait(Timing.RIBBON_DELAY)
    Send("o")       ; Format dropdown
    Wait(Timing.NAV_DELAY)
    Send("h")       ; Row Height
    Wait(Timing.DIALOG_DELAY)
    Send(pts)       ; Type the height
    Send("{Enter}")
    ShowHUD("Row Height: " . pts . "pt", 800)
}

SetColumnWidth(width) {
    ; Use ribbon command for column width
    Send("!h")      ; Home ribbon
    Wait(Timing.RIBBON_DELAY)
    Send("o")       ; Format dropdown
    Wait(Timing.NAV_DELAY)
    Send("w")       ; Column Width
    Wait(Timing.DIALOG_DELAY)
    Send(width)     ; Type the width
    Send("{Enter}")
    ShowHUD("Column Width: " . width, 800)
}

InsertSpacerRow() {
    ; Insert row then set height to 5pt
    Send("^{+}")        ; Insert row
    Wait(75)
    Send("{Up}")        ; Move to inserted row
    Wait(75)
    Send("!h")          ; Home ribbon
    Wait(Timing.RIBBON_DELAY)
    Send("o")           ; Format dropdown
    Wait(Timing.NAV_DELAY)
    Send("h")           ; Row Height
    Wait(100)
    Send("5{Enter}")    ; Set to 5pt
    ShowHUD("Insert Spacer Row (5pt)", 800)
}

InsertDividerColumn() {
    ; Insert column then set width to 0.5
    Send("^{Space}")    ; Select column
    Wait(75)
    Send("^{+}")        ; Insert column
    Wait(75)
    Send("{Left}")      ; Move to inserted column
    Wait(75)
    Send("!h")          ; Home ribbon
    Wait(Timing.RIBBON_DELAY)
    Send("o")           ; Format dropdown
    Wait(Timing.NAV_DELAY)
    Send("w")           ; Column Width
    Wait(100)
    Send("0.5{Enter}")  ; Set to 0.5
    ShowHUD("Insert Divider Column (0.5)", 800)
}

AutoFitColumns() {
    ; Use ribbon command for AutoFit columns
    Send("!h")      ; Home ribbon
    Wait(Timing.RIBBON_DELAY)
    Send("o")       ; Format dropdown
    Wait(Timing.NAV_DELAY)
    Send("i")       ; AutoFit Column Width
    ShowHUD("AutoFit Columns", 800)
}

AutoFitRows() {
    ; Use ribbon command for AutoFit rows
    Send("!h")      ; Home ribbon
    Wait(Timing.RIBBON_DELAY)
    Send("o")       ; Format dropdown
    Wait(Timing.NAV_DELAY)
    Send("a")       ; AutoFit Row Height
    ShowHUD("AutoFit Rows", 800)
}

; -----------------------------------------------------------------------------
; Navigation
; -----------------------------------------------------------------------------
GetDividerColumns() {
    ; Simplified - no COM needed for basic navigation
    return []
}

JumpToPrevDivider() {
    ; Use Ctrl+Left to jump to previous data boundary
    Send("^{Left}")
    ShowHUD("Jump Left", 500)
}

JumpToNextDivider() {
    ; Use Ctrl+Right to jump to next data boundary  
    Send("^{Right}")
    ShowHUD("Jump Right", 500)
}

JumpToBlockEdge(which) {
    ; Use Ctrl+Arrow keys for block edges
    if which = "first" {
        Send("^{Home}")     ; Go to beginning of worksheet
        Send("^{Right}")    ; Then to first data
        ShowHUD("First Block Edge", 500)
    } else {
        Send("^{End}")      ; Go to last used cell
        ShowHUD("Last Block Edge", 500)
    }
}

; -----------------------------------------------------------------------------
; Cleanup
; -----------------------------------------------------------------------------
TrimInPlace() {
    ; Use Find & Replace to trim spaces
    Send("^h")          ; Find & Replace dialog
    Wait(Timing.DIALOG_DELAY)
    Send("^ ")          ; Find: space at start (^ means start of cell)
    Wait(50)
    Send("{Tab}")       ; Move to Replace field
    ; Leave replace field empty
    Wait(50)
    Send("!a")          ; Replace All
    Wait(Timing.DIALOG_DELAY)
    Send("{Escape}")    ; Close dialog
    ShowHUD("Trim Spaces", 800)
}

CleanInPlace() {
    ; Use Find & Replace to clean non-printable characters
    Send("^h")          ; Find & Replace dialog
    Wait(Timing.DIALOG_DELAY)
    Send("^j")          ; Find: line break character
    Wait(50)
    Send("{Tab}")       ; Move to Replace field
    Send(" ")           ; Replace with space
    Wait(50)
    Send("!a")          ; Replace All
    Wait(Timing.DIALOG_DELAY)
    Send("{Escape}")    ; Close dialog
    ShowHUD("Clean Characters", 800)
}

CoerceToNumber() {
    ; Use Text to Columns to convert text to numbers
    Send("!d")          ; Data ribbon
    Wait(Timing.RIBBON_DELAY)
    Send("e")           ; Text to Columns
    Wait(Timing.DIALOG_DELAY)
    Send("{Enter}")     ; Accept defaults (Delimited)
    Wait(Timing.DIALOG_DELAY)
    Send("{Enter}")     ; Accept defaults (Tab delimiter)
    Wait(Timing.DIALOG_DELAY)
    Send("{Enter}")     ; Finish - this forces text-to-number conversion
    ShowHUD("Convert to Numbers", 800)
}

; -----------------------------------------------------------------------------
; Freeze Panes
; -----------------------------------------------------------------------------
FreezeAtActiveCell() {
    ; Use View ribbon to freeze panes
    Send("!w")          ; View ribbon
    Wait(Timing.RIBBON_DELAY)
    Send("f")           ; Freeze Panes dropdown
    Wait(Timing.NAV_DELAY)
    Send("f")           ; Freeze Panes at current position
    ShowHUD("Freeze Panes", 800)
}

; -----------------------------------------------------------------------------
; Row Classification
; -----------------------------------------------------------------------------
ClassifySectionHeader() {
    ; Apply section header formatting via ribbon
    Send("^{Space}")    ; Select entire row
    Wait(50)
    Send("^b")          ; Bold
    Wait(50)
    Send("!h")          ; Home ribbon
    Wait(Timing.RIBBON_DELAY)
    Send("h")           ; Fill color dropdown
    Wait(Timing.NAV_DELAY)
    Send("g")           ; Gray fill
    Wait(50)
    Send("!h")          ; Home ribbon
    Wait(Timing.RIBBON_DELAY)
    Send("b")           ; Borders dropdown
    Wait(Timing.NAV_DELAY)
    Send("b")           ; Bottom border
    ShowHUD("Section Header Format", 800)
}

ClassifySubtotal() {
    ; Apply subtotal formatting
    Send("^{Space}")    ; Select entire row
    Wait(50)
    Send("^b")          ; Bold
    Wait(50)
    Send("!h")          ; Home ribbon
    Wait(Timing.RIBBON_DELAY)
    Send("b")           ; Borders dropdown
    Wait(Timing.NAV_DELAY)
    Send("t")           ; Top border
    ShowHUD("Subtotal Format", 800)
}

ClassifyMajorTotal() {
    ; Apply major total formatting
    Send("^{Space}")    ; Select entire row
    Wait(50)
    Send("^b")          ; Bold
    Wait(50)
    Send("!h")          ; Home ribbon
    Wait(Timing.RIBBON_DELAY)
    Send("b")           ; Borders dropdown
    Wait(Timing.NAV_DELAY)
    Send("t")           ; Top border
    ShowHUD("Major Total Format", 800)
}

ClassifyGrandTotal() {
    ; Apply grand total formatting
    Send("^{Space}")    ; Select entire row
    Wait(50)
    Send("^b")          ; Bold
    Wait(50)
    Send("!h")          ; Home ribbon
    Wait(Timing.RIBBON_DELAY)
    Send("b")           ; Borders dropdown
    Wait(Timing.NAV_DELAY)
    Send("t")           ; Top border (will be double style)
    ShowHUD("Grand Total Format", 800)
}



; -----------------------------------------------------------------------------
; Macro-aware border functions
; -----------------------------------------------------------------------------
ApplyRightBorder() {
    ; Call RightBorder macro via keyboard shortcut
    Send("^+r")     ; Ctrl+Shift+R for RightBorder macro
    ShowHUD("RightBorder macro", 800)
}

ApplyBottomBorder() {
    ; Call CDS_BottomThinBorder macro via keyboard shortcut
    Send("^+b")     ; Ctrl+Shift+B for CDS_BottomThinBorder macro
    ShowHUD("CDS_BottomThinBorder macro", 800)
}

; -----------------------------------------------------------------------------
; Alignment
; -----------------------------------------------------------------------------
SetAlignment(align) {
    ; Use ribbon commands for alignment
    switch align {
        case "left":
            Send("!h")      ; Home ribbon
            Wait(Timing.RIBBON_DELAY)
            Send("al")      ; Align Left
            ShowHUD("Align Left", 800)
        case "center":
            Send("!h")      ; Home ribbon
            Wait(Timing.RIBBON_DELAY)
            Send("ac")      ; Align Center
            ShowHUD("Align Center", 800)
        case "right":
            Send("!h")      ; Home ribbon
            Wait(Timing.RIBBON_DELAY)
            Send("ar")      ; Align Right
            ShowHUD("Align Right", 800)
        default:
            ShowHUD("Alignment not supported: " . align, 1500)
    }
}

ToggleWrapText() {
    ; Use ribbon command for wrap text
    Send("!h")      ; Home ribbon
    Wait(Timing.RIBBON_DELAY)
    Send("w")       ; Wrap Text
    ShowHUD("Toggle Wrap Text", 800)
}

; Helpers for multi-step actions using ribbon commands
OutlineMedium() {
    ; Use ribbon for outline borders
    Send("!h")      ; Home ribbon
    Wait(Timing.RIBBON_DELAY)
    Send("b")       ; Borders dropdown
    Wait(Timing.NAV_DELAY)
    Send("o")       ; Outline
    ShowHUD("Outline Medium", 800)
}

InsideMedium() {
    ; Use ribbon for inside borders
    Send("!h")      ; Home ribbon
    Wait(Timing.RIBBON_DELAY)
    Send("b")       ; Borders dropdown
    Wait(Timing.NAV_DELAY)
    Send("i")       ; Inside borders
    ShowHUD("Inside Medium", 800)
}

IncreaseIndent() {
    ; Use ribbon for increase indent
    Send("!h")      ; Home ribbon
    Wait(Timing.RIBBON_DELAY)
    Send("6")       ; Increase Indent
    ShowHUD("Increase Indent", 800)
}

DecreaseIndent() {
    ; Use ribbon for decrease indent
    Send("!h")      ; Home ribbon
    Wait(Timing.RIBBON_DELAY)
    Send("5")       ; Decrease Indent
    ShowHUD("Decrease Indent", 800)
}

AddDecimalPlace() {
    ; Use ribbon for increase decimal places
    Send("!h")      ; Home ribbon
    Wait(Timing.RIBBON_DELAY)
    Send("00")      ; Increase Decimal Places
    ShowHUD("Add Decimal Place", 800)
}

RemoveDecimalPlace() {
    ; Use ribbon for decrease decimal places
    Send("!h")      ; Home ribbon
    Wait(Timing.RIBBON_DELAY)
    Send("9")       ; Decrease Decimal Places
    ShowHUD("Remove Decimal Place", 800)
}

ClearFormatsSel() {
    ; Use ribbon for clear formats
    Send("!h")      ; Home ribbon
    Wait(Timing.RIBBON_DELAY)
    Send("e")       ; Clear dropdown
    Wait(Timing.NAV_DELAY)
    Send("f")       ; Clear Formats
    ShowHUD("Clear Formats", 800)
}

ClearContentsSel() {
    ; Use Delete key for clear contents
    Send("{Delete}")
    ShowHUD("Clear Contents", 800)
}

ClearAllSel() {
    ; Use ribbon for clear all
    Send("!h")      ; Home ribbon
    Wait(Timing.RIBBON_DELAY)
    Send("e")       ; Clear dropdown
    Wait(Timing.NAV_DELAY)
    Send("a")       ; Clear All
    ShowHUD("Clear All", 800)
}

; -----------------------------------------------------------------------------
; Global CapsLock Hold Layer - Works everywhere
; -----------------------------------------------------------------------------
#HotIf GetKeyState("SC03A","P")

; HELP - Global help system showing Excel-specific shortcuts
SC03A & Space::ShowOSD("CAPSLOCK LAYER - EXCEL SHORTCUTS",
    "PASTE: v,f,t,w,s,x,l,p + Numpad+ -`n" .
    "FILTER: Numpad/ (Toggle) Numpad* (Clear)`n" .
    "DELETE: / (Ctrl=Row)`n" .
    "FORMAT: 1,2,3,4,6 + 9,a,k,5,m,d + F1-F4 + r,b,o,i,c,y,j,; + q,F5,F6,F11,F12 + Numpad.,0`n" .
    "NAV: [,],=,-,.,,g,8,h,Right,Left + Numpad8,2,4,6,7,9`n" .
    "DATA: u,F8,n,e,F7,F9 + z,Backspace,Delete`n" .
    "Note: Excel-specific hotkeys only work in Excel",
    2500, "top-center", 720)

#HotIf

; -----------------------------------------------------------------------------
; Excel-Specific CapsLock Hold Layer - All hotkeys in one block
; -----------------------------------------------------------------------------
#HotIf (IsExcel() && GetKeyState("SC03A","P"))

; PASTE operations
SC03A & v::Do(() => PasteSpecial("values"), "Paste Values")                    ; Paste Values
SC03A & f::                                                                            ; CapsLock+F or CapsLock+Ctrl+F
{
    if GetKeyState("Ctrl","P") {
        Do(() => FreezeAtActiveCell(), "Freeze Panes")
    } else {
        Do(() => PasteSpecial("formulas"), "Paste Formulas")
    }
}
SC03A & t::Do(() => PasteSpecial("formats"), "Paste Formats")                   ; Paste Formats
SC03A & w::Do(() => PasteSpecial("colwidths"), "Paste Column Widths")                 ; Paste Column Widths
SC03A & s::Do(() => PasteFormulasWithFormat(), "Paste Formulas + Format")                 ; Paste Formulas + Format
SC03A & x::Do(() => PasteSpecial("values", Map("transpose", true)), "Paste Values (Transpose)")  ; Paste Values (Transpose)
SC03A & l::Do(() => PasteLink(), "Paste Link")                               ; Paste Link
SC03A & p::Do(() => Send("^!v"), "Paste Special Dialog")                               ; Paste Special dialog

; Numpad operations
SC03A & /::
{
    if GetKeyState("Ctrl","P") {
        Do(DeleteSheetRow, "Delete Row")
    } else {
        Do(DeleteSheetColumn, "Delete Column")
    }
}
SC03A & NumpadDiv::Do(() => ToggleFilter(), "Toggle Filter")                    ; Toggle Filter (Alt+H+S+F)
SC03A & NumpadMult::Do(() => ClearFilter(), "Clear Filter")                      ; Clear Filter (Alt+H+S+C)
SC03A & NumpadAdd::Do(() => PasteOperation("add"), "Paste Add")             ; Paste Operation Add
SC03A & NumpadSub::Do(() => PasteOperation("subtract"), "Paste Subtract")        ; Paste Operation Subtract

; FORMAT - Row types
SC03A & 1::Do(() => Send("^!+1"), "Custom Font Color Macro")                       ; Custom font color macro
SC03A & 2::Do(() => ClassifySubtotal(), "Subtotal Format")                        ; Subtotal
SC03A & 3::Do(() => ClassifyMajorTotal(), "Major Total Format")                      ; Major total
SC03A & 4::Do(() => ClassifyGrandTotal(), "Grand Total Format")                      ; Grand total
SC03A & 6::Do(() => SetRowHeight(5), "Set Row Height")                           ; Row height 5pt

; FORMAT - Number formats
SC03A & 9::Do(() => SetNumberFormat("general"), "General Format")                ; General format
SC03A & a::Do(() => SetRowHeight(5), "Set Row Height")                           ; Row height 5pt
SC03A & k::Do(() => Send("^!+k"), "Custom Number Format Macro")                    ; Custom number format macro
SC03A & 5::Do(() => SetNumberFormat("percent"), "Percent Format")                ; Percent format
SC03A & m::Do(() => SetNumberFormat("month"), "Month Format")                  ; Month format
SC03A & d::Do(() => SetNumberFormat("date"), "Date Format")                   ; Date format

; ALIGNMENT
SC03A & F1::Do(() => SetAlignment("left"), "Align Left")                     ; Align Left
SC03A & F2::Do(() => SetAlignment("center"), "Align Center")                   ; Align Center
SC03A & F3::Do(() => SetAlignment("right"), "Align Right")                    ; Align Right
SC03A & F4::Do(() => ToggleWrapText(), "Toggle Wrap Text")                         ; Toggle Wrap

; BORDERS / MACROS
SC03A & r::                                                                            ; CapsLock+R or CapsLock+Ctrl+R
{
    if GetKeyState("Ctrl","P") {
        Do(AutoFitRows, "AutoFit Row Height")
    } else {
        Do(() => ApplyRightBorder(), "Apply Right Border")
    }
}
SC03A & b::Do(() => ApplyBottomBorder(), "Apply Bottom Border")                       ; Apply BottomThinBorder macro
SC03A & o::Do(() => SetBorders("outline"), "Outline Borders")                     ; Outline borders
SC03A & i::Do(() => SetBorders("inside"), "Inside Borders")                      ; Inside borders
SC03A & c::                                                                            ; CapsLock+C or CapsLock+Ctrl+C
{
    if GetKeyState("Ctrl","P") {
        Do(AutoFitColumns, "AutoFit Column Width")
    } else {
        Do(() => ClearBorders(), "Clear Borders")
    }
}
SC03A & y::Do(() => SetBorderLine("top", "double"), "Top Double Border")            ; Top double border
SC03A & j::Do(() => SetBorderLine("left", "thick"), "Left Thick Border")            ; Left thick border
SC03A & `;::Do(() => SetBorderLine("right", "thick"), "Right Thick Border")          ; Right thick border

; DIVIDERS / SIZING
SC03A & q::                                                                            ; CapsLock+Q or CapsLock+Ctrl+Q
{
    if GetKeyState("Ctrl","P") {
        Do(() => SetRowHeight(5), "Set Row Height 5pt")
    } else {
        Do(() => SetColumnWidth(0.5), "Set Column Width")
    }
}
SC03A & F5::Do(() => AutoFitColumns(), "AutoFit Columns")                         ; AutoFit Columns
SC03A & F6::Do(() => AutoFitRows(), "AutoFit Rows")                            ; AutoFit Rows
SC03A & F11::Do(IncreaseIndent, "Increase Indent")                                ; Increase indent
SC03A & F12::Do(DecreaseIndent, "Decrease Indent")                                ; Decrease indent
SC03A & NumpadDot::Do(AddDecimalPlace, "Add Decimal Place")                         ; Add decimal place
SC03A & Numpad0::Do(RemoveDecimalPlace, "Remove Decimal Place")                        ; Remove decimal place

; NAVIGATION
SC03A & [::Do(() => JumpToPrevDivider(), "Prev Divider")                       ; Prev divider
SC03A & ]::Do(() => JumpToNextDivider(), "Next Divider")                       ; Next divider
SC03A & =::Do(() => JumpToBlockEdge("first"), "First Block Edge")                  ; First block edge
SC03A & -::Do(() => JumpToBlockEdge("last"), "Last Block Edge")                   ; Last block edge
SC03A & ,::Do(() => Send("^{PgUp}"), "Previous Sheet")                           ; Previous sheet
SC03A & .::Do(() => Send("^{PgDn}"), "Next Sheet")                           ; Next sheet
SC03A & g::
{
    if GetKeyState("Ctrl","P") {
        Do(GroupAndCollapseSelection, "Group and Collapse")
    } else {
        Do(() => Send("^g"), "Go To")
    }
}
SC03A & 8::Do(() => Send("^+8"), "Current Region")                               ; Current Region
SC03A & h::Do(() => Send("^!+h"), "Custom Macro Ctrl+Alt+Shift+H")                  ; Custom macro Ctrl+Alt+Shift+H

SC03A & Right::
{
    if GetKeyState("Ctrl","P") {
        Do(() => Send("{Shift down}{Right 11}{Shift up}"), "Select Right 11")
    } else {
        Do(() => Send("{Right 12}"), "Move Right 12")
    }
}

SC03A & Left::
{
    if GetKeyState("Ctrl","P") {
        Do(() => Send("{Shift down}{Left 11}{Shift up}"), "Select Left 11")
    } else {
        Do(() => Send("{Left 12}"), "Move Left 12")
    }
}

; Numpad navigation
SC03A & Numpad8::Do(() => Send("^{Up}"), "Ctrl+Up")                       ; Ctrl+Arrow Up
SC03A & Numpad2::Do(() => Send("^{Down}"), "Ctrl+Down")                     ; Ctrl+Arrow Down
SC03A & Numpad4::Do(() => Send("^{Left}"), "Ctrl+Left")                     ; Ctrl+Arrow Left
SC03A & Numpad6::Do(() => Send("^{Right}"), "Ctrl+Right")                    ; Ctrl+Arrow Right
SC03A & Numpad7::Do(() => Send("^{Home}"), "Ctrl+Home")                     ; Ctrl+Home (A1)
SC03A & Numpad9::Do(() => Send("^{End}"), "Ctrl+End")                      ; Ctrl+End

; DATA / CLEANUP
SC03A & u::Do(() => TrimInPlace(), "Trim In Place")                             ; TRIM
SC03A & F8::Do(() => CleanInPlace(), "Clean In Place")                           ; CLEAN
SC03A & n::Do(() => CoerceToNumber(), "Convert to Number")                          ; Convert to Number
SC03A & e::Do(() => Send("!de"), "Text to Columns")                               ; Text to Columns
SC03A & F7::Do(() => Send("^+l"), "Toggle AutoFilter")                              ; Toggle AutoFilter
; SC03A & F9::Do(() => FreezeAtActiveCell(), "Freeze Panes")                     ; Freeze panes - moved to Ctrl+CapsLock+F

; CLEARS
SC03A & z::Do(ClearFormatsSel, "Clear Formats")                                 ; Clear Formats
SC03A & Backspace::Do(ClearContentsSel, "Clear Contents")                        ; Clear Contents
SC03A & Delete::Do(ClearAllSel, "Clear All")                                ; Clear All


#HotIf

; (Ctrl-layer removed: Ctrl behavior is handled inside base CapsLock combos)

; Win key safety - prevent Start menu while holding CapsLock in Excel
#HotIf (IsExcel() && GetKeyState("SC03A","P"))
LWin::Return
LWin up::Return
#HotIf




