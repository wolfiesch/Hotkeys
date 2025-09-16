; Hotkeys scaffold (AutoHotkey v2)
; Per-app key fixes for Microsoft Office apps, Chrome, Todoist, File Explorer, Cursor, VS Code, ChatGPT, and Teams
; Requires: AutoHotkey v2 — https://www.autohotkey.com/docs/v2/

#Requires AutoHotkey v2.0
#SingleInstance Force
#Warn All, MsgBox

; Ensure CapsLock starts off and stays off unless explicit logic sets it
SetCapsLockState("AlwaysOff")

; -----------------------------------------------------------------------------
; Global settings
; -----------------------------------------------------------------------------
SetTitleMatchMode(2)  ; allow partial title matching for helper detectors

; tuning
KeyDelay := 40   ; global key/macro delay (ms) - tune this once


; -----------------------------------------------------------------------------
; Utilities
; -----------------------------------------------------------------------------

JoinStrings(items, sep := "`r`n") {
    out := ""
    for i, v in items {
        out .= (i > 1 ? sep : "") v
    }
    return out
}

; Paste clipboard as plain text, preserving the original clipboard contents.
; If pickFromSelection is true, tries to copy current selection first.
PastePlainText(pickFromSelection := true) {
    clipBackup := ClipboardAll()
    try {
        if pickFromSelection {
            A_Clipboard := ""
            Send("^c")
            ClipWait(0.35)
        }
        text := A_Clipboard
        ; Normalize newlines to LF to reduce weird double-spacing in some apps
        text := StrReplace(text, "`r`n", "`n")
        A_Clipboard := text
        Send("^v")
        Sleep(KeyDelay)
    } finally {
        A_Clipboard := clipBackup
    }
}

; Copies selected Explorer item full paths, quoted, one per line.
; Falls back to copying the current folder path if no selection.
Explorer_CopyFullPathQuoted() {
    if !IsExplorer() {
        return
    }
    clipBackup := ClipboardAll()
    try {
        A_Clipboard := ""
        Send("^c")  ; copy selected items (CF_HDROP)
        if !ClipWait(0.35) {
            ; No selection or clipboard did not update: copy current folder path
            path := Explorer_GetCurrentPath()
            if path != "" {
                A_Clipboard := '"' path '"'
            }
            return
        }
        raw := A_Clipboard
        if raw = "" {
            path := Explorer_GetCurrentPath()
            if path != "" {
                A_Clipboard := '"' path '"'
            }
            return
        }
        items := []
        for _, line in StrSplit(raw, "`r`n") {
            if (line != "") {
                items.Push('"' line '"')
            }
        }
        A_Clipboard := JoinStrings(items, "`r`n")
    } finally {
        ; Intentionally keep the quoted paths in clipboard; comment next line to keep
        ; A_Clipboard := clipBackup
    }
}

; Returns current File Explorer path by focusing the address bar momentarily.
; Non-invasive: exits address bar with Escape.
Explorer_GetCurrentPath() {
    if !IsExplorer() {
        return ""
    }
    clipBackup := ClipboardAll()
    try {
        A_Clipboard := ""
        Send("!d")  ; focus address bar
        Sleep(KeyDelay)
        Send("^c")
        if !ClipWait(0.35) {
            Send("{Escape}")
            return ""
        }
        path := Trim(A_Clipboard)
        Send("{Escape}")
        return path
    } finally {
        A_Clipboard := clipBackup
    }
}

; Opens Windows Terminal (wt.exe) at the current Explorer folder. Fallbacks to PowerShell.
Explorer_OpenTerminalHere() {
    path := Explorer_GetCurrentPath()
    if path = "" {
        return
    }
    try {
        Run(Format('wt.exe -d "{1}"', path))
    } catch {
        Run(Format('powershell.exe -NoExit -Command "Set-Location -Path \"{1}\""', path))
    }
}

; Copies the current URL from the omnibox (Chrome/Edge) to clipboard.
; Leaves URL in the clipboard for subsequent pastes.
Browser_CopyCurrentUrl() {
    if !(IsChrome() || IsEdge()) {
        return
    }
    Send("^l")
    Sleep(KeyDelay)
    A_Clipboard := ""
    Send("^c")
    ClipWait(0.35)
}

; Toggles mute in Microsoft Teams.
Teams_ToggleMute() {
    Send("^+m")
}

; Paste as Values in Microsoft Excel via Ribbon: Alt, H, V, V
; Uses the current clipboard contents without copying the selection first.
Excel_PasteValues() {
    Send("!h")
    Sleep(70)
    Send("v")
    Sleep(KeyDelay)
    Send("v")
}

; Send an Excel Alt ribbon sequence like "h,o,i"
Excel_SendAlt(sequence) {
    parts := []
    for _, p in StrSplit(sequence, ",") {
        p := Trim(p)
        if (p != "")
            parts.Push(p)
    }
    if parts.Length = 0 {
        return
    }
    Send("!" . parts[1])
    Sleep(KeyDelay)
    if parts.Length > 1 {
        for idx, key in parts {
            if (idx = 1)
                continue
            Send(key)
            Sleep(KeyDelay)
        }
    }
}

; AutoFit selected columns (robust COM first, then ribbon keys Alt,H,O,I)
Excel_AutoFitColumns() {
    try {
        xl := ComObjActive("Excel.Application")
        sel := xl.Selection
        if sel {
            sel.Columns.AutoFit()
            return
        }
    } catch as err {
        ; fall back to ribbon keys
    }
    Send("!h")
    Sleep(KeyDelay)
    Send("o")
    Sleep(KeyDelay)
    Send("i")
}

; Excel: Paste Special → Formulas and number formats, Skip blanks
Excel_PasteFormulasAndNumberFormats_SkipBlanks() {
    try {
        xl := ComObjActive("Excel.Application")
        sel := xl.Selection
        if !sel {
            return
        }
        xlPasteFormulasAndNumberFormats := 11    ; XlPasteType
        xlPasteSpecialOperationNone := -4142     ; XlPasteSpecialOperation
        sel.PasteSpecial(xlPasteFormulasAndNumberFormats, xlPasteSpecialOperationNone, true, false)
    } catch as err {
        ; Fallback to dialog if COM fails
        Send("^!v")
        Sleep(KeyDelay)
        ; Best-effort: sequence that often selects 'Formulas and number formats' and toggles Skip blanks
        ; Depending on locale, adjust accelerators; user can refine later
        Send("f")   ; Formulas (fallback)
        Sleep(KeyDelay)
        Send("b")   ; Toggle Skip blanks (fallback)
        Sleep(KeyDelay)
        Send("{Enter}")
    }
}

; -----------------------------------------------------------------------------
; CapsLock dual-role and combos
; -----------------------------------------------------------------------------
; Tap: toggles CapsLock. Hold + key: triggers a unique per-key combo (placeholders provided).
CapsLayer_DEBUG := false

gCapsLayerUsed := false

CapsLock::{
    global gCapsLayerUsed
    gCapsLayerUsed := false
}

CapsLock up::{
    global gCapsLayerUsed
    ; Do not toggle CapsLock
    gCapsLayerUsed := false
}

CapsLayer_Dispatch(key) {
    global gCapsLayerUsed, CapsLayer_DEBUG
    gCapsLayerUsed := true
    try {
        switch key {
            case "t":
                Debug_AltTab()
            case "a","s","d","f","h","j","k","l","u","i","o":
                CapsLayer_Placeholder("Caps+" . key)
            case "Left","Right","Up","Down","Space","Enter","Tab","Backspace":
                CapsLayer_Placeholder("Caps+" . key)
            default:
                if CapsLayer_DEBUG {
                    ToolTip "Caps+" . key, , , 1
                    SetTimer(() => ToolTip(), -800)
                }
        }
    } catch as err {
        if CapsLayer_DEBUG {
            ToolTip "Caps+" . key . " error: " . err.Message, , , 1
            SetTimer(() => ToolTip(), -800)
        }
    }
}

CapsLayer_Placeholder(label) {
    ToolTip(label " (placeholder)", , , 1)
    SetTimer(() => ToolTip(), -600)
}

Debug_AltTab() {
    SendEvent("{Alt down}{Tab}{Alt up}")
}

; Excel-specific CapsLock subgroups
; Ctrl+CapsLock → Paste mode; Alt+CapsLock → Format mode
Debug_Excel_PasteMode(action) {
    ToolTip("Excel Paste Mode: " . action, , , 1)
    SetTimer(() => ToolTip(), -700)
}

Debug_Excel_FormatMode(action) {
    ToolTip("Excel Format Mode: " . action, , , 1)
    SetTimer(() => ToolTip(), -700)
}

#HotIf IsExcel() && GetKeyState("Ctrl", "P") && GetKeyState("CapsLock", "P")
; Paste mode (Ctrl+CapsLock held)
CapsLock & e::{
    global gCapsLayerUsed
    gCapsLayerUsed := true
    Debug_Excel_PasteMode("Formulas+NumberFormats, Skip blanks")
    ; Later: Excel_PasteFormulasAndNumberFormats_SkipBlanks()
}
; Ctrl+CapsLock+C → Alt,H,F,P
CapsLock & c::{
	global gCapsLayerUsed
	gCapsLayerUsed := true
	Send("!h")
	Sleep(KeyDelay)
	Send("f")
	Sleep(KeyDelay)
	Send("p")
}
#HotIf

#HotIf IsExcel() && GetKeyState("Alt", "P") && GetKeyState("CapsLock", "P")
; Format mode (Alt+CapsLock held)
CapsLock & r::{
    global gCapsLayerUsed
    gCapsLayerUsed := true
    Debug_Excel_FormatMode("Right border: dashed style #2")
    ; Later: implement dashed border style #2 and apply right border to selection via COM
}
#HotIf

; Static Caps layer bindings using #HotIf for reliability
#HotIf GetKeyState("CapsLock", "P")
; Letters
CapsLock & a::CapsLayer_Dispatch("a")
CapsLock & s::CapsLayer_Dispatch("s")
CapsLock & d::CapsLayer_Dispatch("d")
CapsLock & f::CapsLayer_Dispatch("f")
CapsLock & h::CapsLayer_Dispatch("h")
CapsLock & j::CapsLayer_Dispatch("j")
CapsLock & k::CapsLayer_Dispatch("k")
CapsLock & l::CapsLayer_Dispatch("l")
CapsLock & u::CapsLayer_Dispatch("u")
CapsLock & i::CapsLayer_Dispatch("i")
CapsLock & o::CapsLayer_Dispatch("o")
CapsLock & t::{
    global gCapsLayerUsed
    gCapsLayerUsed := true
    SendEvent("{Alt down}{Tab}{Alt up}")
}
; Digits
CapsLock & 1::CapsLayer_Dispatch("1")
CapsLock & 2::CapsLayer_Dispatch("2")
CapsLock & 3::CapsLayer_Dispatch("3")
CapsLock & 4::CapsLayer_Dispatch("4")
CapsLock & 5::CapsLayer_Dispatch("5")
CapsLock & 6::CapsLayer_Dispatch("6")
CapsLock & 7::CapsLayer_Dispatch("7")
CapsLock & 8::CapsLayer_Dispatch("8")
CapsLock & 9::CapsLayer_Dispatch("9")
CapsLock & 0::CapsLayer_Dispatch("0")
; Navigation / control
CapsLock & Left::CapsLayer_Dispatch("Left")
CapsLock & Right::CapsLayer_Dispatch("Right")
CapsLock & Up::CapsLayer_Dispatch("Up")
CapsLock & Down::CapsLayer_Dispatch("Down")
CapsLock & Space::CapsLayer_Dispatch("Space")
CapsLock & Enter::CapsLayer_Dispatch("Enter")
CapsLock & Tab::{
    global gCapsLayerUsed
    gCapsLayerUsed := true
    SendEvent("{Alt down}{Tab}{Alt up}")
}
CapsLock & Backspace::CapsLayer_Dispatch("Backspace")
; Year jumps
CapsLock & [::Send("{Left 12}")
CapsLock & ]::Send("{Right 12}")
#HotIf

; -----------------------------------------------------------------------------
; App detectors (for use with #HotIf)
; -----------------------------------------------------------------------------
IsExe(exeName) {
    return WinActive("ahk_exe " exeName)
}

IsWord() {
    return IsExe("WINWORD.EXE")
}

IsExcel() {
    return IsExe("EXCEL.EXE")
}

IsPowerPoint() {
    return IsExe("POWERPNT.EXE")
}

IsOutlook() {
    return IsExe("OUTLOOK.EXE")
}

IsOneNote() {
    ; Desktop OneNote
    if IsExe("ONENOTE.EXE")
        return true
    ; Store/UWP OneNote runs in ApplicationFrameHost.exe; match by title contains OneNote
    return WinActive("ahk_class ApplicationFrameWindow") && RegExMatch(WinGetTitle("A"), "i)OneNote")
}

IsChrome() {
    return IsExe("chrome.exe")
}

IsEdge() {
    return IsExe("msedge.exe")
}

IsTodoist() {
    if !(IsChrome() || IsEdge())
        return false
    return RegExMatch(WinGetTitle("A"), "i)Todoist")
}

IsExplorer() {
    return WinActive("ahk_class CabinetWClass") || WinActive("ahk_class ExploreWClass")
}

IsVSCode() {
    return IsExe("Code.exe")
}

IsCursor() {
    return IsExe("Cursor.exe")
}

IsChatGPT() {
    if !(IsChrome() || IsEdge())
        return false
    ; Titles often contain ChatGPT, sometimes with suffixes
    return RegExMatch(WinGetTitle("A"), "i)\b(ChatGPT|OpenAI)\b")
}

IsTeams() {
    return IsExe("ms-teams.exe") || IsExe("Teams.exe")
}

; -----------------------------------------------------------------------------
; Global hotkeys (active everywhere EXCEPT where app-specific overrides exist)
; -----------------------------------------------------------------------------
#HotIf !IsExcel()  ; Excel gets its own Ctrl+Shift+V handler
^+v::{
    PastePlainText(true)
}
#HotIf

; Copy current URL quickly in browsers
#HotIf (IsChrome() || IsEdge())
^+y::{
    Browser_CopyCurrentUrl()
}
#HotIf

; -----------------------------------------------------------------------------
; Microsoft Word
; -----------------------------------------------------------------------------
#HotIf IsWord()
; Add Word-specific remaps below
; Example: keep plain-text paste in Word as well (in addition to global)
^+v::{
    PastePlainText(true)
}
#HotIf

; -----------------------------------------------------------------------------
; Microsoft Excel
; -----------------------------------------------------------------------------
#HotIf IsExcel()
; Ctrl+Shift+V → Paste Values
^+v::{
    Excel_PasteValues()
}
; Horizontal scroll fix: Shift+WheelUp/Down → WheelLeft/Right
+WheelUp::Send("{WheelLeft}")
+WheelDown::Send("{WheelRight}")
#HotIf

; Excel: CapsLock combos
#HotIf IsExcel() && GetKeyState("CapsLock", "P")
; CapsLock+V → Paste Values (use COM if available; else keys)
CapsLock & v::{
	Excel_PasteValues()
}
; CapsLock+C → Ctrl+C (copy)
CapsLock & c::{
	Send("^c")
}
; CapsLock+W → Wrap Text (Alt, H, W)
CapsLock & w::{
	Send("!h")
	Sleep(KeyDelay)
	Send("w")
}
; CapsLock+Numpad2 → Ctrl+B, then Ctrl+Shift+B (macro)
CapsLock & Numpad2::{
	Send("^b")
	Sleep(KeyDelay)
	Send("^+b")
}
; CapsLock+F → AutoFit Column Width (COM or Alt,H,O,I fallback)
CapsLock & f::{
	Excel_PasteFormulasAndNumberFormats_SkipBlanks()
}
#HotIf

; -----------------------------------------------------------------------------
; Microsoft PowerPoint
; -----------------------------------------------------------------------------
#HotIf IsPowerPoint()
; Example placeholder: add your PowerPoint-specific hotkeys here
#HotIf

; -----------------------------------------------------------------------------
; Microsoft Outlook
; -----------------------------------------------------------------------------
#HotIf IsOutlook()
; Example placeholder: add your Outlook-specific hotkeys here
; Keep plain-text paste
^+v::{
    PastePlainText(true)
}
#HotIf

; -----------------------------------------------------------------------------
; Microsoft OneNote
; -----------------------------------------------------------------------------
#HotIf IsOneNote()
^+v::{
    PastePlainText(true)
}
#HotIf

; -----------------------------------------------------------------------------
; Google Chrome
; -----------------------------------------------------------------------------
#HotIf IsChrome()
; Example placeholder: add your Chrome-specific hotkeys here
#HotIf

; -----------------------------------------------------------------------------
; Todoist (in Chrome or Edge tab)
; -----------------------------------------------------------------------------
#HotIf IsTodoist()
; Example: quick add task (focus add input), customize as desired
; Send("a")  ; Uncomment and adapt to your Todoist workflow
#HotIf

; -----------------------------------------------------------------------------
; File Explorer
; -----------------------------------------------------------------------------
#HotIf IsExplorer()
; Copy selected items' full paths (quoted)
^+c::{
    Explorer_CopyFullPathQuoted()
}
; Open Windows Terminal here
!t::{
    Explorer_OpenTerminalHere()
}
#HotIf

; -----------------------------------------------------------------------------
; VS Code
; -----------------------------------------------------------------------------
#HotIf IsVSCode()
; Example placeholder: add your VS Code-specific hotkeys here
#HotIf

; -----------------------------------------------------------------------------
; Cursor
; -----------------------------------------------------------------------------
#HotIf IsCursor()
; Example placeholder: add your Cursor-specific hotkeys here
#HotIf

; -----------------------------------------------------------------------------
; ChatGPT (tab in Chrome/Edge)
; -----------------------------------------------------------------------------
#HotIf IsChatGPT()
; Example placeholder: add your ChatGPT-specific hotkeys here
#HotIf

; -----------------------------------------------------------------------------
; Microsoft Teams (classic or new)
; -----------------------------------------------------------------------------
#HotIf IsTeams()
; Alt+M → mute/unmute
!m::{
    Teams_ToggleMute()
}
#HotIf


