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

; Global variables for sticky tab navigation
global TabSwitchStates := Map()

; Initialize Status Overlay and Hotkey Tracking
global StatusOverlay := {
    gui: 0,
    timer: 0,
    visible: true,
    hovering: false,
    lastHotkey: "",
    lastDescription: "",
    lastTimestamp: 0
}

; Track recent hotkey history (up to 5 items)
global HotkeyHistory := []
global HotkeyDescriptions := Map()

; Shared layer hotkey metadata used for registration and overlay descriptions
global LayerHotkeyConfig := [
    {
        leader: "SC03A",
        modifiers: ["Ctrl"],
        key: "f",
        actionFn: FreezeAtActiveCell,
        description: "Freeze Panes",
        contextFn: IsExcel
    }
]

; -----------------------------------------------------------------------------
; Hotkey Tracking and Description System - Must be defined before use
; -----------------------------------------------------------------------------

InitializeHotkeyDescriptions() {
    ; Initialize the map with common hotkey descriptions
    global HotkeyDescriptions

    ; PASTE operations
    HotkeyDescriptions["CapsLock+V"] := "Paste Values"
    HotkeyDescriptions["CapsLock+F"] := "Paste Formulas"
    HotkeyDescriptions["CapsLock+T"] := "Paste Formats"
    HotkeyDescriptions["CapsLock+W"] := "Column Widths"
    HotkeyDescriptions["CapsLock+S"] := "Formulas + Format"
    HotkeyDescriptions["CapsLock+X"] := "Values Transpose"
    HotkeyDescriptions["CapsLock+L"] := "Paste Link"
    HotkeyDescriptions["CapsLock+P"] := "Paste Special Dialog"

    ; FORMAT operations
    HotkeyDescriptions["CapsLock+1"] := "Font Color"
    HotkeyDescriptions["CapsLock+2"] := "Subtotal Format"
    HotkeyDescriptions["CapsLock+3"] := "Major Total Format"
    HotkeyDescriptions["CapsLock+4"] := "Grand Total Format"
    HotkeyDescriptions["CapsLock+5"] := "Percent Format"
    HotkeyDescriptions["CapsLock+6"] := "Row Height 5pt"
    HotkeyDescriptions["CapsLock+9"] := "General Format"
    HotkeyDescriptions["CapsLock+K"] := "Number Format"
    HotkeyDescriptions["CapsLock+M"] := "Month Format"
    HotkeyDescriptions["CapsLock+D"] := "Date Format"

    ; ALIGNMENT
    HotkeyDescriptions["CapsLock+F1"] := "Align Left"
    HotkeyDescriptions["CapsLock+F2"] := "Align Center"
    HotkeyDescriptions["CapsLock+F3"] := "Align Right"

    ; NAVIGATION
    HotkeyDescriptions["CapsLock+N"] := "Name Manager"
    HotkeyDescriptions["CapsLock+G"] := "Go To Dialog"
    HotkeyDescriptions["CapsLock+;"] := "Keep On Top"
    HotkeyDescriptions["CapsLock+Enter"] := "Terminal Here"
    HotkeyDescriptions["CapsLock+Down"] := "Next Tab"
    HotkeyDescriptions["CapsLock+Up"] := "Previous Tab"
    HotkeyDescriptions["CapsLock+Ctrl+Right"] := "Next Sheet"
    HotkeyDescriptions["CapsLock+Ctrl+Left"] := "Previous Sheet"

    ; Other operations
    HotkeyDescriptions["CapsLock+/"] := "Delete Column"
    HotkeyDescriptions["CapsLock+Ctrl+/"] := "Delete Row"
    HotkeyDescriptions["CapsLock+H"] := "Highlight Cell"
    HotkeyDescriptions["CapsLock+O"] := "Open File"
    HotkeyDescriptions["Ctrl+Alt+H"] := "Toggle Overlay"
}

TrackHotkey(key, description := "") {
    global StatusOverlay, HotkeyHistory, HotkeyDescriptions

    ; Get description from map if not provided
    if (description == "" && HotkeyDescriptions.Has(key)) {
        description := HotkeyDescriptions[key]
    }

    ; Update last hotkey info
    StatusOverlay.lastHotkey := key
    StatusOverlay.lastDescription := description
    StatusOverlay.lastTimestamp := A_TickCount

    ; Add to history (keep last 3)
    historyItem := {
        key: key,
        description: description,
        timestamp: A_TickCount
    }

    HotkeyHistory.InsertAt(1, historyItem)
    if (HotkeyHistory.Length > 3) {
        HotkeyHistory.Pop()
    }
}

; -----------------------------------------------------------------------------
; Layer Hotkey Registration Helpers
; -----------------------------------------------------------------------------

RegisterLayerHotkeys(config) {
    ; Dynamically bind layer-aware hotkeys defined in the shared metadata.
    ; The helper also keeps HotkeyDescriptions synchronized so the overlay
    ; reflects every registered shortcut without redundant manual entries.
    global HotkeyDescriptions

    if !IsObject(config) {
        return
    }

    for entry in config {
        ; Build a context predicate that respects the application guard,
        ; leader state, and any additional modifier requirements.
        contextPredicate := (*) => (
            entry.contextFn.Call()
            && GetKeyState(entry.leader, "P")
            && AreLayerModifiersActive(entry.modifiers)
        )

        hotkeyLabel := BuildLayerHotkeyLabel(entry)

        ; Wrap the registered action in Do() so tracking and error handling
        ; remain consistent with the existing static hotkey definitions.
        callback := (*) => Do(entry.actionFn, entry.description, hotkeyLabel)

        HotIf(contextPredicate)
        Hotkey(entry.leader . " & " . entry.key, callback, "On")
        HotkeyDescriptions[hotkeyLabel] := entry.description
    }

    ; Reset HotIf to avoid leaking the context to unrelated hotkeys.
    HotIf()
}

AreLayerModifiersActive(requiredModifiers) {
    ; Ensure every required modifier is pressed while allowing the helper to
    ; succeed when no modifiers are specified in the metadata.
    if !IsObject(requiredModifiers) || requiredModifiers.Length = 0 {
        return true
    }

    for modifier in requiredModifiers {
        if !GetKeyState(modifier, "P") {
            return false
        }
    }

    return true
}

BuildLayerHotkeyLabel(entry) {
    ; Produce a human-readable label (e.g., CapsLock+Ctrl+F) that mirrors the
    ; existing convention in HotkeyDescriptions for overlay consumption.
    leaderName := GetLayerLeaderDisplayName(entry.leader)
    labelParts := [leaderName]

    if IsObject(entry.modifiers) {
        for modifier in entry.modifiers {
            labelParts.Push(modifier)
        }
    }

    labelParts.Push(StrUpper(entry.key))

    label := ""
    for index, part in labelParts {
        label .= (index > 1 ? "+" : "") . part
    }

    return label
}

GetLayerLeaderDisplayName(leader) {
    ; Map scan codes to descriptive names so overlay strings stay friendly.
    leaderNames := Map("SC03A", "CapsLock")

    if leaderNames.Has(leader) {
        return leaderNames[leader]
    }

    return leader
}

; Initialize the hotkey descriptions now that function is defined
InitializeHotkeyDescriptions()
RegisterLayerHotkeys(LayerHotkeyConfig)

; Create the overlay GUI
CreateStatusOverlay()
SetTimer(UpdateStatusOverlay, 150)  ; Update every 150ms for responsive feedback

; Treat CapsLock as a global modifier
SC03A::Return

; Release any held modifiers when CapsLock is released
SC03A up::ReleaseTabSwitchModifiers()

; Global overlay toggle hotkey (works without CapsLock)
^!h::{
    TrackHotkey("Ctrl+Alt+H", "Toggle Overlay")
    ToggleStatusOverlay()
}

; Set consistent key timing for better Excel compatibility
SetKeyDelay(50, 50)  ; 50ms press duration, 50ms release delay

; Multi-application support: Excel (CapsLock layers) + PowerPoint (Ctrl+Alt+Shift+S)

; Active when Excel or PowerPoint is focused

; -----------------------------------------------------------------------------
; PowerPoint CapsLock Integration
; -----------------------------------------------------------------------------
; Only active when PowerPoint is focused
#HotIf WinActive("ahk_exe POWERPNT.EXE")

; Ctrl+Alt+Shift+S -> Send specific sequence: Tab, Tab, 70, Enter, Tab x4, 4.57, Tab x2, 21.47
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

; Ctrl+Alt+Shift+F -> Format Object Pane Macro (Alt+4, wait, Ctrl+A, wait, "70", Enter)
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

; CapsLock + Down Arrow -> Next tab (Ctrl held, then Tab)
SC03A & Down::HandleCapsTabSwitch(Func("IsChrome"), "Chrome", "next")

; CapsLock + Up Arrow -> Previous tab (Ctrl+Shift held, then Tab)
SC03A & Up::HandleCapsTabSwitch(Func("IsChrome"), "Chrome", "previous")

#HotIf

; -----------------------------------------------------------------------------
; Cursor CapsLock Integration
; -----------------------------------------------------------------------------
; Only active when Cursor is focused and CapsLock is held
#HotIf (IsCursor() && GetKeyState("SC03A","P"))

; CapsLock + M -> Insert commit message template and press Enter
SC03A & m::
{
    TrackHotkey("CapsLock+M", "Generate Commit Msg")
    Send("^k")  ; Ctrl+K to open AI command palette
    Sleep(200)  ; Wait for AI command palette to open
    Send("Generate a commit message summarizing recent changes")
    Send("{Enter}")
}

; CapsLock + Down Arrow -> Next tab (Ctrl held, then Tab)
SC03A & Down::HandleCapsTabSwitch(Func("IsCursor"), "Cursor", "next")

; CapsLock + Up Arrow -> Previous tab (Ctrl+Shift held, then Tab)
SC03A & Up::HandleCapsTabSwitch(Func("IsCursor"), "Cursor", "previous")

#HotIf

; -----------------------------------------------------------------------------
; File Explorer CapsLock Integration
; -----------------------------------------------------------------------------
; Only active when File Explorer is focused and CapsLock is held
#HotIf (IsFileExplorer() && GetKeyState("SC03A","P"))

; CapsLock + Enter -> Open selected folder in Windows Terminal
SC03A & Enter::
{
    TrackHotkey("CapsLock+Enter", "Terminal Here")
    try {
        ; Get the selected items in File Explorer
        for window in ComObject("Shell.Application").Windows {
            if (window.hwnd = WinExist("A")) {
                ; Check if any items are selected
                selectedItems := window.Document.SelectedItems()

                ; If no items selected, get the current folder
                if (selectedItems.Count = 0) {
                    currentPath := window.Document.Folder.Self.Path
                    if (DirExist(currentPath)) {
                        ; Try Windows Terminal first, fall back to cmd
                        try {
                            Run('wt.exe -d "' . currentPath . '"')
                            ShowHUD("Opening terminal in current folder", 1500)
                        } catch {
                            Run('cmd.exe /k cd /d "' . currentPath . '"')
                            ShowHUD("Opening CMD in current folder", 1500)
                        }
                    }
                    return
                }

                ; Process the first selected item
                for item in selectedItems {
                    itemPath := item.Path

                    ; Check if it's a folder (not a file)
                    if (DirExist(itemPath)) {
                        ; Try Windows Terminal first, fall back to cmd
                        try {
                            Run('wt.exe -d "' . itemPath . '"')
                            ShowHUD("Opening terminal in: " . item.Name, 1500)
                        } catch {
                            Run('cmd.exe /k cd /d "' . itemPath . '"')
                            ShowHUD("Opening CMD in: " . item.Name, 1500)
                        }
                    } else {
                        ShowHUD("Selected item is not a folder", 1500)
                    }
                    break  ; Only process first selected item
                }
                break
            }
        }
    } catch as e {
        ShowHUD("Error: " . e.Message, 2000)
    }
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
IsFileExplorer() => WinActive("ahk_class CabinetWClass") || WinActive("ahk_class ExploreWClass")
IsVSCode() => WinActive("ahk_exe Code.exe")

HandleCapsTabSwitch(appCheckFn, stateKey, direction) {
    global TabSwitchStates

    if (!appCheckFn.Call()) {
        return
    }

    if (!TabSwitchStates.Has(stateKey)) {
        TabSwitchStates[stateKey] := { ctrlHeld: false, shiftHeld: false }
    }

    state := TabSwitchStates[stateKey]
    isNext := (direction = "next")
    TrackHotkey(isNext ? "CapsLock+Down" : "CapsLock+Up", isNext ? "Next Tab" : "Previous Tab")

    if (!state.ctrlHeld) {
        Send("{Ctrl Down}")
        state.ctrlHeld := true
    }

    if (isNext) {
        if (state.shiftHeld) {
            Send("{Shift Up}")
            state.shiftHeld := false
        }
    } else {
        if (!state.shiftHeld) {
            Send("{Shift Down}")
            state.shiftHeld := true
        }
    }

    Send("{Tab}")
}

ReleaseTabSwitchModifiers() {
    global TabSwitchStates

    for stateKey, state in TabSwitchStates {
        if (state.shiftHeld) {
            Send("{Shift Up}")
            state.shiftHeld := false
        }
        if (state.ctrlHeld) {
            Send("{Ctrl Up}")
            state.ctrlHeld := false
        }
        TabSwitchStates[stateKey] := state
    }
}

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
; Persistent Layer Status Overlay - Enhanced
; -----------------------------------------------------------------------------

CreateStatusOverlay() {
    try StatusOverlay.gui.Destroy()

    StatusOverlay.gui := Gui("+AlwaysOnTop -Caption +ToolWindow +E0x20 -DPIScale", "LayerStatus")
    StatusOverlay.gui.BackColor := "0x1a1a1a"
    StatusOverlay.gui.MarginX := 12, StatusOverlay.gui.MarginY := 10

    ; === LAYER SECTION ===
    StatusOverlay.gui.SetFont("s11 Bold", "Segoe UI")
    StatusOverlay.layerLabel := StatusOverlay.gui.AddText("xm ym c0x888888 w240", "CURRENT LAYER")

    StatusOverlay.gui.SetFont("s14 Bold", "Segoe UI")
    StatusOverlay.layerText := StatusOverlay.gui.AddText("xm y+2 cWhite w240", "🎯 Unknown")

    ; Divider line
    StatusOverlay.gui.SetFont("s8", "Segoe UI")
    StatusOverlay.gui.AddText("xm y+8 w240 h1 0x10")  ; Horizontal line

    ; === CAPSLOCK STATUS ===
    StatusOverlay.gui.SetFont("s11 Bold", "Segoe UI")
    StatusOverlay.capslockLabel := StatusOverlay.gui.AddText("xm y+8 c0x888888 w240", "CAPSLOCK STATUS")

    StatusOverlay.gui.SetFont("s13 Bold", "Segoe UI")
    StatusOverlay.capslockText := StatusOverlay.gui.AddText("xm y+2 c0x666666 w240", "⚫ OFF")

    ; Divider line
    StatusOverlay.gui.AddText("xm y+8 w240 h1 0x10")

    ; === LAST HOTKEY SECTION ===
    StatusOverlay.gui.SetFont("s11 Bold", "Segoe UI")
    StatusOverlay.lastHotkeyLabel := StatusOverlay.gui.AddText("xm y+8 c0x888888 w240", "LAST HOTKEY")

    StatusOverlay.gui.SetFont("s12 Bold", "Segoe UI")
    StatusOverlay.lastHotkeyText := StatusOverlay.gui.AddText("xm y+2 c0x00FFFF w240", "None")

    StatusOverlay.gui.SetFont("s10", "Segoe UI")
    StatusOverlay.lastDescText := StatusOverlay.gui.AddText("xm y+2 cWhite w240", "")

    ; Divider line
    StatusOverlay.gui.AddText("xm y+8 w240 h1 0x10")

    ; === RECENT HISTORY ===
    StatusOverlay.gui.SetFont("s11 Bold", "Segoe UI")
    StatusOverlay.historyLabel := StatusOverlay.gui.AddText("xm y+8 c0x888888 w240", "RECENT HISTORY")

    StatusOverlay.gui.SetFont("s9", "Segoe UI")
    StatusOverlay.history1 := StatusOverlay.gui.AddText("xm y+4 c0x999999 w240", "")
    StatusOverlay.history2 := StatusOverlay.gui.AddText("xm y+2 c0x888888 w240", "")
    StatusOverlay.history3 := StatusOverlay.gui.AddText("xm y+2 c0x777777 w240", "")

    ; Set transparency and rounded corners effect
    WinSetTransparent(230, StatusOverlay.gui.Hwnd)

    ; Try to set rounded corners (Windows 11)
    try DllCall("dwmapi\DwmSetWindowAttribute", "ptr", StatusOverlay.gui.Hwnd, "int", 33, "int*", 2, "int", 4)

    ; Position in top-right corner with padding
    StatusOverlay.gui.Show("NoActivate x" . (A_ScreenWidth - 280) . " y30")

    return StatusOverlay.gui
}

GetCurrentLayer() {
    if IsExcel()
        return "Excel"
    else if IsChrome()
        return "Chrome"
    else if IsVSCode()
        return "VS Code"
    else if IsCursor()
        return "Cursor"
    else if IsFileExplorer()
        return "Explorer"
    else if WinActive("ahk_exe POWERPNT.EXE")
        return "PowerPoint"
    else
        return "Other"
}

GetLayerIcon(layer) {
    switch layer {
        case "Excel":      return "📊"
        case "Chrome":     return "🌐"
        case "VS Code":    return "💻"
        case "Cursor":     return "✨"
        case "Explorer":   return "📁"
        case "PowerPoint": return "📽️"
        default:           return "🔲"
    }
}

GetLayerColor(layer) {
    switch layer {
        case "Excel":      return "0x00FF00"  ; Green
        case "Chrome":     return "0x4285F4"  ; Blue
        case "VS Code":    return "0x007ACC"  ; VS Code Blue
        case "Cursor":     return "0xFF69B4"  ; Pink
        case "Explorer":   return "0xFFD700"  ; Gold
        case "PowerPoint": return "0xFF4500"  ; OrangeRed
        default:           return "0xFFFFFF"  ; White
    }
}

UpdateStatusOverlay() {
    global HotkeyHistory

    if (!StatusOverlay.visible || !StatusOverlay.gui)
        return

    ; Check if mouse is hovering over overlay
    mouseHovering := IsMouseOverOverlay()

    ; Handle hover state changes
    if (mouseHovering && !StatusOverlay.hovering) {
        ; Mouse just started hovering - fade out
        StatusOverlay.hovering := true
        WinSetTransparent(50, StatusOverlay.gui.Hwnd)  ; Very transparent when hovering
    } else if (!mouseHovering && StatusOverlay.hovering) {
        ; Mouse stopped hovering - restore normal transparency
        StatusOverlay.hovering := false
        WinSetTransparent(230, StatusOverlay.gui.Hwnd)  ; Normal transparency
    }

    currentLayer := GetCurrentLayer()
    capslockHeld := GetKeyState("SC03A", "P")

    ; Update layer text with icon and color
    layerIcon := GetLayerIcon(currentLayer)
    layerColor := GetLayerColor(currentLayer)
    StatusOverlay.layerText.Text := layerIcon . " " . currentLayer
    StatusOverlay.gui.SetFont("s14 Bold", "Segoe UI")
    StatusOverlay.layerText.SetFont("c" . layerColor)

    ; Update CapsLock status with color coding
    if (capslockHeld) {
        StatusOverlay.capslockText.Text := "🔴 ON - ACTIVE"
        StatusOverlay.gui.SetFont("s13 Bold", "Segoe UI")
        StatusOverlay.capslockText.SetFont("c0xFF4444")  ; Bright Red
    } else {
        StatusOverlay.capslockText.Text := "⚫ OFF"
        StatusOverlay.gui.SetFont("s13 Bold", "Segoe UI")
        StatusOverlay.capslockText.SetFont("c0x666666")  ; Gray
    }

    ; Update last hotkey info - clear if older than 30 seconds
    if (StatusOverlay.lastHotkey != "" && A_TickCount - StatusOverlay.lastTimestamp > 30000) {
        StatusOverlay.lastHotkey := ""
        StatusOverlay.lastDescription := ""
    }

    ; Update last hotkey display
    if (StatusOverlay.lastHotkey != "") {
        StatusOverlay.lastHotkeyText.Text := StatusOverlay.lastHotkey
        StatusOverlay.lastDescText.Text := "→ " . StatusOverlay.lastDescription
    } else {
        StatusOverlay.lastHotkeyText.Text := "None"
        StatusOverlay.lastDescText.Text := ""
    }

    ; Update history display
    Loop 3 {
        if (A_Index <= HotkeyHistory.Length) {
            item := HotkeyHistory[A_Index]
            ; Format: "Key → Description"
            historyText := item.key . " → " . item.description

            ; Get the control reference
            if (A_Index == 1)
                StatusOverlay.history1.Text := historyText
            else if (A_Index == 2)
                StatusOverlay.history2.Text := historyText
            else if (A_Index == 3)
                StatusOverlay.history3.Text := historyText
        } else {
            ; Clear empty history slots
            if (A_Index == 1)
                StatusOverlay.history1.Text := ""
            else if (A_Index == 2)
                StatusOverlay.history2.Text := ""
            else if (A_Index == 3)
                StatusOverlay.history3.Text := ""
        }
    }
}

IsMouseOverOverlay() {
    if (!StatusOverlay.gui || !StatusOverlay.visible)
        return false

    ; Get mouse position
    MouseGetPos(&mouseX, &mouseY)

    ; Get overlay position and size
    try {
        WinGetPos(&overlayX, &overlayY, &overlayW, &overlayH, StatusOverlay.gui.Hwnd)

        ; Check if mouse is within overlay boundaries (with small margin)
        margin := 5
        return (mouseX >= overlayX - margin && mouseX <= overlayX + overlayW + margin &&
                mouseY >= overlayY - margin && mouseY <= overlayY + overlayH + margin)
    } catch {
        return false
    }
}

ToggleStatusOverlay() {
    if (StatusOverlay.visible) {
        StatusOverlay.gui.Hide()
        StatusOverlay.visible := false
        ShowHUD("Status Overlay: Hidden", 1000)
    } else {
        StatusOverlay.gui.Show("NoActivate")
        StatusOverlay.visible := true
        ShowHUD("Status Overlay: Visible", 1000)
    }
}

; -----------------------------------------------------------------------------
; Layer System - Now using held-down keys
; -----------------------------------------------------------------------------
Do(fn, operation := "Unknown Operation", hotkeyKey := "") {
    try {
        ; Track the hotkey if key is provided
        if (hotkeyKey != "") {
            TrackHotkey(hotkeyKey, operation)
        }
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
    ; Alt + E -> S -> R -> B -> Enter
    ; Opens Paste Special dialog, selects Formulas, enables Skip blanks
    Send("!es")     ; Alt+E+S to open Paste Special dialog
    Wait(Timing.DIALOG_DELAY)
    Send("r")       ; R for Formulas
    Wait(50)
    Send("b")       ; Skip blanks
    Wait(50)
    Send("{Enter}")
    ShowHUD("Paste Formulas + Formats (Skip Blanks)", 800)
}

; Filter operations
ToggleFilter() {
    ; Alt + H -> S -> F (Toggle Filter)
    Send("!h")      ; Home ribbon
    Wait(Timing.RIBBON_DELAY)
    Send("s")       ; Sort & Filter
    Wait(Timing.NAV_DELAY)
    Send("f")       ; Filter
    ShowHUD("Toggle Filter", 800)
}

ClearFilter() {
    ; Alt + H -> S -> C (Clear Filter)
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
SC03A & Space::{
    helpText := "GLOBAL: semicolon->Ctrl+Windows+Alt+T | Ctrl+Alt+H->Toggle Overlay"
    helpText .= Chr(10) . "PASTE: v,f,t,w,s,x,l,p + Numpad+ -"
    helpText .= Chr(10) . "FILTER: Numpad/ (Toggle) Numpad* (Clear)"
    helpText .= Chr(10) . "DELETE: / (Ctrl=Row)"
    helpText .= Chr(10) . "FORMAT: 1,2,3,4,6 + 9,a,k,5,m,d + F1-F4 + r,b,o,i,c,y,j + q,F5,F6,F11,F12 + Numpad.,0"
    helpText .= Chr(10) . "NAV: [,],=,-,.,,g,8,h,Right,Left + Numpad8,2,4,6,7,9"
    helpText .= Chr(10) . "DATA: u,F8,n,e,F7,F9 + z,Backspace,Delete"
    helpText .= Chr(10) . "Note: Excel-specific hotkeys only work in Excel"
    ShowOSD("CAPSLOCK LAYER - EXCEL SHORTCUTS", helpText, 2500, "top-center", 720)
}

; CapsLock + Semicolon hotkey - sends Ctrl+Windows+Alt+T for keep-on-top
SC03A & SC027::Send("^#!t")

#HotIf

; -----------------------------------------------------------------------------
; Excel-Specific CapsLock Hold Layer - All hotkeys in one block
; -----------------------------------------------------------------------------
#HotIf (IsExcel() && GetKeyState("SC03A","P"))

; PASTE operations
SC03A & v::Do(() => PasteSpecial("values"), "Paste Values", "CapsLock+V")                    ; Paste Values
SC03A & f::                                                                            ; CapsLock+F
{
    if GetKeyState("Ctrl", "P") {
        return
    }
    Do(() => PasteSpecial("formulas"), "Paste Formulas", "CapsLock+F")
}
SC03A & t::Do(() => PasteSpecial("formats"), "Paste Formats", "CapsLock+T")                   ; Paste Formats
SC03A & w::Do(() => PasteSpecial("colwidths"), "Paste Column Widths", "CapsLock+W")                 ; Paste Column Widths
SC03A & s::Do(() => PasteFormulasWithFormat(), "Paste Formulas + Formats", "CapsLock+S")   ; Paste Formulas + Number Formats, Skip Blanks
SC03A & x::Do(() => PasteSpecial("values", Map("transpose", true)), "Paste Values (Transpose)", "CapsLock+X")  ; Paste Values (Transpose)
SC03A & l::Do(() => PasteLink(), "Paste Link", "CapsLock+L")                               ; Paste Link
SC03A & p::Do(() => Send("^!v"), "Paste Special Dialog", "CapsLock+P")                               ; Paste Special dialog

; Numpad operations
SC03A & /::
{
    if GetKeyState("Ctrl","P") {
        Do(DeleteSheetRow, "Delete Row", "CapsLock+Ctrl+/")
    } else {
        Do(DeleteSheetColumn, "Delete Column", "CapsLock+/")
    }
}
SC03A & NumpadDiv::Do(() => ToggleFilter(), "Toggle Filter", "CapsLock+Numpad/")                    ; Toggle Filter (Alt+H+S+F)
SC03A & NumpadMult::Do(() => ClearFilter(), "Clear Filter", "CapsLock+Numpad*")                      ; Clear Filter (Alt+H+S+C)
SC03A & NumpadAdd::Do(() => PasteOperation("add"), "Paste Add", "CapsLock+Numpad+")             ; Paste Operation Add
SC03A & NumpadSub::Do(() => PasteOperation("subtract"), "Paste Subtract", "CapsLock+Numpad-")        ; Paste Operation Subtract

; FORMAT - Row types
SC03A & 1::Do(() => Send("^!+1"), "Font Color", "CapsLock+1")                       ; Custom font color macro
SC03A & 2::Do(() => ClassifySubtotal(), "Subtotal Format", "CapsLock+2")                        ; Subtotal
SC03A & 3::Do(() => ClassifyMajorTotal(), "Major Total Format", "CapsLock+3")                      ; Major total
SC03A & 4::Do(() => ClassifyGrandTotal(), "Grand Total Format", "CapsLock+4")                      ; Grand total
SC03A & 6::Do(() => SetRowHeight(5), "Row Height 5pt", "CapsLock+6")                           ; Row height 5pt

; FORMAT - Number formats
SC03A & 9::Do(() => SetNumberFormat("general"), "General Format", "CapsLock+9")                ; General format
SC03A & a::Do(() => SetRowHeight(5), "Row Height 5pt", "CapsLock+A")                           ; Row height 5pt
SC03A & k::Do(() => Send("^!+k"), "Number Format", "CapsLock+K")                    ; Custom number format macro
SC03A & 5::Do(() => SetNumberFormat("percent"), "Percent Format", "CapsLock+5")                ; Percent format
SC03A & m::Do(() => SetNumberFormat("month"), "Month Format", "CapsLock+M")                  ; Month format
SC03A & d::Do(() => SetNumberFormat("date"), "Date Format", "CapsLock+D")                   ; Date format

; ALIGNMENT
SC03A & F1::Do(() => SetAlignment("left"), "Align Left", "CapsLock+F1")                     ; Align Left
SC03A & F2::Do(() => SetAlignment("center"), "Align Center", "CapsLock+F2")                   ; Align Center
SC03A & F3::Do(() => SetAlignment("right"), "Align Right", "CapsLock+F3")                    ; Align Right
SC03A & F4::Do(() => ToggleWrapText(), "Toggle Wrap Text", "CapsLock+F4")                         ; Toggle Wrap

; BORDERS / MACROS
SC03A & r::                                                                            ; CapsLock+R or CapsLock+Ctrl+R
{
    if GetKeyState("Ctrl","P") {
        Do(AutoFitRows, "AutoFit Row Height", "CapsLock+Ctrl+R")
    } else {
        Do(() => ApplyRightBorder(), "Apply Right Border", "CapsLock+R")
    }
}
SC03A & b::Do(() => ApplyBottomBorder(), "Apply Bottom Border", "CapsLock+B")                       ; Apply BottomThinBorder macro
SC03A & o::Do(() => SetBorders("outline"), "Outline Borders", "CapsLock+O")                     ; Outline borders
SC03A & i::Do(() => SetBorders("inside"), "Inside Borders", "CapsLock+I")                      ; Inside borders
SC03A & c::                                                                            ; CapsLock+C or CapsLock+Ctrl+C
{
    if GetKeyState("Ctrl","P") {
        Do(AutoFitColumns, "AutoFit Column Width", "CapsLock+Ctrl+C")
    } else {
        Do(() => ClearBorders(), "Clear Borders", "CapsLock+C")
    }
}
SC03A & y::Do(() => SetBorderLine("top", "double"), "Top Double Border", "CapsLock+Y")            ; Top double border
SC03A & j::Do(() => SetBorderLine("left", "thick"), "Left Thick Border", "CapsLock+J")            ; Left thick border
; SC03A & `;:: removed - now used globally for Ctrl+Windows+Alt+T

; DIVIDERS / SIZING
SC03A & q::                                                                            ; CapsLock+Q or CapsLock+Ctrl+Q
{
    if GetKeyState("Ctrl","P") {
        Do(() => SetRowHeight(5), "Set Row Height 5pt", "CapsLock+Ctrl+Q")
    } else {
        Do(() => SetColumnWidth(0.5), "Set Column Width", "CapsLock+Q")
    }
}
SC03A & F5::Do(() => AutoFitColumns(), "AutoFit Columns", "CapsLock+F5")                         ; AutoFit Columns
SC03A & F6::Do(() => AutoFitRows(), "AutoFit Rows", "CapsLock+F6")                            ; AutoFit Rows
SC03A & F11::Do(IncreaseIndent, "Increase Indent", "CapsLock+F11")                                ; Increase indent
SC03A & F12::Do(DecreaseIndent, "Decrease Indent", "CapsLock+F12")                                ; Decrease indent
SC03A & NumpadDot::Do(AddDecimalPlace, "Add Decimal Place", "CapsLock+Numpad.")                         ; Add decimal place
SC03A & Numpad0::Do(RemoveDecimalPlace, "Remove Decimal Place", "CapsLock+Numpad0")                        ; Remove decimal place

; NAVIGATION
SC03A & [::Do(() => JumpToPrevDivider(), "Prev Divider", "CapsLock+[")                       ; Prev divider
SC03A & ]::Do(() => JumpToNextDivider(), "Next Divider", "CapsLock+]")                       ; Next divider
SC03A & =::Do(() => JumpToBlockEdge("first"), "First Block Edge", "CapsLock+=")                  ; First block edge
SC03A & -::Do(() => JumpToBlockEdge("last"), "Last Block Edge", "CapsLock+-")                   ; Last block edge
SC03A & ,::Do(() => Send("^{PgUp}"), "Previous Sheet", "CapsLock+,")                           ; Previous sheet
SC03A & .::Do(() => Send("^{PgDn}"), "Next Sheet", "CapsLock+.")                           ; Next sheet
SC03A & g::
{
    if GetKeyState("Ctrl","P") {
        Do(GroupAndCollapseSelection, "Group and Collapse", "CapsLock+Ctrl+G")
    } else {
        Do(() => Send("^g"), "Go To", "CapsLock+G")
    }
}
SC03A & 8::Do(() => Send("^+8"), "Current Region", "CapsLock+8")                               ; Current Region
SC03A & h::Do(() => Send("^!+h"), "Highlight Cell", "CapsLock+H")                  ; Custom macro Ctrl+Alt+Shift+H

SC03A & Right::
{
    if GetKeyState("Ctrl","P") {
        Do(() => Send("^{PgDn}"), "Next Sheet", "CapsLock+Ctrl+Right")
    } else if GetKeyState("Shift","P") {
        Do(() => Send("{Shift down}{Right 11}{Shift up}"), "Select Right 11", "CapsLock+Shift+Right")
    } else {
        Do(() => Send("{Right 12}"), "Move Right 12", "CapsLock+Right")
    }
}

SC03A & Left::
{
    if GetKeyState("Ctrl","P") {
        Do(() => Send("^{PgUp}"), "Previous Sheet", "CapsLock+Ctrl+Left")
    } else if GetKeyState("Shift","P") {
        Do(() => Send("{Shift down}{Left 11}{Shift up}"), "Select Left 11", "CapsLock+Shift+Left")
    } else {
        Do(() => Send("{Left 12}"), "Move Left 12", "CapsLock+Left")
    }
}

; Numpad navigation
SC03A & Numpad8::Do(() => Send("^{Up}"), "Ctrl+Up", "CapsLock+Numpad8")                       ; Ctrl+Arrow Up
SC03A & Numpad2::Do(() => Send("^{Down}"), "Ctrl+Down", "CapsLock+Numpad2")                     ; Ctrl+Arrow Down
SC03A & Numpad4::Do(() => Send("^{Left}"), "Ctrl+Left", "CapsLock+Numpad4")                     ; Ctrl+Arrow Left
SC03A & Numpad6::Do(() => Send("^{Right}"), "Ctrl+Right", "CapsLock+Numpad6")                    ; Ctrl+Arrow Right
SC03A & Numpad7::Do(() => Send("^{Home}"), "Ctrl+Home", "CapsLock+Numpad7")                     ; Ctrl+Home (A1)
SC03A & Numpad9::Do(() => Send("^{End}"), "Ctrl+End", "CapsLock+Numpad9")                      ; Ctrl+End

; DATA / CLEANUP
SC03A & u::Do(() => TrimInPlace(), "Trim In Place", "CapsLock+U")                             ; TRIM
SC03A & F8::Do(() => CleanInPlace(), "Clean In Place", "CapsLock+F8")                           ; CLEAN
SC03A & n::Do(() => CoerceToNumber(), "Convert to Number", "CapsLock+N")                          ; Convert to Number
SC03A & e::Do(() => Send("!de"), "Text to Columns", "CapsLock+E")                               ; Text to Columns
SC03A & F7::Do(() => Send("^+l"), "Toggle AutoFilter", "CapsLock+F7")                              ; Toggle AutoFilter
; SC03A & F9::Do(() => FreezeAtActiveCell(), "Freeze Panes")                     ; Freeze panes - moved to Ctrl+CapsLock+F

; CLEARS
SC03A & z::Do(ClearFormatsSel, "Clear Formats", "CapsLock+Z")                                 ; Clear Formats
SC03A & Backspace::Do(ClearContentsSel, "Clear Contents", "CapsLock+Backspace")                        ; Clear Contents
SC03A & Delete::Do(ClearAllSel, "Clear All", "CapsLock+Delete")                                ; Clear All


#HotIf

; (Ctrl-layer removed: Ctrl behavior is handled inside base CapsLock combos)

; Win key safety - prevent Start menu while holding CapsLock in Excel
#HotIf (IsExcel() && GetKeyState("SC03A","P"))
LWin::Return
LWin up::Return
#HotIf




