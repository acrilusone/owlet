/*
Meeting Widget (AutoHotkey v2 version, v11)
Windows-only compact Outlook calendar widget.
The gear menu includes About / Help and identical widget clone controls.
Each clone runs from the same script or compiled EXE and remembers its own position.

Requires AutoHotkey v2.0+  (https://www.autohotkey.com)
Run: double-click this file.

Controls
- Main window: hover or click = popup, drag = move, right-click = join meeting.
- Gear icon: click = About / Help, right-click = menu (Add another widget,
  Close this widget, Refresh now, Restart this widget, Open log, Reset position,
  Start with Windows, Exit).
- Popup: click a meeting to show MEETING DETAILS / JOIN MEETING buttons next to it,
  right-click a meeting to join it. The bottom buttons act on the main-window meeting.

v5.1: fixed "This line will never execute" warning (startup code moved into Main()).
v6:   all fonts (size + bold/italic) are set in the FONTS section of the configuration;
      popup meetings no longer show countdowns.
v7:   added a configurable border around the main taskbar window
      (MAIN_BORDER_COLOR / MAIN_BORDER_WIDTH in the Appearance section).
v8:   added identical widget clones from one script/EXE; each clone has its own
      remembered position, and the desired clone count is restored at startup.
      Added Add another widget, Close this widget, and About / Help.
v9:   redesigned meeting interactions. Double-click joins or shows the physical
      location; right-click opens a meeting-specific menu. Added persistent
      user hiding through User-Hidden Meetings.txt. Removed duplicate main meeting,
      popup gear, popup action buttons, and bottom action bar. Added a light-gray
      main gear and a separator between NEXT and the gear.
      2026-11-05 16:49 CST - Fixed clone-slot allocation by reading Win32
      GetLastError immediately after CreateMutexW instead of A_LastError.
      2026-11-05 16:52 CST - Fixed meeting right-click menu variable collision and
      changed every gear click to open the gear menu without changing gear width.
      2026-11-05 16:57 CST - Restored the original v7 outer-border and popup-
      separator colors/implementation; retained only the new 1px gear separator.
      2026-11-05 16:59 CST - Corrected border geometry: removed the extra pixel
      added to main width for the gear separator, overlaid the gear separator on
      the existing boundary, and restored v7 popup row positioning/insets.
      2026-11-05 17:02 CST - Found the remaining thickness cause: v7 uses the
      border-colored GUI background for separators, and NEXT already reserved a
      2px trailing gap. Removed the added separator control and changed only the
      NEXT-to-gear gap to 1px. Original v7 outer border and other geometry remain.
      2026-11-05 17:05 CST - Confirmed nextSpace uses gNextW + 1 exactly as requested.
      2026-11-05 17:12 CST - Corrected the remaining doubled meeting-to-NEXT
      divider by changing both innerW and nextX from +2 to +1. All main internal
      divider gaps are now 1px while preserving the original outer border.
      2026-11-05 17:15 CST - Reverted the pixel-by-pixel experiments and restored
      the exact v7 BuildMain geometry (+2 values) and popup perimeter geometry
      (inner=gW-2, y=1, x=1 cells), while retaining v9 behavior changes.
      2026-11-05 17:20 CST - Rendering experiment: preserved v7 dimensions and
      coordinates but changed GUI backgrounds to black and drew explicit 1px outer
      borders, internal dividers, popup perimeter, and popup row separators.
      2026-11-05 17:27 CST - Renamed generated files to remove "taskbar"; consolidated
      clone positions/count into meeting widget position.ini; changed hidden-meeting
      lines to short IDs plus friendly date/time and subject, with old-format migration.
      2026-11-05 17:37 CST - Moved every user setting to Meeting Widget Preferences.txt,
      made logging off by default, and added visible status suffixes for displayed
      tentative, free, cancelled, declined, expired, and all-day meetings.
      2026-11-05 17:49 CST - Separated Outlook RSVP ResponseStatus from Show As/
      BusyStatus. Added independent visibility filters and accurate title labels such
      as not responded, tentative response, show as tentative, free, and declined.
      2026-11-05 19:29 CST - Added a tabbed Settings window that edits the external
      preferences file with checkboxes, dropdowns, color pickers, and numeric fields.
      Reordered and renamed the gear menu and moved startup/help controls to General.
      2026-11-05 20:01 CST - Fixed AutoHotkey v2 catch syntax in both numeric
      Settings code paths; catch blocks now use braces and numeric values are clamped.
*/
#Requires AutoHotkey v2.0
#SingleInstance Off
Persistent

;@Ahk2Exe-SetName Meeting Widget
;@Ahk2Exe-SetDescription Outlook calendar meeting widget with identical clone support
;@Ahk2Exe-SetVersion 11.0.0
;@Ahk2Exe-SetProductName Meeting Widget
;@Ahk2Exe-SetProductVersion 11.0.0

; ============================================================================
; USER PREFERENCES
; All user-editable settings are loaded from Meeting Widget Preferences.txt.
; Keep that file beside this script or compiled EXE, save changes, then restart.
; ============================================================================

PREFERENCES_FILE := A_ScriptDir "\Meeting Widget Preferences.txt"

MAIN_TEXT_CHARS := PrefInt("Main Window", "TextChars", 27)
MAIN_ANCHOR := PrefText("Main Window", "Anchor", "right")
MAIN_HORIZONTAL_OFFSET := PrefInt("Main Window", "HorizontalOffset", 185)
MAIN_VERTICAL_OFFSET := PrefInt("Main Window", "VerticalOffset", 0)
ALLOW_DRAG := PrefBool("Main Window", "AllowDrag", true)
SAVE_POSITION := PrefBool("Main Window", "SavePosition", true)
SNAP_TO_TASKBAR := PrefBool("Main Window", "SnapToTaskbar", true)
SNAP_TO_TOP_EDGE := PrefBool("Main Window", "SnapToTopEdge", true)
SNAP_DISTANCE_PX := PrefInt("Main Window", "SnapDistancePx", 20)
SHOW_COUNTDOWN := PrefBool("Main Window", "ShowCountdown", true)
SHOW_NEXT_MEETING_TIME := PrefBool("Main Window", "ShowNextMeetingTime", true)
PROMOTE_NEXT_AFTER_ACTIVE_MINUTES := PrefInt("Main Window", "PromoteNextAfterActiveMinutes", 10)

POPUP_MEETING_COUNT := PrefInt("Popup", "MeetingCount", 4)
POPUP_OPEN_ON_HOVER := PrefBool("Popup", "OpenOnHover", true)
POPUP_OPEN_ON_CLICK := PrefBool("Popup", "OpenOnClick", true)
POPUP_HOVER_DELAY_MS := PrefInt("Popup", "HoverDelayMs", 300)
POPUP_CLOSE_DELAY_MS := PrefInt("Popup", "CloseDelayMs", 450)

USE_OUTLOOK_EVENTS := PrefBool("Refresh", "UseOutlookEvents", true)
CHANGE_DEBOUNCE_SECONDS := PrefInt("Refresh", "ChangeDebounceSeconds", 2)
SAFETY_REFRESH_MINUTES := PrefInt("Refresh", "SafetyRefreshMinutes", 15)
FALLBACK_POLL_SECONDS := PrefInt("Refresh", "FallbackPollSeconds", 60)
RETRY_AFTER_ERROR_SECONDS := PrefInt("Refresh", "RetryAfterErrorSeconds", 60)
LOOKAHEAD_DAYS := PrefInt("Refresh", "LookaheadDays", 30)
MAX_MEETINGS_TO_SCAN := PrefInt("Refresh", "MaxMeetingsToScan", 60)
KEEP_ON_TOP_INTERVAL_MS := PrefInt("Refresh", "KeepOnTopIntervalMs", 1500)

MAIN_HIDDEN_STATUSES := PrefList("Meeting Visibility", "MainHiddenStatuses", "cancelled,expired,all_day")
POPUP_HIDDEN_STATUSES := PrefList("Meeting Visibility", "PopupHiddenStatuses", "cancelled,expired,all_day")
MAIN_HIDDEN_RESPONSES := PrefList("Meeting Visibility", "MainHiddenResponses", "not_responded,declined")
POPUP_HIDDEN_RESPONSES := PrefList("Meeting Visibility", "PopupHiddenResponses", "not_responded,declined")
MAIN_HIDDEN_AVAILABILITY := PrefList("Meeting Visibility", "MainHiddenAvailability", "free")
POPUP_HIDDEN_AVAILABILITY := PrefList("Meeting Visibility", "PopupHiddenAvailability", "free")
MAIN_OUTLOOK_RESTRICT_FILTER := PrefText("Meeting Visibility", "MainOutlookRestrictFilter", "")
POPUP_OUTLOOK_RESTRICT_FILTER := PrefText("Meeting Visibility", "PopupOutlookRestrictFilter", "")

BACKGROUND := PrefText("Appearance", "Background", "#000000")
TEXT_COLOR := PrefText("Appearance", "TextColor", "#FFB000")
SECONDARY_TEXT_COLOR := PrefText("Appearance", "SecondaryTextColor", "#C09200")
NEXT_TIME_COLOR := PrefText("Appearance", "NextTimeColor", "#BF4E14")
HOVER_BACKGROUND := PrefText("Appearance", "HoverBackground", "#FFB000")
HOVER_TEXT_COLOR := PrefText("Appearance", "HoverTextColor", "#000000")
SEPARATOR_COLOR := PrefText("Appearance", "SeparatorColor", "#7A5400")
GEAR_COLOR := PrefText("Appearance", "GearColor", "#D0D0D0")
DISABLED_TEXT_COLOR := PrefText("Appearance", "DisabledTextColor", "#4D3A00")

MAIN_BORDER_COLOR := PrefText("Borders", "MainBorderColor", "#FFB000")
MAIN_BORDER_WIDTH := PrefInt("Borders", "MainBorderWidth", 1)

FONT_FAMILY := PrefText("Fonts", "FontFamily", "Segoe UI")
FONT_TIME := PrefFont("Time", 10, "bold")
FONT_COUNTDOWN := PrefFont("Countdown", 8, "normal")
FONT_TITLE := PrefFont("Title", 8, "normal")
FONT_LOCATION := PrefFont("Location", 8, "normal")
FONT_NEXT := PrefFont("Next", 7, "normal")
FONT_POPUP_BUTTONS := {size: 10, style: "bold"}
FONT_ROW_BUTTONS := {size: 10, style: "bold"}
GEAR_FONT_FAMILY := PrefText("Fonts", "GearFontFamily", "Segoe UI Symbol")
FONT_GEAR := PrefFont("Gear", 12, "normal")

DEBUG := PrefBool("Logging", "Enabled", false)
LOG_MAX_KB := PrefInt("Logging", "MaxKB", 256)


POSITION_FILE := A_ScriptDir "\meeting widget position.ini"
APP_STATE_FILE := POSITION_FILE
HIDDEN_MEETINGS_FILE := A_ScriptDir "\User-Hidden Meetings.txt"
INSTANCE_ID := ClaimInstanceSlot()
LOG_FILE := A_ScriptDir "\meeting widget " INSTANCE_ID ".log"
STARTUP_SHORTCUT := A_Startup "\Meeting Widget.lnk"
DRAG_THRESHOLD_PX := 4
PAD := 6

; ---------- state ----------
gInstanceMutex := 0
gLastMeetingClickTick := 0, gLastMeetingClickKey := ""
gNs := "", gOutlook := "", gWatchItems := "", gEventsOn := false
gLinkCache := Map()
gMainMeetings := [], gPopupMeetings := []
gCurrentMain := ""
gLoaded := false, gError := "", gLastSig := ""
gRefreshing := false, gChangePending := false
gLastDay := FormatTime(A_Now, "yyyyMMdd")
gPosX := 0, gPosY := 0, gW := 0, gH := 0, gPopupH := 0, gBorder := 0
gPopupVisible := false, gPressing := false
gHoverStart := 0, gOutsideSince := 0, gHoverTimerOn := false
gExpanded := 0, gHoverKey := ""
gRows := []
gSettingsGui := "", gSettingsTabs := "", gSettingsControls := Map()


PrefText(section, key, defaultValue) {
    global PREFERENCES_FILE
    if !FileExist(PREFERENCES_FILE)
        return defaultValue
    try return IniRead(PREFERENCES_FILE, section, key, defaultValue)
    catch
        return defaultValue
}

PrefInt(section, key, defaultValue) {
    value := PrefText(section, key, defaultValue)
    try return Integer(value)
    catch
        return defaultValue
}

PrefBool(section, key, defaultValue) {
    value := StrLower(Trim(String(PrefText(section, key, defaultValue ? "true" : "false"))))
    if value = "true" || value = "1" || value = "yes" || value = "on"
        return true
    if value = "false" || value = "0" || value = "no" || value = "off"
        return false
    return defaultValue
}

PrefList(section, key, defaultValue) {
    text := Trim(PrefText(section, key, defaultValue))
    values := []
    if text = ""
        return values
    for item in StrSplit(text, ",") {
        item := StrLower(Trim(item))
        if item != ""
            values.Push(item)
    }
    return values
}

PrefFont(prefix, defaultSize, defaultStyle) {
    return {size: PrefInt("Fonts", prefix "Size", defaultSize),
            style: PrefText("Fonts", prefix "Style", defaultStyle)}
}

OpenPreferences(*) {
    global PREFERENCES_FILE
    if !FileExist(PREFERENCES_FILE) {
        MsgBox("Meeting Widget Preferences.txt was not found beside the app.", "Meeting Widget", "Iconx")
        return
    }
    Run('notepad.exe "' PREFERENCES_FILE '"')
}


SettingsSchema() {
    return [
        {tab:"Meeting Visibility", section:"Meeting Visibility", key:"MainHiddenStatuses", label:"Hide lifecycle statuses in main widget", type:"multi", options:["cancelled","expired","all_day"]},
        {tab:"Meeting Visibility", section:"Meeting Visibility", key:"PopupHiddenStatuses", label:"Hide lifecycle statuses in popup", type:"multi", options:["cancelled","expired","all_day"]},
        {tab:"Meeting Visibility", section:"Meeting Visibility", key:"MainHiddenResponses", label:"Hide RSVP responses in main widget", type:"multi", options:["not_responded","tentative","accepted","declined","organizer","none"]},
        {tab:"Meeting Visibility", section:"Meeting Visibility", key:"PopupHiddenResponses", label:"Hide RSVP responses in popup", type:"multi", options:["not_responded","tentative","accepted","declined","organizer","none"]},
        {tab:"Meeting Visibility", section:"Meeting Visibility", key:"MainHiddenAvailability", label:"Hide Show As values in main widget", type:"multi", options:["free","tentative","busy","oof","working_elsewhere"]},
        {tab:"Meeting Visibility", section:"Meeting Visibility", key:"PopupHiddenAvailability", label:"Hide Show As values in popup", type:"multi", options:["free","tentative","busy","oof","working_elsewhere"]},
        {tab:"Meeting Visibility", section:"Meeting Visibility", key:"MainOutlookRestrictFilter", label:"Advanced Outlook filter - main", type:"text"},
        {tab:"Meeting Visibility", section:"Meeting Visibility", key:"PopupOutlookRestrictFilter", label:"Advanced Outlook filter - popup", type:"text"},

        {tab:"Fonts", section:"Fonts", key:"FontFamily", label:"Main font family", type:"font"},
        {tab:"Fonts", section:"Fonts", key:"TimeSize", label:"Time size", type:"number", min:6, max:72},
        {tab:"Fonts", section:"Fonts", key:"TimeStyle", label:"Time style", type:"select", options:["normal","bold","italic","bold italic"]},
        {tab:"Fonts", section:"Fonts", key:"CountdownSize", label:"Countdown size", type:"number", min:6, max:72},
        {tab:"Fonts", section:"Fonts", key:"CountdownStyle", label:"Countdown style", type:"select", options:["normal","bold","italic","bold italic"]},
        {tab:"Fonts", section:"Fonts", key:"TitleSize", label:"Title size", type:"number", min:6, max:72},
        {tab:"Fonts", section:"Fonts", key:"TitleStyle", label:"Title style", type:"select", options:["normal","bold","italic","bold italic"]},
        {tab:"Fonts", section:"Fonts", key:"LocationSize", label:"Location size", type:"number", min:6, max:72},
        {tab:"Fonts", section:"Fonts", key:"LocationStyle", label:"Location style", type:"select", options:["normal","bold","italic","bold italic"]},
        {tab:"Fonts", section:"Fonts", key:"NextSize", label:"NEXT size", type:"number", min:6, max:72},
        {tab:"Fonts", section:"Fonts", key:"NextStyle", label:"NEXT style", type:"select", options:["normal","bold","italic","bold italic"]},
        {tab:"Fonts", section:"Fonts", key:"GearFontFamily", label:"Gear font family", type:"font"},
        {tab:"Fonts", section:"Fonts", key:"GearSize", label:"Gear size", type:"number", min:6, max:72},
        {tab:"Fonts", section:"Fonts", key:"GearStyle", label:"Gear style", type:"select", options:["normal","bold","italic","bold italic"]},

        {tab:"Appearance", section:"Appearance", key:"Background", label:"Background", type:"color"},
        {tab:"Appearance", section:"Appearance", key:"TextColor", label:"Text color", type:"color"},
        {tab:"Appearance", section:"Appearance", key:"SecondaryTextColor", label:"Secondary text color", type:"color"},
        {tab:"Appearance", section:"Appearance", key:"NextTimeColor", label:"NEXT time color", type:"color"},
        {tab:"Appearance", section:"Appearance", key:"HoverBackground", label:"Hover background", type:"color"},
        {tab:"Appearance", section:"Appearance", key:"HoverTextColor", label:"Hover text color", type:"color"},
        {tab:"Appearance", section:"Appearance", key:"SeparatorColor", label:"Popup separators", type:"color"},
        {tab:"Appearance", section:"Appearance", key:"GearColor", label:"Gear color", type:"color"},
        {tab:"Appearance", section:"Appearance", key:"DisabledTextColor", label:"Disabled text color", type:"color"},

        {tab:"Borders", section:"Borders", key:"MainBorderColor", label:"Main border color", type:"color"},
        {tab:"Borders", section:"Borders", key:"MainBorderWidth", label:"Main border width (px)", type:"number", min:0, max:10},

        {tab:"Refresh", section:"Refresh", key:"UseOutlookEvents", label:"Use Outlook change events", type:"bool"},
        {tab:"Refresh", section:"Refresh", key:"ChangeDebounceSeconds", label:"Change debounce (seconds)", type:"number", min:0, max:60},
        {tab:"Refresh", section:"Refresh", key:"SafetyRefreshMinutes", label:"Safety refresh (minutes)", type:"number", min:1, max:1440},
        {tab:"Refresh", section:"Refresh", key:"FallbackPollSeconds", label:"Fallback poll (seconds)", type:"number", min:10, max:3600},
        {tab:"Refresh", section:"Refresh", key:"RetryAfterErrorSeconds", label:"Retry after error (seconds)", type:"number", min:5, max:3600},
        {tab:"Refresh", section:"Refresh", key:"LookaheadDays", label:"Look-ahead days", type:"number", min:1, max:365},
        {tab:"Refresh", section:"Refresh", key:"MaxMeetingsToScan", label:"Maximum meetings to scan", type:"number", min:10, max:1000},
        {tab:"Refresh", section:"Refresh", key:"KeepOnTopIntervalMs", label:"Keep-on-top interval (ms)", type:"number", min:250, max:60000},

        {tab:"Logging", section:"Logging", key:"Enabled", label:"Enable logging", type:"bool"},
        {tab:"Logging", section:"Logging", key:"MaxKB", label:"Maximum log size (KB)", type:"number", min:16, max:10240},

        {tab:"Main Window", section:"Main Window", key:"TextChars", label:"Title width (characters)", type:"number", min:10, max:100},
        {tab:"Main Window", section:"Main Window", key:"Anchor", label:"Default anchor", type:"select", options:["left","center","right"]},
        {tab:"Main Window", section:"Main Window", key:"HorizontalOffset", label:"Horizontal offset", type:"number", min:-5000, max:5000},
        {tab:"Main Window", section:"Main Window", key:"VerticalOffset", label:"Vertical offset", type:"number", min:-5000, max:5000},
        {tab:"Main Window", section:"Main Window", key:"AllowDrag", label:"Allow dragging", type:"bool"},
        {tab:"Main Window", section:"Main Window", key:"SavePosition", label:"Remember widget position", type:"bool"},
        {tab:"Main Window", section:"Main Window", key:"SnapToTaskbar", label:"Snap to taskbar", type:"bool"},
        {tab:"Main Window", section:"Main Window", key:"SnapToTopEdge", label:"Snap to top edge", type:"bool"},
        {tab:"Main Window", section:"Main Window", key:"SnapDistancePx", label:"Snap distance (px)", type:"number", min:0, max:500},
        {tab:"Main Window", section:"Main Window", key:"ShowCountdown", label:"Show countdown", type:"bool"},
        {tab:"Main Window", section:"Main Window", key:"ShowNextMeetingTime", label:"Show NEXT meeting time", type:"bool"},
        {tab:"Main Window", section:"Main Window", key:"PromoteNextAfterActiveMinutes", label:"Promote next after active minutes", type:"number", min:0, max:1440},

        {tab:"Popup", section:"Popup", key:"MeetingCount", label:"Meetings to show", type:"number", min:1, max:20},
        {tab:"Popup", section:"Popup", key:"OpenOnHover", label:"Open on hover", type:"bool"},
        {tab:"Popup", section:"Popup", key:"OpenOnClick", label:"Open on click", type:"bool"},
        {tab:"Popup", section:"Popup", key:"HoverDelayMs", label:"Hover delay (ms)", type:"number", min:0, max:10000},
        {tab:"Popup", section:"Popup", key:"CloseDelayMs", label:"Close delay (ms)", type:"number", min:0, max:10000}
    ]
}

OpenSettingsWindow(*) {
    global gSettingsGui, gSettingsTabs, gSettingsControls, STARTUP_SHORTCUT
    if IsObject(gSettingsGui) {
        try gSettingsGui.Show()
        return
    }
    gSettingsControls := Map()
    gSettingsGui := Gui("+Resize +MinSize700x560", "Meeting Widget Settings")
    gSettingsGui.SetFont("s9", "Segoe UI")
    tabs := ["General","Meeting Visibility","Fonts","Appearance","Borders","Refresh","Logging","Main Window","Popup"]
    gSettingsTabs := gSettingsGui.Add("Tab3", "x10 y10 w760 h520", tabs)

    gSettingsTabs.UseTab("General")
    gSettingsGui.SetFont("s12 bold")
    gSettingsGui.Add("Text", "x35 y55 w680", "Meeting Widget")
    gSettingsGui.SetFont("s9 norm")
    gSettingsGui.Add("Text", "x35 y88 w680", "Version 11.0.0")
    startup := gSettingsGui.Add("CheckBox", "x35 y125 w300", "Start with Windows")
    startup.Value := FileExist(STARTUP_SHORTCUT) ? 1 : 0
    startup.OnEvent("Click", SettingsToggleStartup)
    gSettingsGui.Add("Text", "x55 y153 w650 h45", "When enabled, a shortcut is placed in the current user's Windows Startup folder. Windows launches items in this folder after sign-in.")
    gSettingsGui.Add("Button", "x35 y215 w165", "Open Startup Folder").OnEvent("Click", (*) => Run("shell:startup"))
    gSettingsGui.Add("Button", "x215 y215 w165", "Open preferences file").OnEvent("Click", OpenPreferences)
    gSettingsGui.Add("Button", "x395 y215 w165", "Open app folder").OnEvent("Click", (*) => Run(A_ScriptDir))
    gSettingsGui.Add("Button", "x35 y270 w165", "About / Help").OnEvent("Click", OpenSettings)
    gSettingsGui.Add("Text", "x35 y325 w670 h80", "Refresh meetings rereads Outlook without restarting the app. Restart this widget reloads fonts, colors, dimensions, filters, and all other saved preferences.")

    yByTab := Map()
    for tab in tabs
        yByTab[tab] := 52
    for spec in SettingsSchema() {
        gSettingsTabs.UseTab(spec.tab)
        y := yByTab[spec.tab]
        AddSettingsControl(spec, y)
        yByTab[spec.tab] := y + (spec.type = "multi" ? 60 : 34)
    }
    gSettingsTabs.UseTab()
    gSettingsGui.Add("Button", "x345 y540 w120", "Cancel").OnEvent("Click", SettingsCancel)
    gSettingsGui.Add("Button", "x475 y540 w120", "Apply").OnEvent("Click", SettingsApply)
    gSettingsGui.Add("Button", "x605 y540 w165 Default", "Save and Restart").OnEvent("Click", SettingsSaveRestart)
    gSettingsGui.OnEvent("Close", SettingsCancel)
    gSettingsGui.Show("w780 h585")
}

AddSettingsControl(spec, y) {
    global gSettingsGui, gSettingsControls
    current := PrefText(spec.section, spec.key, "")
    id := spec.section "|" spec.key
    if spec.type = "multi" {
        gSettingsGui.Add("Text", "x35 y" y " w700", spec.label)
        controls := Map(), x := 45, yy := y + 24
        selected := PrefList(spec.section, spec.key, "")
        for option in spec.options {
            label := SettingsFriendly(option)
            ctl := gSettingsGui.Add("CheckBox", "x" x " y" yy " w115", label)
            ctl.Value := HasValue(selected, option) ? 1 : 0
            controls[option] := ctl
            x += 120
            if x > 650
                x := 45, yy += 25
        }
        gSettingsControls[id] := {spec:spec, type:"multi", controls:controls}
        return
    }
    gSettingsGui.Add("Text", "x35 y" (y+5) " w275", spec.label)
    if spec.type = "bool" {
        ctl := gSettingsGui.Add("CheckBox", "x325 y" (y+3) " w300")
        ctl.Value := PrefBool(spec.section, spec.key, false) ? 1 : 0
    } else if spec.type = "select" || spec.type = "font" {
        options := spec.type = "font" ? SettingsFontList(current) : spec.options
        ctl := gSettingsGui.Add("DropDownList", "x325 y" y " w260", options)
        SettingsChoose(ctl, current, options)
    } else {
        ctl := gSettingsGui.Add("Edit", "x325 y" y " w210", current)
        if spec.type = "number" {
            initialNumber := spec.min
            try {
                initialNumber := Integer(current)
            } catch {
                initialNumber := spec.min
            }
            initialNumber := Max(spec.min, Min(spec.max, initialNumber))
            gSettingsGui.Add("UpDown", "Range" spec.min "-" spec.max, initialNumber)
        }
        if spec.type = "color" {
            preview := gSettingsGui.Add("Text", "x545 y" (y+2) " w30 h21 Background" C(current))
            btn := gSettingsGui.Add("Button", "x585 y" y " w110", "Choose color")
            btn.OnEvent("Click", SettingsPickColor.Bind(ctl, preview))
        }
    }
    gSettingsControls[id] := {spec:spec, type:spec.type, control:ctl}
}

SettingsFriendly(value) {
    labels := Map("all_day","All-day", "not_responded","Not responded", "working_elsewhere","Working elsewhere", "oof","Out of office")
    return labels.Has(value) ? labels[value] : StrTitle(StrReplace(value, "_", " "))
}

SettingsFontList(current) {
    fonts := ["Segoe UI","Segoe UI Symbol","Arial","Calibri","Cambria","Consolas","Tahoma","Times New Roman","Verdana"]
    if current != "" && !HasValue(fonts, current)
        fonts.InsertAt(1, current)
    return fonts
}

SettingsChoose(ctl, value, options) {
    for i, item in options
        if StrLower(item) = StrLower(value) {
            ctl.Choose(i)
            return
        }
    ctl.Choose(1)
}

SettingsPickColor(editCtl, previewCtl, *) {
    global gSettingsGui
    color := ChooseColorDialog(editCtl.Text, gSettingsGui.Hwnd)
    if color != "" {
        editCtl.Text := color
        previewCtl.Opt("Background" C(color))
        previewCtl.Redraw()
    }
}

ChooseColorDialog(initial, owner) {
    rgb := Integer("0x" LTrim(initial, "#"))
    bgr := ((rgb & 0xFF) << 16) | (rgb & 0xFF00) | ((rgb >> 16) & 0xFF)
    custom := Buffer(16 * 4, 0)
    cc := Buffer(A_PtrSize = 8 ? 72 : 36, 0)
    NumPut("UInt", cc.Size, cc, 0)
    NumPut("Ptr", owner, cc, A_PtrSize = 8 ? 8 : 4)
    NumPut("UInt", bgr, cc, A_PtrSize = 8 ? 24 : 12)
    NumPut("Ptr", custom.Ptr, cc, A_PtrSize = 8 ? 32 : 16)
    NumPut("UInt", 0x00000103, cc, A_PtrSize = 8 ? 40 : 20) ; RGBINIT|FULLOPEN|ANYCOLOR
    if !DllCall("Comdlg32.dll\ChooseColorW", "Ptr", cc.Ptr)
        return ""
    bgr := NumGet(cc, A_PtrSize = 8 ? 24 : 12, "UInt")
    rgb := ((bgr & 0xFF) << 16) | (bgr & 0xFF00) | ((bgr >> 16) & 0xFF)
    return Format("#{:06X}", rgb)
}

SettingsToggleStartup(ctl, *) {
    global STARTUP_SHORTCUT
    desired := !!ctl.Value
    exists := FileExist(STARTUP_SHORTCUT)
    if desired != !!exists
        ToggleStartup()
}

SettingsApply(showMessage := true, *) {
    global gSettingsControls, PREFERENCES_FILE
    for id, item in gSettingsControls {
        spec := item.spec
        if item.type = "multi" {
            values := []
            for value, ctl in item.controls
                if ctl.Value
                    values.Push(value)
            text := JoinSettingsValues(values)
        } else if item.type = "bool" {
            text := item.control.Value ? "true" : "false"
        } else {
            text := Trim(item.control.Text)
        }
        if item.type = "number" {
            number := spec.min
            try {
                number := Integer(text)
            } catch {
                number := spec.min
            }
            number := Max(spec.min, Min(spec.max, number))
            text := String(number)
        }
        IniWrite(text, PREFERENCES_FILE, spec.section, spec.key)
    }
    if showMessage
        MsgBox("Preferences saved. Use Save and Restart for visual, layout, font, and visibility changes.", "Meeting Widget", "Iconi")
}

JoinSettingsValues(values) {
    text := ""
    for i, value in values
        text .= (i > 1 ? "," : "") value
    return text
}

SettingsSaveRestart(*) {
    SettingsApply(false)
    Reload()
}

SettingsCancel(*) {
    global gSettingsGui, gSettingsTabs, gSettingsControls
    if IsObject(gSettingsGui)
        gSettingsGui.Destroy()
    gSettingsGui := "", gSettingsTabs := "", gSettingsControls := Map()
}

ClaimInstanceSlot() {
    global gInstanceMutex
    ; Use GetLastError immediately after CreateMutexW. A_LastError is not reliable
    ; for this mutex-claim test in AutoHotkey v2 and caused every slot to look occupied.
    loop 50 {
        id := A_Index
        mutexName := "Local\MeetingTaskbarWidgetClone" id
        hMutex := DllCall("Kernel32\CreateMutexW", "Ptr", 0, "Int", 1, "Str", mutexName, "Ptr")
        if !hMutex
            continue
        lastError := DllCall("Kernel32\GetLastError", "UInt")
        if lastError != 183 { ; ERROR_ALREADY_EXISTS
            gInstanceMutex := hMutex
            return id
        }
        DllCall("Kernel32\CloseHandle", "Ptr", hMutex)
    }
    throw Error("Could not allocate a widget clone slot.")
}

GetSavedWidgetCount() {
    global APP_STATE_FILE
    count := Integer(IniRead(APP_STATE_FILE, "Widgets", "Count", 1))
    return Max(1, Min(count, 20))
}

SetSavedWidgetCount(count) {
    global APP_STATE_FILE
    IniWrite(Max(1, Min(count, 20)), APP_STATE_FILE, "Widgets", "Count")
}

LaunchWidgetClone(*) {
    count := GetSavedWidgetCount()
    if count >= 20 {
        MsgBox("The widget limit is 20.", "Meeting Widget", "Iconi")
        return
    }
    SetSavedWidgetCount(count + 1)
    try Run('"' A_ScriptFullPath '"')
    catch as e {
        SetSavedWidgetCount(count)
        MsgBox("Could not add another widget.`n`n" e.Message, "Meeting Widget", "Iconx")
    }
}

RestoreSavedWidgetClones() {
    global INSTANCE_ID
    if INSTANCE_ID != 1
        return
    count := GetSavedWidgetCount()
    loop count - 1 {
        try Run('"' A_ScriptFullPath '"')
        Sleep(150)
    }
}

CloseThisWidget(*) {
    global POSITION_FILE
    count := GetSavedWidgetCount()
    if count > 1
        SetSavedWidgetCount(count - 1)
    try FileDelete(POSITION_FILE)
    ExitApp()
}

Main()

; ============================================================================
; Startup
; ============================================================================
Main() {
    Log("Starting Meeting Widget (AHK), clone " INSTANCE_ID "...")
    BuildMain()
    BuildPopup()
    BuildMenu()
    PlaceInitially()
    gTitle.Text := "Loading Outlook calendar..."

    OnMessage(0x200, OnMouseMove)        ; WM_MOUSEMOVE
    OnMessage(0x201, OnLButtonDown)      ; WM_LBUTTONDOWN
    OnMessage(0x205, OnRButtonUp)        ; WM_RBUTTONUP
    OnMessage(0x218, OnPowerBroadcast)   ; WM_POWERBROADCAST (wake from sleep)
    OnExit((*) => Log("Exiting clone " INSTANCE_ID "."))

    SetTimer(KeepOnTop, KEEP_ON_TOP_INTERVAL_MS)
    SetTimer(StartupRefresh, -100)
    SetTimer(RestoreSavedWidgetClones, -500)
    Log("Window shown.")
}

StartupRefresh() => RefreshCalendar("startup")
ManualRefresh(*) => RefreshCalendar("manual")
SafetyRefresh() => RefreshCalendar("safety check")
WakeRefresh() => RefreshCalendar("wake from sleep")
NewDayRefresh() => RefreshCalendar("new day")

; ============================================================================
; Logging / helpers
; ============================================================================
C(color) => LTrim(color, "#")

Log(message) {
    if !DEBUG
        return
    try {
        if FileExist(LOG_FILE) && FileGetSize(LOG_FILE) > LOG_MAX_KB * 1024 {
            if FileExist(LOG_FILE ".1")
                FileMove(LOG_FILE ".1", LOG_FILE ".2", 1)
            FileMove(LOG_FILE, LOG_FILE ".1", 1)
        }
        FileAppend("[" FormatTime(A_Now, "yyyy-MM-dd HH:mm:ss") "] " message "`n", LOG_FILE, "UTF-8")
    }
}

HasValue(arr, value) {
    for item in arr
        if item = value
            return true
    return false
}

Truncate(text, limit := MAIN_TEXT_CHARS) {
    return StrLen(text) <= limit ? text : RTrim(SubStr(text, 1, Max(0, limit - 3))) "..."
}

StrRepeat(s, n) {
    out := ""
    loop n
        out .= s
    return out
}

CleanText(value) {
    return Trim(RegExReplace(String(value ?? ""), "\s+", " "))
}

MeasureText(text, size, style := "", font := "") {
    static cache := Map()
    key := size "|" style "|" font "|" text
    if cache.Has(key)
        return cache[key]
    g := Gui("-DPIScale")
    g.SetFont("s" size " " style, font = "" ? FONT_FAMILY : font)
    ctl := g.Add("Text", , text)
    ctl.GetPos(, , &w, &h)
    g.Destroy()
    return cache[key] := {w: w, h: h}
}

; "normal" -> "norm"; everything else is passed straight to Gui.SetFont.
StyleOpts(style) {
    style := Trim(StrLower(style))
    style := RegExReplace(style, "\bnormal\b", "norm")
    return "norm " style
}

FontOpts(spec) => "s" spec.size " " StyleOpts(spec.style)

MeasureSpec(text, spec, family := "") => MeasureText(text, spec.size, StyleOpts(spec.style), family)

TextWidth(ctrl, text) {
    hdc := DllCall("GetDC", "Ptr", ctrl.Hwnd, "Ptr")
    hFont := SendMessage(0x31, 0, 0, ctrl)          ; WM_GETFONT
    old := DllCall("SelectObject", "Ptr", hdc, "Ptr", hFont, "Ptr")
    size := Buffer(8, 0)
    DllCall("GetTextExtentPoint32W", "Ptr", hdc, "WStr", text, "Int", StrLen(text), "Ptr", size)
    DllCall("SelectObject", "Ptr", hdc, "Ptr", old)
    DllCall("ReleaseDC", "Ptr", ctrl.Hwnd, "Ptr", hdc)
    return NumGet(size, 0, "Int")
}

FitText(ctrl, text, maxPx) {
    if maxPx <= 0 || TextWidth(ctrl, text) <= maxPx
        return text
    low := 0, high := StrLen(text)
    while low < high {
        mid := (low + high + 1) // 2
        if TextWidth(ctrl, RTrim(SubStr(text, 1, mid)) "...") <= maxPx
            low := mid
        else
            high := mid - 1
    }
    return RTrim(SubStr(text, 1, low)) "..."
}

; ============================================================================
; Screen / position helpers
; ============================================================================
GetTaskbarRect() {
    try {
        WinGetPos(&x, &y, &w, &h, "ahk_class Shell_TrayWnd")
        return {l: x, t: y, r: x + w, b: y + h}
    }
    return {l: 0, t: A_ScreenHeight - 48, r: A_ScreenWidth, b: A_ScreenHeight}
}

GetVirtualScreen() {
    l := SysGet(76), t := SysGet(77)
    return {l: l, t: t, r: l + SysGet(78), b: t + SysGet(79)}
}

DefaultPosition(w, h) {
    tb := GetTaskbarRect()
    tbw := tb.r - tb.l, tbh := tb.b - tb.t
    if tbw >= tbh {
        y := tb.t + (tbh - h) // 2 - MAIN_VERTICAL_OFFSET
        if MAIN_ANCHOR = "left"
            x := tb.l + MAIN_HORIZONTAL_OFFSET
        else if MAIN_ANCHOR = "center"
            x := tb.l + (tbw - w) // 2 + MAIN_HORIZONTAL_OFFSET
        else
            x := tb.r - w - MAIN_HORIZONTAL_OFFSET
    } else {
        x := tb.l + Max(0, (tbw - w) // 2)
        if MAIN_ANCHOR = "left"
            y := tb.t + MAIN_HORIZONTAL_OFFSET
        else if MAIN_ANCHOR = "center"
            y := tb.t + (tbh - h) // 2 + MAIN_HORIZONTAL_OFFSET
        else
            y := tb.b - h - MAIN_HORIZONTAL_OFFSET
    }
    return {x: x, y: y}
}

Clamp(x, y, w, h) {
    vs := GetVirtualScreen()
    return {x: Max(vs.l, Min(x, vs.r - w)), y: Max(vs.t, Min(y, vs.b - h))}
}

Snap(x, y, w, h) {
    tb := GetTaskbarRect()
    horizontal := (tb.r - tb.l) >= (tb.b - tb.t)
    cy := y + h // 2, cx := x + w // 2
    nearTaskbar := cy >= tb.t - SNAP_DISTANCE_PX && cy <= tb.b + SNAP_DISTANCE_PX
                && cx >= tb.l && cx <= tb.r
    if SNAP_TO_TASKBAR && horizontal && nearTaskbar
        y := tb.t + (tb.b - tb.t - h) // 2
    else if SNAP_TO_TOP_EDGE {
        top := GetVirtualScreen().t
        if y - top <= SNAP_DISTANCE_PX
            y := top
    }
    return Clamp(x, y, w, h)
}

; ============================================================================
; Main window
;   The main window gets an optional MAIN_BORDER_WIDTH-pixel border, drawn the
;   same way the popup draws its separator lines: the Gui's own background is
;   MAIN_BORDER_COLOR, and every control is inset by MAIN_BORDER_WIDTH so the
;   uncovered strip around the edge shows through as a border. With
;   MAIN_BORDER_WIDTH = 0 this is identical to having no border at all.
; ============================================================================
BuildMain() {
    global gMain, gTime, gCountdown, gTitle, gPlace, gNextLabel, gNextTime, gGear
    global gW, gH, gTitleW, gGearW, gNextW, gBorder
    global gTimeH, gTitleH, gPlaceH, gNextH, gTimeRowH, gBtnH, gRowBtnLineH

    gBorder := Max(0, MAIN_BORDER_WIDTH)

    gTimeH := MeasureSpec("Ag", FONT_TIME).h
    gTitleH := MeasureSpec("Ag", FONT_TITLE).h
    gPlaceH := MeasureSpec("Ag", FONT_LOCATION).h
    gNextH := MeasureSpec("Ag", FONT_NEXT).h
    gBtnH := MeasureSpec("Ag", FONT_POPUP_BUTTONS).h
    gRowBtnLineH := MeasureSpec("Ag", FONT_ROW_BUTTONS).h
    gTimeRowH := Max(gTimeH, MeasureSpec("Ag", FONT_COUNTDOWN).h)
    gTitleW := MeasureSpec(StrRepeat("0", MAIN_TEXT_CHARS), FONT_TITLE).w
    gGearW := MeasureSpec(Chr(0x2699), FONT_GEAR, GEAR_FONT_FAMILY).w + 8
    gNextW := Max(MeasureSpec("NEXT", FONT_NEXT).w, MeasureSpec("12:00", FONT_NEXT).w) + 6

    ; Content size (everything inside the border).
    innerH := gTimeRowH + gTitleH + gPlaceH + 2
    if SHOW_NEXT_MEETING_TIME
        innerH := Max(innerH, 3 * gNextH + 2)
    nextSpace := SHOW_NEXT_MEETING_TIME ? gNextW + 2 : 0
    innerW := PAD + gTitleW + 2 + nextSpace + gGearW

    ; Full window size = content + border on all sides.
    gH := innerH + 2 * gBorder
    gW := innerW + 2 * gBorder

    gMain := Gui("+AlwaysOnTop -Caption +ToolWindow -DPIScale", "Meeting Widget")
    gMain.BackColor := C(BACKGROUND)
    gMain.MarginX := 0, gMain.MarginY := 0

    ; Explicit 1px border. Do not rely on uncovered GUI background.
    borderBg := " Background" C(MAIN_BORDER_COLOR)
    gMain.Add("Text", "x0 y0 w" gW " h1" borderBg)
    gMain.Add("Text", "x0 y" (gH - 1) " w" gW " h1" borderBg)
    gMain.Add("Text", "x0 y0 w1 h" gH borderBg)
    gMain.Add("Text", "x" (gW - 1) " y0 w1 h" gH borderBg)

    bg := " Background" C(BACKGROUND)
    gMain.SetFont(FontOpts(FONT_TIME) " c" C(TEXT_COLOR), FONT_FAMILY)
    gTime := gMain.Add("Text", "x" (PAD + gBorder) " y" (1 + gBorder) " w" gTitleW " h" gTimeRowH " 0x200" bg)
    gMain.SetFont(FontOpts(FONT_COUNTDOWN) " c" C(SECONDARY_TEXT_COLOR), FONT_FAMILY)
    gCountdown := gMain.Add("Text", "x" (PAD + gBorder) " y" (1 + gBorder) " w1 h" gTimeRowH " 0x200" bg)
    gMain.SetFont(FontOpts(FONT_TITLE) " c" C(TEXT_COLOR), FONT_FAMILY)
    gTitle := gMain.Add("Text", "x" (PAD + gBorder) " y" (1 + gBorder + gTimeRowH) " w" gTitleW " h" gTitleH bg)
    gMain.SetFont(FontOpts(FONT_LOCATION) " c" C(TEXT_COLOR), FONT_FAMILY)
    gPlace := gMain.Add("Text", "x" (PAD + gBorder) " y" (1 + gBorder + gTimeRowH + gTitleH) " w" gTitleW " h" gPlaceH bg)

    nextX := PAD + gTitleW + 2 + gBorder
    nextY := gBorder + Max(0, (innerH - 3 * gNextH) // 2)
    gMain.SetFont(FontOpts(FONT_NEXT) " c" C(NEXT_TIME_COLOR), FONT_FAMILY)
    gNextLabel := gMain.Add("Text", "x" nextX " y" nextY " w" gNextW " h" gNextH " Center" bg, "NEXT")
    gNextTime := gMain.Add("Text", "x" nextX " y" (nextY + gNextH) " w" gNextW " h" (2 * gNextH) " Center" bg)

    ; Explicit 1px meeting-to-NEXT and NEXT-to-gear dividers.
    dividerBg := " Background" C(MAIN_BORDER_COLOR)
    meetingDividerX := nextX - 1
    gearX := gBorder + innerW - gGearW
    gMain.Add("Text", "x" meetingDividerX " y" gBorder " w1 h" innerH dividerBg)
    gMain.Add("Text", "x" (gearX - 1) " y" gBorder " w1 h" innerH dividerBg)

    if !SHOW_NEXT_MEETING_TIME {
        gNextLabel.Visible := false
        gNextTime.Visible := false
    }

    gMain.SetFont(FontOpts(FONT_GEAR) " c" C(GEAR_COLOR), GEAR_FONT_FAMILY)
    gGear := gMain.Add("Text", "x" (gBorder + innerW - gGearW) " y" gBorder " w" gGearW " h" innerH " Center 0x200" bg, Chr(0x2699))
}

PositionSection() {
    global INSTANCE_ID
    return "Widget" INSTANCE_ID
}

PlaceInitially() {
    global gPosX, gPosY
    pos := ""
    section := PositionSection()
    if SAVE_POSITION && FileExist(POSITION_FILE) {
        try pos := {x: Integer(IniRead(POSITION_FILE, section, "x")),
                    y: Integer(IniRead(POSITION_FILE, section, "y"))}
    }
    if !IsObject(pos)
        pos := DefaultPosition(gW, gH)
    pos := Clamp(pos.x, pos.y, gW, gH)
    gPosX := pos.x, gPosY := pos.y
    gMain.Show("NA x" gPosX " y" gPosY " w" gW " h" gH)
    KeepOnTop()
}

SavePosition() {
    if SAVE_POSITION {
        section := PositionSection()
        IniWrite(gPosX, POSITION_FILE, section, "x")
        IniWrite(gPosY, POSITION_FILE, section, "y")
    }
}

ResetPosition(*) {
    global gPosX, gPosY
    try IniDelete(POSITION_FILE, PositionSection())
    def := DefaultPosition(gW, gH)
    pos := Clamp(def.x, def.y, gW, gH)
    gPosX := pos.x, gPosY := pos.y
    WinMove(gPosX, gPosY, , , gMain)
}

KeepOnTop() {
    if gPressing
        return
    ; NOSIZE | NOMOVE | NOACTIVATE | NOREDRAW (0x0008) - a bare z-order reassertion
    ; needs no repaint; without NOREDRAW, Windows flashes the Gui's own background
    ; (now the border color) across the whole window on every call, which looks
    ; like blinking.
    flags := 0x0001 | 0x0002 | 0x0010 | 0x0008
    DllCall("SetWindowPos", "Ptr", gMain.Hwnd, "Ptr", -1, "Int", 0, "Int", 0, "Int", 0, "Int", 0, "UInt", flags)
    if gPopupVisible
        DllCall("SetWindowPos", "Ptr", gPopup.Hwnd, "Ptr", -1, "Int", 0, "Int", 0, "Int", 0, "Int", 0, "UInt", flags)
}

RenderMain() {
    global gCurrentMain
    if gError != "" && !gLoaded {
        SetTimeText("OUTLOOK ERROR", "")
        gTitle.Text := Truncate(gError)
        gPlace.Text := "Right-click " Chr(0x2699) " > Open log"
        return
    }
    if !gLoaded
        return
    now := A_Now
    visible := []
    for m in gMainMeetings
        if m.end > now
            visible.Push(m)
    if visible.Length && ShouldPromote(visible[1], now)
        visible.RemoveAt(1)
    gCurrentMain := visible.Length ? visible[1] : ""
    following := visible.Length > 1 ? visible[2] : ""

    if !IsObject(gCurrentMain) {
        SetTimeText("", "")
        gTitle.Text := "No upcoming meetings"
        gPlace.Text := ""
        gNextTime.Text := ""
        return
    }
    SetTimeText(FormatWhen(gCurrentMain.start, now),
                SHOW_COUNTDOWN ? "  " FormatCountdown(gCurrentMain.start, now) : "")
    gTitle.Text := Truncate(DisplayMeetingTitle(gCurrentMain))
    gPlace.Text := Truncate("[" PlaceText(gCurrentMain) "]")
    if SHOW_NEXT_MEETING_TIME
        gNextTime.Text := IsObject(following) ? StrReplace(FormatClock(following.start), " ", "`n") : ""
}


DisplayMeetingTitle(m) {
    if !IsObject(m)
        return ""
    labels := []

    if m.status = "cancelled"
        labels.Push("cancelled")
    else if m.status = "expired"
        labels.Push("expired")
    else if m.status = "all_day"
        labels.Push("all-day")

    if m.response = "not_responded"
        labels.Push("not responded")
    else if m.response = "tentative"
        labels.Push("tentative response")
    else if m.response = "declined"
        labels.Push("declined")

    if m.availability = "free"
        labels.Push("free")
    else if m.availability = "tentative" && m.response != "tentative"
        labels.Push("show as tentative")
    else if m.availability = "oof"
        labels.Push("out of office")
    else if m.availability = "working_elsewhere"
        labels.Push("working elsewhere")

    suffix := labels.Length ? " (" JoinLabels(labels) ")" : ""
    return m.subject suffix
}

JoinLabels(labels) {
    text := ""
    for i, label in labels
        text .= (i > 1 ? ", " : "") label
    return text
}

PlaceText(m) {
    if m.platform != ""
        return m.platform
    return m.location != "" ? m.location : "No location"
}

SetTimeText(timeText, countdownText) {
    gTime.Text := timeText
    tw := Min(gTitleW, TextWidth(gTime, timeText) + 1)
    gTime.Move(, , tw)
    gCountdown.Move(PAD + gBorder + tw, , Max(1, gTitleW - tw))
    gCountdown.Text := countdownText
}

; ============================================================================
; Popup window
;   Popup background = SEPARATOR_COLOR, and every cell is inset by 1px,
;   so the 1px gaps between cells draw the separator lines and the outer border.
; ============================================================================
class HoverBtn {
    __New(g, text, fontOpts, fontName, lineH, lines, action) {
        this.action := action
        this.enabled := true
        this.hover := false
        this.lines := lines
        this.lineH := lineH
        this.visible := true
        this.x := 0, this.y := 0, this.w := 0, this.h := 0
        bg := " Background" C(BACKGROUND)
        g.SetFont(fontOpts " c" C(TEXT_COLOR), fontName)
        this.top := g.Add("Text", "x0 y0 w1 h1" bg)
        this.label := g.Add("Text", "x0 y0 w1 h1 Center" (lines = 1 ? " 0x200" : "") bg, text)
        this.bottom := g.Add("Text", "x0 y0 w1 h1" bg)
    }

    Place(x, y, w, h) {
        this.x := x, this.y := y, this.w := w, this.h := h
        if this.lines = 1 {
            this.top.Move(x, y, w, 0)
            this.bottom.Move(x, y + h, w, 0)
            this.label.Move(x, y, w, h)
        } else {
            textH := Min(h, this.lines * this.lineH)
            topH := (h - textH) // 2
            this.top.Move(x, y, w, topH)
            this.label.Move(x, y + topH, w, textH)
            this.bottom.Move(x, y + topH + textH, w, h - topH - textH)
        }
    }

    SetVisible(v) {
        this.visible := v
        this.top.Visible := v
        this.label.Visible := v
        this.bottom.Visible := v
    }

    SetEnabled(e) {
        e := !!e
        if e != this.enabled {
            this.enabled := e
            this.Paint()
        }
    }

    SetHover(h) {
        if h != this.hover {
            this.hover := h
            this.Paint()
        }
    }

    Paint() {
        if !this.enabled {
            bg := BACKGROUND, fg := DISABLED_TEXT_COLOR
        } else if this.hover {
            bg := HOVER_BACKGROUND, fg := HOVER_TEXT_COLOR
        } else {
            bg := BACKGROUND, fg := TEXT_COLOR
        }
        for ctl in [this.top, this.label, this.bottom] {
            ctl.Opt("Background" C(bg))
            ctl.SetFont("c" C(fg))
            ctl.Redraw()
        }
    }

    Contains(cx, cy) {
        if !this.visible
            return false
        return cx >= this.x && cx < this.x + this.w && cy >= this.y && cy < this.y + this.h
    }

    Click() {
        if this.enabled && this.action
            this.action.Call()
    }
}

BuildPopup() {
    global gPopup, gRows
    global gRowH, gPTimeRowH, gPPlaceRowH, gBottomH, gRowBtnW

    gPopup := Gui("+AlwaysOnTop -Caption +ToolWindow -DPIScale +Owner" gMain.Hwnd)
    gPopup.BackColor := C(BACKGROUND)
    gPopup.MarginX := 0, gPopup.MarginY := 0

    gPTimeRowH := gTimeH + 3
    gPPlaceRowH := gPlaceH + 3
    gRowH := gPTimeRowH + gTitleH + gPPlaceRowH
    gBottomH := 0
    gRowBtnW := 0

    bg := " Background" C(BACKGROUND)
    ; Explicit popup perimeter. Bottom edge comes from the last row separator.
    gPopup.Add("Text", "x0 y0 w" gW " h1 Background" C(SEPARATOR_COLOR))
    gPopup.Add("Text", "x0 y0 w1 h10000 Background" C(SEPARATOR_COLOR))
    gPopup.Add("Text", "x" (gW - 1) " y0 w1 h10000 Background" C(SEPARATOR_COLOR))
    loop POPUP_MEETING_COUNT {
        row := {}
        gPopup.SetFont(FontOpts(FONT_TIME) " c" C(TEXT_COLOR), FONT_FAMILY)
        row.pad := gPopup.Add("Text", "x0 y0 w1 h1" bg)
        row.time := gPopup.Add("Text", "x0 y0 w1 h1 0x200" bg)
        gPopup.SetFont(FontOpts(FONT_TITLE) " c" C(TEXT_COLOR), FONT_FAMILY)
        row.title := gPopup.Add("Text", "x0 y0 w1 h1" bg)
        gPopup.SetFont(FontOpts(FONT_LOCATION) " c" C(TEXT_COLOR), FONT_FAMILY)
        row.place := gPopup.Add("Text", "x0 y0 w1 h1 0x200" bg)
        row.sep := gPopup.Add("Text", "x0 y0 w1 h1 Background" C(SEPARATOR_COLOR))
        row.y := 0
        row.visible := true
        gRows.Push(row)
    }
}


VisiblePopupMeetings() {
    global gPopupMeetings, gCurrentMain
    now := A_Now
    result := []
    for m in gPopupMeetings {
        if m.end <= now
            continue
        if IsObject(gCurrentMain) && m.key = gCurrentMain.key
            continue
        result.Push(m)
        if result.Length >= POPUP_MEETING_COUNT
            break
    }
    return result
}

RenderPopup() {
    global gPopupH
    meetings := VisiblePopupMeetings()
    inner := gW - 2
    y := 1
    for i, row in gRows {
        show := i <= meetings.Length
        row.visible := show
        if !show {
            for ctl in [row.pad, row.time, row.title, row.place, row.sep]
                ctl.Visible := false
            continue
        }
        m := meetings[i]
        row.y := y
        ; Use the original v7 row geometry. The separator is only the popup
        ; background exposed between rows, not an added inset border.
        textW := inner - PAD
        row.pad.Move(1, y, PAD, gRowH)
        row.time.Move(1 + PAD, y, textW, gPTimeRowH)
        row.title.Move(1 + PAD, y + gPTimeRowH, textW, gTitleH)
        row.place.Move(1 + PAD, y + gPTimeRowH + gTitleH, textW, gPPlaceRowH)
        row.sep.Move(0, y + gRowH, gW, 1)
        for ctl in [row.pad, row.time, row.title, row.place, row.sep]
            ctl.Visible := true
        row.time.Text := FormatWhen(m.start, A_Now)
        row.title.Text := Truncate(DisplayMeetingTitle(m))
        row.place.Text := Truncate("[" PlaceText(m) "]")
        PaintRow(i)
        y += gRowH + 1
    }
    gPopupH := Max(1, y)
}



PaintRow(i) {
    row := gRows[i]
    active := (i = gExpanded) || (gHoverKey = "row" i)
    bg := active ? HOVER_BACKGROUND : BACKGROUND
    fg := active ? HOVER_TEXT_COLOR : TEXT_COLOR
    for ctl in [row.pad, row.time, row.title, row.place] {
        ctl.Opt("Background" C(bg))
        ctl.SetFont("c" C(fg))
        ctl.Redraw()
    }
}

PositionPopup() {
    vs := GetVirtualScreen()
    x := gPosX
    if gPosY + gH // 2 < (vs.t + vs.b) // 2
        y := gPosY + gH + 2                 ; main window in top half: open downward
    else
        y := gPosY - gPopupH - 2
    x := Max(vs.l, Min(x, vs.r - gW))
    y := Max(vs.t, Min(y, vs.b - gPopupH))
    gPopup.Show("NA x" x " y" y " w" gW " h" gPopupH)
}

ShowPopup() {
    global gPopupVisible, gOutsideSince
    if VisiblePopupMeetings().Length = 0
        return
    RenderPopup()
    PositionPopup()
    gPopupVisible := true
    gOutsideSince := 0
    KeepOnTop()
    StartHoverTimer()
}

HidePopup() {
    global gPopupVisible, gExpanded
    gPopup.Hide()
    gPopupVisible := false
    gExpanded := 0
    ClearHover()
}



RowMeeting(i) {
    meetings := VisiblePopupMeetings()
    return i <= meetings.Length ? meetings[i] : ""
}



IsDoubleMeetingClick(m) {
    global gLastMeetingClickTick, gLastMeetingClickKey
    if !IsObject(m)
        return false
    nowTick := A_TickCount
    isDouble := (m.key = gLastMeetingClickKey && nowTick - gLastMeetingClickTick <= DllCall("GetDoubleClickTime", "UInt"))
    gLastMeetingClickKey := m.key
    gLastMeetingClickTick := nowTick
    if isDouble {
        gLastMeetingClickKey := ""
        gLastMeetingClickTick := 0
    }
    return isDouble
}

PhysicalLocation(m) {
    if !IsObject(m)
        return ""
    loc := Trim(m.location)
    if loc = ""
        return ""
    loc := RegExReplace(loc, 'i)https?://[^\s<>"`'\[\](){}]+', "")
    loc := RegExReplace(loc, 'i)\b(Microsoft Teams Meeting|Microsoft Teams|Teams Meeting|Zoom Meeting|Zoom|Webex Meeting|Webex)\b', "")
    loc := Trim(loc, " `t`r`n-|,;[]()")
    return loc
}

HandleMeetingDoubleClick(m) {
    if !IsObject(m)
        return
    if gPopupVisible
        HidePopup()
    if m.joinUrl != "" {
        JoinMeeting(m)
        return
    }
    loc := PhysicalLocation(m)
    message := loc != "" ? "Location: " loc : "No location provided."
    choice := MsgBox(message "`n`nOpen the Outlook meeting details?", m.subject, "YesNo Iconi")
    if choice = "Yes"
        OpenDetails(m)
}

EscapeHiddenField(text) {
    return StrReplace(StrReplace(StrReplace(String(text), "`t", " "), "`r", " "), "`n", " ")
}

ShortMeetingId(realKey) {
    ; FNV-1a style 32-bit hash. The real Outlook key never needs to appear in the
    ; user-facing file. Recomputing this value maps the short ID back to a meeting.
    hash := 2166136261
    loop parse String(realKey) {
        hash := ((hash ^ Ord(A_LoopField)) * 16777619) & 0xFFFFFFFF
    }
    return Format("M-{:08X}", hash)
}

FriendlyHiddenLine(shortId, startTime, subject) {
    when := FormatTime(startTime, "yyyy-MM-dd h:mm tt")
    return shortId " | " when " | " EscapeHiddenField(subject)
}


LoadHiddenMeetingKeys() {
    global HIDDEN_MEETINGS_FILE
    keys := Map(), keep := []
    if !FileExist(HIDDEN_MEETINGS_FILE)
        return keys

    for line in StrSplit(FileRead(HIDDEN_MEETINGS_FILE, "UTF-8"), "`n", "`r") {
        text := Trim(line)
        if text = "" || SubStr(text, 1, 1) = "#"
            continue

        ; Current friendly format:
        ; M-12AB34CD | 2026-11-06 8:30 AM | Meeting name
        if RegExMatch(text, "^M-[0-9A-Fa-f]{8}(?=\\s*\\|)", &match) {
            id := StrUpper(match[0])
            keys[id] := true
            keep.Push(text)
            continue
        }

        ; One-time migration support for the old tab-delimited format. Convert it
        ; to the friendly format automatically without exposing the long Outlook ID.
        parts := StrSplit(line, "`t")
        if parts.Length >= 4 {
            if parts[2] <= A_Now
                continue
            id := ShortMeetingId(parts[1])
            keys[id] := true
            keep.Push(FriendlyHiddenLine(id, parts[3], parts[4]))
        }
    }
    WriteHiddenMeetingFile(keep)
    return keys
}

WriteHiddenMeetingFile(lines) {
    global HIDDEN_MEETINGS_FILE
    header := "# User-Hidden Meetings`r`n# Delete a meeting line, save this file, then from the Gear icon, select [Refresh meetings] to show it again.`r`n"
    body := ""
    for line in lines
        body .= line "`r`n"
    try FileDelete(HIDDEN_MEETINGS_FILE)
    FileAppend(header body, HIDDEN_MEETINGS_FILE, "UTF-8")
}

HideMeeting(m, *) {
    global HIDDEN_MEETINGS_FILE
    if !IsObject(m)
        return
    if !FileExist(HIDDEN_MEETINGS_FILE)
        WriteHiddenMeetingFile([])
    keys := LoadHiddenMeetingKeys()
    id := ShortMeetingId(m.key)
    if !keys.Has(id) {
        line := FriendlyHiddenLine(id, m.start, m.subject)
        FileAppend(line "`r`n", HIDDEN_MEETINGS_FILE, "UTF-8")
    }
    RefreshCalendar("meeting hidden by user")
}

ApplyUserHidden(meetings) {
    keys := LoadHiddenMeetingKeys()
    if keys.Count = 0
        return meetings
    visible := []
    for m in meetings {
        if !keys.Has(ShortMeetingId(m.key))
            visible.Push(m)
    }
    return visible
}

OpenHiddenMeetings(*) {
    global HIDDEN_MEETINGS_FILE
    if !FileExist(HIDDEN_MEETINGS_FILE)
        WriteHiddenMeetingFile([])
    Run('notepad.exe "' HIDDEN_MEETINGS_FILE '"')
}

ShowMeetingMenu(m) {
    if !IsObject(m)
        return
    meetingMenu := Menu()
    meetingMenu.Add("Meeting details", (*) => OpenDetails(m))
    if m.joinUrl != ""
        meetingMenu.Add("Join meeting", (*) => JoinMeeting(m))
    loc := PhysicalLocation(m)
    if loc != "" {
        label := "Location: " TruncateMenuText(loc, 80)
        meetingMenu.Add(label, (*) => 0)
        meetingMenu.Disable(label)
    }
	meetingMenu.Add("Hide this meeting", HideMeeting.Bind(m))
    meetingMenu.Show()
}

TruncateMenuText(text, maxChars) {
    text := StrReplace(text, "&", "&&")
    return StrLen(text) <= maxChars ? text : SubStr(text, 1, maxChars - 3) "..."
}

JoinMeeting(m) {
    if IsObject(m) && m.joinUrl != ""
        try Run(m.joinUrl)
}

OpenDetails(m) {
    if !IsObject(m) || m.entryId = ""
        return
    try {
        ConnectOutlook().GetItemFromID(m.entryId).Display(false)
    } catch as e {
        Log("Could not open meeting: " e.Message)
    }
}

; ---------- popup hit-testing / hover ----------
PopupHit(cx, cy) {
    for i, row in gRows {
        if row.visible && cy >= row.y && cy < row.y + gRowH
            return {key: "row" i, row: i}
    }
    return ""
}

ApplyHover(hit) {
    global gHoverKey
    key := IsObject(hit) ? hit.key : ""
    if key = gHoverKey
        return
    old := gHoverKey
    gHoverKey := key
    if RegExMatch(old, "^row(\d+)$", &m)
        PaintRow(Integer(m[1]))
    if IsObject(hit) && hit.HasOwnProp("row")
        PaintRow(hit.row)
}

ClearHover() {
    global gHoverKey
    gHoverKey := ""
    for i, row in gRows
        if row.visible
            PaintRow(i)
}

AllButtons() {
    return []
}

StartHoverTimer() {
    global gHoverTimerOn
    if !gHoverTimerOn {
        gHoverTimerOn := true
        SetTimer(HoverTick, 60)
    }
}

HoverTick() {
    global gHoverStart, gOutsideSince, gHoverTimerOn
    CoordMode("Mouse", "Screen")
    MouseGetPos(&mx, &my, &win)
    overMain := (win = gMain.Hwnd)
    overPopup := gPopupVisible && (win = gPopup.Hwnd)
    now := A_TickCount

    if overMain && !gPressing && !gPopupVisible && POPUP_OPEN_ON_HOVER {
        if !gHoverStart
            gHoverStart := now
        else if now - gHoverStart >= POPUP_HOVER_DELAY_MS
            ShowPopup()
    } else if !overMain {
        gHoverStart := 0
    }

    if overMain || overPopup {
        gOutsideSince := 0
    } else if gPopupVisible {
        if !gOutsideSince
            gOutsideSince := now
        else if now - gOutsideSince >= POPUP_CLOSE_DELAY_MS
            HidePopup()
    }

    if overPopup {
        WinGetPos(&px, &py, , , gPopup)
        ApplyHover(PopupHit(mx - px, my - py))
    } else if gHoverKey != "" {
        ApplyHover("")
    }

    if !overMain && !overPopup && !gPopupVisible {
        SetTimer(HoverTick, 0)
        gHoverTimerOn := false
        gHoverStart := 0
    }
}

; ============================================================================
; Mouse messages (Text controls are click-through, so the Gui gets these)
; ============================================================================
OwnerGui(hwnd) {
    root := DllCall("GetAncestor", "Ptr", hwnd, "UInt", 2, "Ptr")   ; GA_ROOT
    if root = gMain.Hwnd
        return "main"
    if root = gPopup.Hwnd
        return "popup"
    return ""
}

OnMouseMove(wParam, lParam, msg, hwnd) {
    if OwnerGui(hwnd) != ""
        StartHoverTimer()
}

InGear(sx, sy) {
    return sx >= gPosX + gW - gBorder - gGearW && sx < gPosX + gW - gBorder
        && sy >= gPosY + gBorder && sy < gPosY + gH - gBorder
}

OnLButtonDown(wParam, lParam, msg, hwnd) {
    global gPressing, gPosX, gPosY
    which := OwnerGui(hwnd)
    if which = ""
        return
    CoordMode("Mouse", "Screen")
    MouseGetPos(&sx, &sy)
    if which = "popup" {
        WinGetPos(&px, &py, , , gPopup)
        hit := PopupHit(sx - px, sy - py)
        if IsObject(hit) && hit.HasOwnProp("row") {
            m := RowMeeting(hit.row)
            if IsDoubleMeetingClick(m)
                HandleMeetingDoubleClick(m)
        }
        return 0
    }
    if InGear(sx, sy) {
        ShowGearMenu()
        return 0
    }
    gPressing := true
    startX := gPosX, startY := gPosY, moved := false
    while GetKeyState("LButton", "P") {
        MouseGetPos(&x, &y)
        if !moved && (Abs(x - sx) >= DRAG_THRESHOLD_PX || Abs(y - sy) >= DRAG_THRESHOLD_PX) {
            moved := true
            if gPopupVisible
                HidePopup()
        }
        if moved && ALLOW_DRAG {
            pos := Clamp(startX + x - sx, startY + y - sy, gW, gH)
            gPosX := pos.x, gPosY := pos.y
            WinMove(gPosX, gPosY, , , gMain)
        }
        Sleep(10)
    }
    gPressing := false
    if moved && ALLOW_DRAG {
        pos := Snap(gPosX, gPosY, gW, gH)
        gPosX := pos.x, gPosY := pos.y
        WinMove(gPosX, gPosY, , , gMain)
        SavePosition()
    } else if !moved {
        if IsDoubleMeetingClick(gCurrentMain)
            HandleMeetingDoubleClick(gCurrentMain)
        else if POPUP_OPEN_ON_CLICK
            ShowPopup()
    }
    return 0
}

OnRButtonUp(wParam, lParam, msg, hwnd) {
    which := OwnerGui(hwnd)
    if which = ""
        return
    CoordMode("Mouse", "Screen")
    MouseGetPos(&sx, &sy)
    if which = "main" {
        if InGear(sx, sy)
            ShowGearMenu()
        else
            ShowMeetingMenu(gCurrentMain)
        return 0
    }
    WinGetPos(&px, &py, , , gPopup)
    hit := PopupHit(sx - px, sy - py)
    if IsObject(hit) && hit.HasOwnProp("row")
        ShowMeetingMenu(RowMeeting(hit.row))
    return 0
}

; ============================================================================
; Gear menu
; ============================================================================
BuildMenu() {
    global gMenu
    gMenu := Menu()
    gMenu.Add("Add another widget", LaunchWidgetClone)
    gMenu.Add("Open user-hidden meetings", OpenHiddenMeetings)
    gMenu.Add("Refresh meetings", ManualRefresh)
    gMenu.Add("Restart this widget", (*) => Reload())
    gMenu.Add("Close this widget", CloseThisWidget)
    gMenu.Add()
    gMenu.Add("Settings...", OpenSettingsWindow)
    gMenu.Add()
    gMenu.Add("Exit app", (*) => ExitApp())
}

ShowGearMenu() {
    if gPopupVisible
        HidePopup()
    gMenu.Show()
}

OpenLog(*) {
    if FileExist(LOG_FILE)
        Run('notepad.exe "' LOG_FILE '"')
}

ToggleStartup(*) {
    global STARTUP_SHORTCUT
    if FileExist(STARTUP_SHORTCUT) {
        FileDelete(STARTUP_SHORTCUT)
        Log("Start with Windows disabled")
    } else {
        FileCreateShortcut(A_ScriptFullPath, STARTUP_SHORTCUT, A_ScriptDir, , "Meeting Widget")
        Log("Start with Windows enabled: " STARTUP_SHORTCUT)
    }
}

OpenSettings(*) {
    global INSTANCE_ID
    text := "Meeting Widget - Clone " INSTANCE_ID "`n`n"
        . "Double-click a meeting to join it. If there is no virtual link, the widget shows the physical location and offers to open Outlook details.`n`n"
        . "Right-click a meeting for Meeting details, Join meeting, Meeting physical location, and to Hide the meeting.`n`n"
        . "Hidden meetings are stored in User-Hidden Meetings.txt using a short ID, date/time, and meeting name. Delete a meeting line, save the file, then from the Gear icon, select [Refresh meetings] to show it again.`n`n"
        . "Use the gear for widget settings, clones, refresh, logs, and startup options."
    MsgBox(text, "Meeting Widget", "Iconi")
}

; ============================================================================
; Outlook
; ============================================================================
ConnectOutlook() {
    global gNs, gOutlook
    if IsObject(gNs)
        return gNs
    Log("Connecting to Outlook...")
    try {
        gOutlook := ComObjActive("Outlook.Application")
    } catch {
        gOutlook := ComObject("Outlook.Application")
    }
    gNs := gOutlook.GetNamespace("MAPI")
    return gNs
}

DropOutlook() {
    global gNs, gOutlook, gWatchItems, gEventsOn
    if IsObject(gWatchItems)
        try ComObjConnect(gWatchItems)
    gWatchItems := "", gEventsOn := false, gNs := "", gOutlook := ""
}

Subscribe() {
    global gWatchItems, gEventsOn
    if !USE_OUTLOOK_EVENTS || gEventsOn
        return
    try {
        ; Keep a reference to this exact Items object or events stop firing.
        gWatchItems := ConnectOutlook().GetDefaultFolder(9).Items
        ComObjConnect(gWatchItems, CalendarEvents)
        gEventsOn := true
        Log("Listening for Outlook calendar changes (safety re-check every "
            SAFETY_REFRESH_MINUTES " min)")
    } catch as e {
        gWatchItems := "", gEventsOn := false
        Log("Outlook event hookup failed; polling every " FALLBACK_POLL_SECONDS "s: " e.Message)
    }
}

class CalendarEvents {
    static ItemAdd(params*) {
        OnOutlookChange("added")
    }
    static ItemChange(params*) {
        OnOutlookChange("changed")
    }
    static ItemRemove(params*) {
        OnOutlookChange("removed")
    }
}

OnOutlookChange(what) {
    global gChangePending
    if !gChangePending
        Log("Outlook reported a calendar item " what "; re-reading shortly")
    gChangePending := true
    SetTimer(ChangeRefresh, -CHANGE_DEBOUNCE_SECONDS * 1000)   ; re-arms on every event
}

ChangeRefresh() {
    global gChangePending
    gChangePending := false
    RefreshCalendar("Outlook change")
}

OnPowerBroadcast(wParam, lParam, msg, hwnd) {
    if wParam = 7 || wParam = 18                 ; PBT_APMRESUMESUSPEND / RESUMEAUTOMATIC
        SetTimer(WakeRefresh, -5000)
}

RefreshCalendar(reason := "safety check") {
    global gMainMeetings, gPopupMeetings, gLoaded, gError, gLastSig, gRefreshing
    if gRefreshing
        return
    gRefreshing := true
    try {
        Subscribe()
        t0 := A_TickCount
        mainAll := FetchMeetings(MAIN_OUTLOOK_RESTRICT_FILTER)
        if Trim(MAIN_OUTLOOK_RESTRICT_FILTER) = Trim(POPUP_OUTLOOK_RESTRICT_FILTER)
            popupAll := mainAll
        else
            popupAll := FetchMeetings(POPUP_OUTLOOK_RESTRICT_FILTER)
        sig := Signature(mainAll) "||" Signature(popupAll)
        if sig != gLastSig {
            gLastSig := sig
            Log("Calendar updated (" reason "): " mainAll.Length " meetings, "
                Round((A_TickCount - t0) / 1000, 1) "s")
        }
        gMainMeetings := ApplyUserHidden(FilterMeetings(mainAll, MAIN_HIDDEN_STATUSES, MAIN_HIDDEN_RESPONSES, MAIN_HIDDEN_AVAILABILITY))
        gPopupMeetings := ApplyUserHidden(FilterMeetings(popupAll, POPUP_HIDDEN_STATUSES, POPUP_HIDDEN_RESPONSES, POPUP_HIDDEN_AVAILABILITY))
        gLoaded := true
        gError := ""
        interval := gEventsOn ? SAFETY_REFRESH_MINUTES * 60000 : FALLBACK_POLL_SECONDS * 1000
    } catch as e {
        Log("Outlook error (retrying in " RETRY_AFTER_ERROR_SECONDS "s): " e.Message " | " e.What)
        DropOutlook()
        gError := e.Message
        interval := RETRY_AFTER_ERROR_SECONDS * 1000
    }
    gRefreshing := false
    SetTimer(SafetyRefresh, interval)
    Tick()
}

IsNothing(v) {
    if !IsObject(v)
        return true
    try return ComObjValue(v) = 0
    return false
}

FetchMeetings(extra) {
    items := ConnectOutlook().GetDefaultFolder(9).Items
    items.Sort("[Start]")
    items.IncludeRecurrences := true
    now := A_Now
    searchStart := DateAdd(now, -1, "Days")
    searchEnd := DateAdd(now, LOOKAHEAD_DAYS, "Days")
    filter := "[Start] <= '" FormatTime(searchEnd, "MM/dd/yyyy hh:mm tt") "' AND [End] >= '"
            . FormatTime(searchStart, "MM/dd/yyyy hh:mm tt") "'"
    if Trim(extra) != ""
        filter := "(" filter ") AND (" extra ")"
    restricted := items.Restrict(filter)

    meetings := [], seen := Map(), scanned := 0
    item := restricted.GetFirst()
    while !IsNothing(item) && scanned < MAX_MEETINGS_TO_SCAN {
        scanned++
        stop := false
        try {
            if item.Class = 26 {                                   ; olAppointment
                st := ComDateToTs(item.Start)
                if st > searchEnd {
                    stop := true
                } else {
                    en := ComDateToTs(item.End)
                    status := MeetingStatus(item, now, st, en)
                    response := MeetingResponse(item)
                    availability := MeetingAvailability(item)
                    subject := CleanText(item.Subject)
                    if subject = ""
                        subject := "(No title)"
                    entryId := String(item.EntryID)
                    globalId := ""
                    try globalId := String(item.GlobalAppointmentID)
                    key := (globalId != "" ? globalId : entryId) "|" st "|" en
                    if !seen.Has(key) {
                        seen[key] := true
                        location := CleanText(item.Location)
                        info := JoinInfo(item, key, location, status)
                        meetings.Push({subject: subject, start: st, end: en, location: location,
                                       platform: info.platform, joinUrl: info.url,
                                       status: status, response: response, availability: availability,
                                       entryId: entryId, key: key})
                    }
                }
            }
        } catch as e {
            Log("Skipped a calendar item: " e.Message)
        }
        if stop
            break
        item := restricted.GetNext()
    }
    return meetings
}

; Converts an Outlook COM date into an AHK timestamp (YYYYMMDDHH24MISS).
ComDateToTs(v) {
    d := ""
    if IsObject(v) {
        raw := ComObjValue(v)
        if IsFloat(raw) {
            d := raw
        } else {
            buf := Buffer(8)
            NumPut("Int64", raw, buf)
            d := NumGet(buf, "Double")
        }
    } else if RegExMatch(String(v), "^\d{14}$") {
        return String(v)
    } else if IsNumber(v) {
        d := v + 0.0
    }
    if d = "" || d < 1 || d > 2958465 {               ; not a plausible OLE date: parse as text
        d := 0.0
        if DllCall("oleaut32\VarDateFromStr", "WStr", String(v), "UInt", 0x400, "UInt", 0, "Double*", &d) != 0
            throw Error("Unrecognized date: " String(v))
    }
    days := Floor(d)
    secs := Round((d - days) * 86400)
    return DateAdd(DateAdd("18991230000000", days, "Days"), secs, "Seconds")
}

MeetingStatus(item, now, st, en) {
    if en <= now
        return "expired"
    if item.AllDayEvent
        return "all_day"
    ms := item.MeetingStatus
    if ms = 5 || ms = 7
        return "cancelled"
    return (st <= now && now < en) ? "active" : "scheduled"
}

MeetingResponse(item) {
    ; Outlook OlResponseStatus:
    ; 0 none, 1 organizer, 2 tentative, 3 accepted, 4 declined, 5 not responded.
    try response := Integer(item.ResponseStatus)
    catch
        return "unknown"
    switch response {
        case 0: return "none"
        case 1: return "organizer"
        case 2: return "tentative"
        case 3: return "accepted"
        case 4: return "declined"
        case 5: return "not_responded"
        default: return "unknown"
    }
}

MeetingAvailability(item) {
    ; Outlook OlBusyStatus:
    ; 0 free, 1 tentative, 2 busy, 3 out of office, 4 working elsewhere.
    try busy := Integer(item.BusyStatus)
    catch
        return "unknown"
    switch busy {
        case 0: return "free"
        case 1: return "tentative"
        case 2: return "busy"
        case 3: return "oof"
        case 4: return "working_elsewhere"
        default: return "unknown"
    }
}

JoinInfo(item, key, location, status) {
    global gLinkCache
    if gLinkCache.Has(key)
        return gLinkCache[key]
    url := PickJoinUrl(ExtractUrls(location), location)
    platform := ClassifyPlatform(url, location)
    if url = "" && status != "expired" && status != "cancelled" && status != "declined" {
        body := ""
        try body := String(item.Body)
        both := location " " body
        url := PickJoinUrl(ExtractUrls(both), both)
        platform := ClassifyPlatform(url, both)
    }
    return gLinkCache[key] := {url: url, platform: platform}
}

ExtractUrls(text) {
    urls := [], pos := 1
    while pos := RegExMatch(text, 'i)https?://[^\s<>"`'\[\](){}]+', &m, pos) {
        u := RTrim(m[0], ".,;:!?)}]>'`"")
        if !HasValue(urls, u)
            urls.Push(u)
        pos += m.Len
    }
    return urls
}

ClassifyPlatform(url, text := "") {
    if url != "" && RegExMatch(url, "i)^https?://([^/:?#]+)[^/?#]*([^?#]*)", &m) {
        host := StrLower(m[1]), path := StrLower(m[2])
        if RegExMatch(host, "(^|\.)(zoom\.us|zoomgov\.com|zoom\.com)$") && RegExMatch(path, "^/(j|w|my|s)/")
            return "Zoom"
        if (host = "teams.microsoft.com" || host = "teams.live.com") && RegExMatch(path, "^/(l/meetup-join|meet)/")
            return "Microsoft Teams"
        if RegExMatch(host, "(^|\.)webex\.com$") && RegExMatch(path, "/meet/|/join/|/wbxmjs/joinservice/")
            return "Webex"
    }
    lowered := StrLower(text)
    if InStr(lowered, "zoom")
        return "Zoom"
    if InStr(lowered, "microsoft teams") || InStr(lowered, "teams meeting")
        return "Microsoft Teams"
    if InStr(lowered, "webex")
        return "Webex"
    return url != "" ? "Online" : ""
}

PickJoinUrl(urls, text) {
    for wanted in ["Zoom", "Microsoft Teams", "Webex"]
        for u in urls
            if ClassifyPlatform(u) = wanted
                return u
    lowered := StrLower(text)
    for marker in ["join meeting", "online meeting", "zoom meeting", "microsoft teams meeting", "webex meeting"] {
        if InStr(lowered, marker) {
            for u in urls {
                lu := StrLower(u)
                if !InStr(lu, "support.microsoft.com") && !InStr(lu, "zoom.us/download") && !InStr(lu, "zoom.us/privacy")
                    return u
            }
            break
        }
    }
    return ""
}

FilterMeetings(meetings, hiddenStatuses, hiddenResponses, hiddenAvailability) {
    out := []
    for m in meetings {
        if HasValue(hiddenStatuses, m.status)
            continue
        if HasValue(hiddenResponses, m.response)
            continue
        if HasValue(hiddenAvailability, m.availability)
            continue
        out.Push(m)
    }
    return out
}

Signature(meetings) {
    s := ""
    for m in meetings
        s .= m.subject "|" m.start "|" m.end "|" m.status "|" m.response "|" m.availability "|" m.location "|" m.joinUrl "`n"
    return s
}

; ============================================================================
; Time formatting + minute tick
; ============================================================================
FormatClock(ts) => StrLower(FormatTime(ts, "h:mm tt"))

FormatWhen(ts, now) {
    day := SubStr(ts, 1, 8)
    if day = SubStr(now, 1, 8)
        return FormatClock(ts)
    if day = FormatTime(DateAdd(now, 1, "Days"), "yyyyMMdd")
        return "Tomorrow " FormatClock(ts)
    return FormatTime(ts, "ddd") " " FormatClock(ts)
}

FormatCountdown(ts, now) {
    seconds := DateDiff(ts, now, "Seconds")
    if seconds <= 0
        return "now"
    minutes := Max(1, (seconds + 59) // 60)
    if minutes < 60
        return minutes = 1 ? "in 1 min" : "in " minutes " mins"
    hours := (minutes + 30) // 60
    if hours < 24
        return hours = 1 ? "in 1 hr" : "in " hours " hrs"
    days := Max(1, (hours + 12) // 24)
    return days = 1 ? "in 1 day" : "in " days " days"
}

ShouldPromote(m, now) {
    if !(m.start <= now && now < m.end)
        return false
    return DateDiff(now, m.start, "Minutes") >= PROMOTE_NEXT_AFTER_ACTIVE_MINUTES
}

Tick() {
    global gLastDay
    today := FormatTime(A_Now, "yyyyMMdd")
    if today != gLastDay {
        gLastDay := today
        SetTimer(NewDayRefresh, -10)
    }
    RenderMain()
    if gPopupVisible {
        RenderPopup()
        PositionPopup()
    }
    ms := (60 - A_Sec) * 1000 - A_MSec + 50
    SetTimer(Tick, -ms)
}
