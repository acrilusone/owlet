# Meeting Widget

A compact Windows desktop widget that displays upcoming Outlook calendar meetings and keeps them easily accessible near the Windows taskbar.

The widget is written in AutoHotkey v2 and uses the Classic Outlook desktop application's COM interface to read the user's calendar.

## Features

- Shows your next Outlook meeting
- Displays meeting start time and countdown
- Shows the following meeting time at a glance
- Hover or click to see additional upcoming meetings
- Double-click a meeting to join it
- Supports Microsoft Teams, Zoom, Webex, and other online meeting links
- Displays physical meeting locations when available
- Right-click access to meeting details and actions
- Hide individual meetings from the widget
- Filter meetings by RSVP status, availability, and meeting status
- Customizable colors, fonts, layout, and behavior
- Multiple independent widget instances
- Remembers widget positions
- Optional Start with Windows support
- Automatically refreshes when Outlook calendar items change

## Requirements

- Windows
- Classic Outlook for Windows installed
- An Outlook profile configured in Classic Outlook
- AutoHotkey v2 if running the `.ahk` source directly

> **Important:** New Outlook for Windows does not provide the COM automation interface used by Meeting Widget. Classic Outlook must be installed. You can continue using New Outlook normally if Classic Outlook is also installed.

## Installation

### Using the compiled application

1. Download the latest Meeting Widget zipped file and extract it.
3. Keep the application and `Meeting Widget Preferences.txt` together in the same folder.
4. Run `Meeting Widget.exe`.
5. The widget will connect to your Outlook calendar.

Classic Outlook does not need to already be open. If an Outlook COM session is not running, Meeting Widget will attempt to start one automatically.

### Running the AutoHotkey source

1. Install [AutoHotkey v2](https://www.autohotkey.com/).
2. Download the repository.
3. Double-click the Meeting Widget `.ahk` file.

## Start with Windows

Open:

**Gear > Settings > General**

Enable **Start with Windows**.

When enabled, Meeting Widget creates a shortcut in the current user's Windows Startup folder.

You can also use **Open Startup Folder** from Settings to view that folder directly.

## Using the Widget

### Main widget

- **Hover** over the widget to display upcoming meetings.
- **Click** the widget to display upcoming meetings.
- **Double-click** a meeting to join it.
- **Right-click** a meeting for additional options.
- **Drag** the widget to reposition it.
- Click the **gear icon** for settings and application controls.

### Meeting menu

Right-click a meeting to access options such as:

- Meeting details
- Join meeting
- Physical location
- Hide this meeting

### Hidden meetings

Meetings hidden manually are recorded in:

`User-Hidden Meetings.txt`

To restore one:

1. Open **Gear > Open user-hidden meetings**.
2. Delete the line for the meeting.
3. Save the file.
4. Select **Gear > Refresh meetings**.

## Multiple Widgets

Select:

**Gear > Add another widget**

Each widget can be positioned independently. Meeting Widget remembers the desired number of widgets and restores them the next time the application starts.

## Settings

Open:

**Gear > Settings**

Settings include:

- Meeting visibility
- RSVP filtering
- Availability filtering
- Fonts
- Colors
- Borders
- Refresh behavior
- Logging
- Main window layout
- Popup behavior
- Start with Windows

Most configuration is stored in:

`Meeting Widget Preferences.txt`

Use **Save and Restart** after changing settings that affect appearance, layout, fonts, or meeting visibility.

## Troubleshooting

### The widget says Outlook Error

Confirm that:

1. Classic Outlook for Windows is installed.
2. Your Outlook account/profile has been configured in Classic Outlook.
3. You can open your calendar successfully in Classic Outlook.

Then select:

**Gear > Refresh meetings**

If necessary, restart Meeting Widget.

### Meetings are missing

Check the meeting's:

- RSVP response
- Show As status
- Meeting status

The default filters can hide declined, not-responded, free, cancelled, expired, and all-day calendar items depending on the configured view.

Review these under:

**Settings > Meeting Visibility**

Also check `User-Hidden Meetings.txt` if you previously hid the meeting manually.

### New Outlook

Meeting Widget currently uses the Outlook COM Object Model. New Outlook for Windows does not support this automation interface.

Classic Outlook therefore needs to remain installed for Meeting Widget to access the calendar.

## Privacy

Meeting Widget reads Outlook calendar information locally through the installed Classic Outlook application.

User-specific runtime files such as saved widget positions, hidden-meeting lists, preferences, and logs should not be committed to the GitHub repository.

## License

This project is available under the MIT License. See `LICENSE` for details.
