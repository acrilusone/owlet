\# Meeting Widget



A compact Windows desktop widget that displays upcoming Outlook calendar meetings and keeps them easily accessible near the Windows taskbar.



The widget is written in AutoHotkey v2 and uses the Classic Outlook desktop application's COM interface to read the user's calendar.



\## Features



\- Shows your next Outlook meeting

\- Displays meeting start time and countdown

\- Shows the following meeting time at a glance

\- Hover or click to see additional upcoming meetings

\- Double-click a meeting to join it

\- Supports Microsoft Teams, Zoom, Webex, and other online meeting links

\- Displays physical meeting locations when available

\- Right-click access to meeting details and actions

\- Hide individual meetings from the widget

\- Filter meetings by RSVP status, availability, and meeting status

\- Customizable colors, fonts, layout, and behavior

\- Multiple independent widget instances

\- Remembers widget positions

\- Optional Start with Windows support

\- Automatically refreshes when Outlook calendar items change



\## Requirements



\- Windows

\- Classic Outlook for Windows installed

\- An Outlook profile configured in Classic Outlook

\- AutoHotkey v2 if running the `.ahk` source directly



> \*\*Important:\*\* New Outlook for Windows does not provide the COM automation interface used by Meeting Widget. Classic Outlook must be installed. You can continue using New Outlook normally if Classic Outlook is also installed.



\## Installation



\### Using the compiled application



1\. Download the latest Meeting Widget files.

2\. Keep the application and `Meeting Widget Preferences.txt` together in the same folder.

3\. Run `Meeting Widget.exe`.

4\. The widget will connect to your Outlook calendar.



Classic Outlook does not need to already be open. If an Outlook COM session is not running, Meeting Widget will attempt to start one automatically.



\### Running the AutoHotkey source



1\. Install \[AutoHotkey v2](https://www.autohotkey.com/).

2\. Download the repository.

3\. Double-click the Meeting Widget `.ahk` file.



\## Start with Windows



Open:



\*\*Gear > Settings > General\*\*



Enable \*\*Start with Windows\*\*.



When enabled, Meeting Widget creates a shortcut in the current user's Windows Startup folder.



You can also use \*\*Open Startup Folder\*\* from Settings to view that folder directly.



\## Using the Widget



\### Main widget



\- \*\*Hover\*\* over the widget to display upcoming meetings.

\- \*\*Click\*\* the widget to display upcoming meetings.

\- \*\*Double-click\*\* a 

