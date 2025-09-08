# NetNotifier.ahk

`NetNotifier.ahk` is an AutoHotkey v2.0 script designed to monitor your internet connection status and provide visual and optional audio alerts.

## Features

- **Real-time Connectivity Monitoring**: Continuously checks internet availability by pinging google.com (additional targets can be configured).
- **Visual Status Indicators**: Changes the system tray icon to reflect connection status:
    - `green.ico`: Internet is online.
    - `red.ico`: Internet is offline or experiencing issues.
- **Voice Alerts**: Optional voice notifications for connection status changes (e.g., "Connection Restored", "Connection Lost").
- **Configurable Settings**: A user-friendly GUI allows you to customize:
    - **Check Interval**: How frequently the script checks for internet connectivity.
     - **Ping Targets**: Comma-separated list of URLs/IPs for pinging (default: google.com).
    - **Voice Alerts**: Enable or disable spoken notifications.
    Settings are saved to `Settings.ini`.
- **Connection Statistics**: Tracks and displays useful metrics in the tray icon tooltip, including:
    - Total online time.
    - Number of disconnects.
    - Overall connection availability percentage.
    - Your local IP address.
- **Robust Connectivity Checks**: Utilizes advanced methods (WMI, DLL calls) to accurately determine both internet and local network (gateway) status.
- **Non-blocking Operations**: Employs asynchronous pinging to ensure the script remains responsive and does not freeze the user interface.

## Usage

1. Run `NetNotifier.ahk` (or `NetNotifier.exe` if compiled).
2. The script will appear in your system tray.
3. Right-click the tray icon to access settings, check connection manually, or reset statistics.

## Requirements

- AutoHotkey v2.0 (if running the `.ahk` script directly)

## Files

- `NetNotifier.ahk`: The main AutoHotkey script.
- `NetNotifier.exe`: (Optional) Compiled executable version of the script.
- `Settings.ini`: Configuration file for user preferences.
- `green.ico`: Icon for online status.
- `red.ico`: Icon for offline/issues status.
- `issues.ico`: (Currently unused, `red.ico` is used for issues) Icon for local network but no internet status.
