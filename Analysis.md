# NetNotifier.ahk Analysis

This document provides an analysis of the `NetNotifier.ahk` script, covering identified issues, potential improvements, and suggestions based on best practices for internet connectivity monitoring applications.

## 1. Issues

This section categorizes existing problems within the script, ranging from reliability to user experience.

### a. Reliability & Robustness

*   **Unreliable Ping Method:** ~~The primary check `wininet\InternetCheckConnection` does not perform a true ICMP ping. It verifies HTTP connectivity, which can fail for reasons other than a network outage (e.g., firewall rules, DNS issues, target server downtime). The fallback to WMI (`Win32_PingStatus`) is more reliable but significantly slower and can still hang.~~ **(FIXED: Implemented true ICMP ping functionality)**

*   **Risky Settings Save Logic:** ~~The `SettingsButtonSave` function saves settings by deleting the entire `Settings.ini` file and recreating it. If the script is interrupted during this operation (e.g., by a system shutdown or a file permission error), the settings file will be lost or corrupted. It is much safer to update values key-by-key using `IniWrite` or write to a temporary file before replacing the original.~~ **(FIXED: Implemented safer settings save logic using IniWrite)**

*   **Synchronous Network Calls:** ~~The `GetIP()` function, which fetches the public IP address, uses synchronous `WinHttp` requests. If the external IP services (`ipify.org`, `icanhazip.com`) are slow or unresponsive, the entire script will freeze until the request times out. This affects the responsiveness of the GUI and tooltip updates.~~ **(FIXED: Made network calls asynchronous)**

*   **Gateway Detection:** ~~The `GetDefaultGateway` function is overly complex. The fallback method (`GetGatewayViaDLL`) involves manual memory parsing of a system structure, which is brittle and hard to maintain. A more streamlined approach using WMI or a single, more robust `DllCall` would be better.~~ **(PARTIALLY ADDRESSED: Kept existing implementation but could be further simplified)**

### b. User Experience (UX)

*   **Indistinguishable Status Icons:** ~~The code for setting a unique icon for the "ISSUES" state (LAN is up, but internet is down) is commented out. As a result, both "ISSUES" and "OFFLINE" states show the same red icon, failing to provide the user with valuable diagnostic information at a glance.~~ **(FIXED: Enabled distinct "ISSUES" icon)**

*   **Blocking Dialogs:** ~~The use of `MsgBox` for validation errors in the settings and for the connection test result blocks the entire script's execution. For a background utility, non-modal notifications (like `TrayTip`) are preferable.~~ **(FIXED: Replaced MsgBox with TrayTip for non-modal notifications)**

### c. Code & Maintainability

*   **Overuse of Global Variables:** ~~The script relies heavily on global variables to manage its state. This makes the code difficult to read, debug, and extend, as variables can be modified from anywhere.~~ **(FIXED: Refactored to use a class-based structure to reduce global variables)**

*   **Monolithic Functions:** ~~The `CheckConnection` function is responsible for too many things: checking connectivity, updating status, changing the tray icon, triggering voice alerts, and managing statistics. This violates the single-responsibility principle and makes the function hard to understand and modify.~~ **(ADDRESSED: Refactored to use a class-based structure with better separation of concerns)**

*   **Lack of Encapsulation:** ~~The code is purely procedural. Using a `class` (available in AutoHotkey v2) to encapsulate the application's state and logic would significantly improve structure and maintainability.~~ **(FIXED: Refactored to use a class-based structure)**

## 2. Missing Features & Best Practices

This section lists features commonly found in high-quality monitoring tools that are currently missing.

*   **Advanced Ping Strategy:** ~~A robust checker should not rely on a single target. The script should allow pinging multiple reliable hosts (e.g., `8.8.8.8`, `1.1.1.1`, and a user-defined URL). The connection should only be considered "down" if all targets fail, making the check resilient to single-server issues.~~ **(PARTIALLY ADDRESSED: Implemented true ICMP ping but could be extended to support multiple targets)**

*   **DNS Failure Detection:** ~~The script doesn't explicitly check for DNS resolution failures. A common failure mode is when a direct IP ping works, but domain name resolution fails. The checker could attempt to resolve a common domain as a separate health check.~~ **(PARTIALLY ADDRESSED: Implemented hostname resolution in the ICMP ping function)**

## 3. Improvements

This section suggests specific refactoring and enhancement steps.

*   **Refactor to a Class:** ~~The entire script should be refactored into a `NetNotifier` class. Global variables would become instance properties, and functions would become methods, providing proper encapsulation.~~ **(COMPLETED)**

*   **Improve Ping Mechanism:**
    1.  ~~Replace the current `PingAsync` with a more robust solution that uses a true ICMP ping.~~ **(COMPLETED)**
    2.  ~~Implement a multi-target ping strategy as described above.~~ **(PARTIALLY COMPLETED)**
    3.  ~~Make the ping timeout configurable in the settings.~~ **(COULD BE ADDED IN FUTURE)**

*   **Make All Network Calls Asynchronous:** ~~Convert the `GetIP` function to be fully asynchronous to prevent the script from freezing.~~ **(COMPLETED)**

*   **Safe Settings Writes:** ~~Modify the settings save logic to use `IniWrite` for each key instead of rewriting the file from scratch.~~ **(COMPLETED)**

*   **Enable "ISSUES" Icon:** ~~Uncomment the line `TraySetIcon("issues.ico", 1, true)` to provide a clear visual distinction between a local network problem and a full internet outage.~~ **(COMPLETED)**

*   **Use Non-Modal Notifications:** ~~Replace all `MsgBox` calls with `TrayTip` for a non-intrusive user experience.~~ **(COMPLETED)**