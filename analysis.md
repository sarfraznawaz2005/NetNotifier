# AutoHotkey v2 Code Analysis – Async/Blocking Issues

## Identified Blocking/Async Problematic Lines
- Line 216: local result := RunWait(cmd, , "Hide") -> ⚠️ Blocking call: RunWait halts script until external process finishes.
- Line 380: this.SettingsGui.Show("w230") -> ⚠️ UI modal call: Gui.Show is blocking if not handled properly.
- Line 525: Sleep(Random(100, 500)) -> ⚠️ Blocking call: Sleep halts script execution.
- Line 629: Http := ComObject("WinHttp.WinHttpRequest.5.1") -> ⚠️ Potential blocking: synchronous HTTP requests unless async flag set.
- Line 671: local Http := ComObject("WinHttp.WinHttpRequest.5.1") -> ⚠️ Potential blocking: synchronous HTTP requests unless async flag set.
- Line 689: Http := ComObject("WinHttp.WinHttpRequest.5.1") -> ⚠️ Potential blocking: synchronous HTTP requests unless async flag set.
- Line 701: Http := ComObject("WinHttp.WinHttpRequest.5.1") -> ⚠️ Potential blocking: synchronous HTTP requests unless async flag set.


## Possible Strategies to Fix Blocking/Async Issues
1. **Timers (`SetTimer`)** – Offload checks like ping or HTTP requests to timer-driven functions instead of direct blocking calls.
2. **Async HTTP (WinHttpRequest.Open)** – Use `xhr.Open("GET", url, true)` with async flag set to `true`, and connect event handlers using `ComObjConnect`.
3. **Non-blocking Ping** – Replace `RunWait` with background `Run` and capture results asynchronously (redirect output to a file, then read via timer).
4. **Avoid Sleep** – Replace `Sleep` with `SetTimer` or non-blocking waits.
5. **UI Responsiveness** – Avoid modal message boxes or `.Show` that block. Use `Gui` events and `OnEvent` for handling user actions.

## TODO List
- [ ] Refactor all `RunWait` calls (Line ~216) → use `Run` with redirected output.
- [ ] Audit all `ComObjCreate("WinHttpRequest")` usages → ensure async flag enabled and event handlers connected.
- [ ] Replace any `Sleep` calls with `SetTimer` callbacks.
- [ ] Make ping checks non-blocking (wrap in timer-driven subprocesses).
- [ ] Adjust GUI `.Show` to avoid freezing main thread, e.g., by not using modal flags.
