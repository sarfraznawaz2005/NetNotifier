# AutoHotkey v2 Code Analysis – Async/Blocking Issues

## Identified Blocking/Async Problematic Lines
Below are the main lines where blocking or synchronous operations occur:

- Line 216: local result := RunWait(cmd, , "Hide") -> ⚠️ Blocking call: RunWait halts script until external process finishes.
- Line 380: this.SettingsGui.Show("w230") -> ⚠️ UI modal call: Gui.Show is blocking if not handled properly.
- Line 525: Sleep(Random(100, 500)) -> ⚠️ Blocking call: Sleep halts script execution.
- Line 629: Http := ComObject("WinHttp.WinHttpRequest.5.1") -> ⚠️ Potential blocking: synchronous HTTP requests unless async flag set.
- Line 671: local Http := ComObject("WinHttp.WinHttpRequest.5.1") -> ⚠️ Potential blocking: synchronous HTTP requests unless async flag set.
- Line 689: Http := ComObject("WinHttp.WinHttpRequest.5.1") -> ⚠️ Potential blocking: synchronous HTTP requests unless async flag set.
- Line 701: Http := ComObject("WinHttp.WinHttpRequest.5.1") -> ⚠️ Potential blocking: synchronous HTTP requests unless async flag set.

## DllCall and Networking Contexts
The following DllCall usages and networking checks may cause blocking:

### Context around line 229
```
0224:             this.Log("nslookup failed: " . e.Message, "WARN")
0225:         }
0226: 
0227:         try {
0228:             ; Method 2: Fallback to gethostbyname
0229:             local hModule := DllCall("LoadLibrary", "Str", "ws2_32.dll", "Ptr")
0230:             if (!hModule) {
0231:                 this.Log("Failed to load ws2_32.dll", "WARN")
0232:                 return false
0233:             }
0234:             local pHostent := DllCall("ws2_32.dll\gethostbyname", "AStr", this.DNSTestHost, "Ptr")
```

### Context around line 234
```
0229:             local hModule := DllCall("LoadLibrary", "Str", "ws2_32.dll", "Ptr")
0230:             if (!hModule) {
0231:                 this.Log("Failed to load ws2_32.dll", "WARN")
0232:                 return false
0233:             }
0234:             local pHostent := DllCall("ws2_32.dll\gethostbyname", "AStr", this.DNSTestHost, "Ptr")
0235:             DllCall("FreeLibrary", "Ptr", hModule)
0236:             if (pHostent != 0) {
0237:                 this.Log("DNS check passed with gethostbyname", "DEBUG")
0238:                 return true
0239:             } else {
```

### Context around line 235
```
0230:             if (!hModule) {
0231:                 this.Log("Failed to load ws2_32.dll", "WARN")
0232:                 return false
0233:             }
0234:             local pHostent := DllCall("ws2_32.dll\gethostbyname", "AStr", this.DNSTestHost, "Ptr")
0235:             DllCall("FreeLibrary", "Ptr", hModule)
0236:             if (pHostent != 0) {
0237:                 this.Log("DNS check passed with gethostbyname", "DEBUG")
0238:                 return true
0239:             } else {
0240:                 this.Log("DNS check failed with gethostbyname", "DEBUG")
```

### Context around line 546
```
0541:                     this.Log("WMI ping exception for " . target . " on attempt " . A_Index . ": " . e.Message, "WARN")
0542:                 }
0543: 
0544:                 ; Fallback to HTTP check
0545:                 try {
0546:                     result := DllCall("wininet\\InternetCheckConnection", "str", "http://" . target, "uint", 1, "uint", 0)
0547:                     if (result != 0) {
0548:                         this.Log("HTTP check succeeded for " . target . " on attempt " . A_Index, "DEBUG")
0549:                         return true
0550:                     } else {
0551:                         this.Log("HTTP check failed for " . target . " on attempt " . A_Index, "DEBUG")
```

### Context around line 596
```
0591:     ; TCP connectivity check
0592:     CheckTCPConnection(host, port := 80) {
0593:         try {
0594:             ; Initialize Winsock
0595:             local WSADATA := Buffer(400, 0)
0596:             local result := DllCall("ws2_32\WSAStartup", "ushort", 0x0202, "ptr", WSADATA, "int")
0597:             if (result != 0) {
0598:                 return false
0599:             }
0600:             
0601:             ; Create socket
```

### Context around line 602
```
0597:             if (result != 0) {
0598:                 return false
0599:             }
0600:             
0601:             ; Create socket
0602:             local socket := DllCall("ws2_32\socket", "int", 2, "int", 1, "int", 6, "ptr") ; AF_INET, SOCK_STREAM, IPPROTO_TCP
0603:             if (socket == -1) {
0604:                 DllCall("ws2_32\WSACleanup")
0605:                 return false
0606:             }
0607:             
```

### Context around line 604
```
0599:             }
0600:             
0601:             ; Create socket
0602:             local socket := DllCall("ws2_32\socket", "int", 2, "int", 1, "int", 6, "ptr") ; AF_INET, SOCK_STREAM, IPPROTO_TCP
0603:             if (socket == -1) {
0604:                 DllCall("ws2_32\WSACleanup")
0605:                 return false
0606:             }
0607:             
0608:             ; Connect
0609:             local sockaddr := Buffer(16, 0)
```

### Context around line 611
```
0606:             }
0607:             
0608:             ; Connect
0609:             local sockaddr := Buffer(16, 0)
0610:             NumPut("short", 2, sockaddr, 0) ; AF_INET
0611:             NumPut("short", DllCall("ws2_32\htons", "ushort", port, "ushort"), sockaddr, 2)
0612:             NumPut("int", DllCall("ws2_32\inet_addr", "astr", host, "int"), sockaddr, 4)
0613:             
0614:             local connectResult := DllCall("ws2_32\connect", "ptr", socket, "ptr", sockaddr, "int", 16, "int")
0615:             DllCall("ws2_32\closesocket", "ptr", socket)
0616:             DllCall("ws2_32\WSACleanup")
```

### Context around line 612
```
0607:             
0608:             ; Connect
0609:             local sockaddr := Buffer(16, 0)
0610:             NumPut("short", 2, sockaddr, 0) ; AF_INET
0611:             NumPut("short", DllCall("ws2_32\htons", "ushort", port, "ushort"), sockaddr, 2)
0612:             NumPut("int", DllCall("ws2_32\inet_addr", "astr", host, "int"), sockaddr, 4)
0613:             
0614:             local connectResult := DllCall("ws2_32\connect", "ptr", socket, "ptr", sockaddr, "int", 16, "int")
0615:             DllCall("ws2_32\closesocket", "ptr", socket)
0616:             DllCall("ws2_32\WSACleanup")
0617:             
```

### Context around line 614
```
0609:             local sockaddr := Buffer(16, 0)
0610:             NumPut("short", 2, sockaddr, 0) ; AF_INET
0611:             NumPut("short", DllCall("ws2_32\htons", "ushort", port, "ushort"), sockaddr, 2)
0612:             NumPut("int", DllCall("ws2_32\inet_addr", "astr", host, "int"), sockaddr, 4)
0613:             
0614:             local connectResult := DllCall("ws2_32\connect", "ptr", socket, "ptr", sockaddr, "int", 16, "int")
0615:             DllCall("ws2_32\closesocket", "ptr", socket)
0616:             DllCall("ws2_32\WSACleanup")
0617:             
0618:             return connectResult == 0
0619:         } catch as e {
```

### Context around line 615
```
0610:             NumPut("short", 2, sockaddr, 0) ; AF_INET
0611:             NumPut("short", DllCall("ws2_32\htons", "ushort", port, "ushort"), sockaddr, 2)
0612:             NumPut("int", DllCall("ws2_32\inet_addr", "astr", host, "int"), sockaddr, 4)
0613:             
0614:             local connectResult := DllCall("ws2_32\connect", "ptr", socket, "ptr", sockaddr, "int", 16, "int")
0615:             DllCall("ws2_32\closesocket", "ptr", socket)
0616:             DllCall("ws2_32\WSACleanup")
0617:             
0618:             return connectResult == 0
0619:         } catch as e {
0620:             this.Log("TCP check failed for " . host . ":" . port . ": " . e.Message, "WARN")
```

### Context around line 616
```
0611:             NumPut("short", DllCall("ws2_32\htons", "ushort", port, "ushort"), sockaddr, 2)
0612:             NumPut("int", DllCall("ws2_32\inet_addr", "astr", host, "int"), sockaddr, 4)
0613:             
0614:             local connectResult := DllCall("ws2_32\connect", "ptr", socket, "ptr", sockaddr, "int", 16, "int")
0615:             DllCall("ws2_32\closesocket", "ptr", socket)
0616:             DllCall("ws2_32\WSACleanup")
0617:             
0618:             return connectResult == 0
0619:         } catch as e {
0620:             this.Log("TCP check failed for " . host . ":" . port . ": " . e.Message, "WARN")
0621:             return false
```

## Explanation of Issues

- **gethostbyname (Line 229–235)** → Blocking DNS resolution. If DNS server is slow, script hangs.
- **InternetCheckConnection (Line 546)** → Synchronous connectivity check, blocks until completion.
- **Socket connect (Lines 596–616)** → Blocking TCP connect call unless explicitly set non-blocking.
- **RunWait for ping (Line ~216)** → Entire script halts until ping exits.
- **WinHttpRequest (via ComObjCreate)** → Defaults to synchronous unless async flag set.


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
