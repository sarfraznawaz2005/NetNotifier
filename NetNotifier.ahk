#Requires AutoHotkey v2.0
#SingleInstance Force

class NetNotifierApp {
    ; Constants
    static DEFAULT_INTERVAL := 5000
    static DEFAULT_PING_TIMEOUT := 2000  ; 2 seconds for faster detection
    static MIN_INTERVAL := 1000
    static MAX_INTERVAL := 300000
    static MIN_PING_TIMEOUT := 500
    static MAX_PING_TIMEOUT := 10000
    static IP_CACHE_TTL := 5 * 60 * 1000  ; 5 minutes
    static GATEWAY_CACHE_TTL := 10 * 60 * 1000  ; 10 minutes for gateway
    static DEFAULT_DNS_HOST := "www.google.com"
    static DEFAULT_PING_TARGETS := ["google.com"]  ; Single reliable target for faster detection

    ; Properties (formerly global variables)
    Interval := NetNotifierApp.DEFAULT_INTERVAL
    PingTargets := NetNotifierApp.DEFAULT_PING_TARGETS.Clone()
    PingTimeout := NetNotifierApp.DEFAULT_PING_TIMEOUT
    DNSTestHost := NetNotifierApp.DEFAULT_DNS_HOST
    VoiceAlerts := 1
    Online := false
    OnlineTime := 0
    DisconnectsToday := 0
    LastCheck := 0
    TotalChecks := 0
    SuccessfulChecks := 0
    LastStatus := false
    FirstRun := true
    SettingsGui := ""
    IsChecking := false
    PublicIPCache := ""
    PublicIPLastFetch := 0
    PublicIPCacheTTL := NetNotifierApp.IP_CACHE_TTL
    gPublicIPHttpRequest := ""
    GatewayCache := ""
    GatewayLastFetch := 0
    GatewayCacheTTL := NetNotifierApp.GATEWAY_CACHE_TTL
    gVoice := ""
    ; GUI input controls
    IntervalInput := ""
    PingTargetsInput := ""
    PingTimeoutInput := ""
    DNSTestHostInput := ""
    VoiceAlertsInput := ""

    __New() {
        this.gVoice := ComObject("SAPI.SpVoice")
        this.gVoice.Rate := -2
        this.gVoice.Volume := 100
        this.LoadSettings()
    }

    Log(message, level := "INFO") {
        try {
            local LogFile := A_ScriptDir . "\NetNotifier.log"
            local Timestamp := FormatTime(A_Now, "yyyy-MM-dd HH:mm:ss")
            local LogEntry := Format("[{}] [{}] {}", Timestamp, level, message)
            ;FileAppend(LogEntry . "`n", LogFile)
        } catch {
            ; Silent fail if logging fails
        }
    }

    LoadSettings() {
        local SettingsFile := A_ScriptDir . "\Settings.ini"
        
        try {
            ; Create settings file with defaults if it doesn't exist
            if (!FileExist(SettingsFile)) {
                this.CreateDefaultSettings()
            }
            
            this.Interval := Integer(IniRead(SettingsFile, "Settings", "Interval", NetNotifierApp.DEFAULT_INTERVAL))
            local PingTargetsStr := IniRead(SettingsFile, "Settings", "PingTargets", "google.com")
            this.PingTargets := StrSplit(PingTargetsStr, ",")
            this.PingTimeout := Integer(IniRead(SettingsFile, "Settings", "PingTimeout", NetNotifierApp.DEFAULT_PING_TIMEOUT))
            this.DNSTestHost := IniRead(SettingsFile, "Settings", "DNSTestHost", NetNotifierApp.DEFAULT_DNS_HOST)
            this.VoiceAlerts := Integer(IniRead(SettingsFile, "Settings", "VoiceAlerts", "1"))
            this.VoiceAlerts := this.VoiceAlerts ? 1 : 0  ; force 0/1 int

            ; Validate interval
            if (this.Interval < NetNotifierApp.MIN_INTERVAL)
                this.Interval := NetNotifierApp.MIN_INTERVAL
            if (this.Interval > NetNotifierApp.MAX_INTERVAL)
                this.Interval := NetNotifierApp.MAX_INTERVAL
        } catch {
            this.CreateDefaultSettings()
        }
    }

    CreateDefaultSettings() {
        local SettingsFile := A_ScriptDir . "\Settings.ini"
        try {
            local PingTargetsStr := ""
            for i, target in this.DEFAULT_PING_TARGETS {
                PingTargetsStr .= target . (i == this.DEFAULT_PING_TARGETS.Length ? "" : ",")
            }
            FileAppend("[Settings]`nInterval=" . this.DEFAULT_INTERVAL . "`nPingTargets=" . PingTargetsStr . "`nPingTimeout=" . this.DEFAULT_PING_TIMEOUT . "`nDNSTestHost=" . this.DEFAULT_DNS_HOST . "`nVoiceAlerts=1`n", SettingsFile)
        } catch {
            ; Ignore if can't create file
        }
    }

    UpdateStatistics(currentStatus) {
        this.TotalChecks++
        if (currentStatus == "ONLINE" || currentStatus == "ISSUES")
            this.SuccessfulChecks++  ; Count partial connectivity as successful
    }

    HandleStatusChange(newStatus, oldStatus) {
        if (newStatus == "ONLINE") {
            this.Online := true
            this.SetIconAndSpeak("green.ico", "Connection Restored")
            if (oldStatus != "ONLINE" || this.FirstRun) {  ; Starting or reconnecting
                this.OnlineTime := A_TickCount
                ; refresh public IP on becoming online
                this.GetPublicIPFetch() ; This is now async
            }
        } else if (newStatus == "ISSUES") {
            this.Online := false
            this.SetIconAndSpeak("issues.ico", "Connection Lost")
            if (oldStatus == "ONLINE" && !this.FirstRun)  ; Was online, now has issues
                this.DisconnectsToday++
        } else if (newStatus == "DNS_FAILURE") {
            this.Online := false
            this.SetIconAndSpeak("red.ico", "Connection Lost")
            if (oldStatus == "ONLINE" && !this.FirstRun)
                this.DisconnectsToday++
        } else { ; OFFLINE
            this.Online := false
            this.SetIconAndSpeak("red.ico", "Connection Lost")
            if (oldStatus == "ONLINE" && !this.FirstRun)  ; Was online, now disconnected
                this.DisconnectsToday++
        }
    }

    SetIconAndSpeak(icon, message) {
        TraySetIcon(icon, 1, true)  ; Force icon refresh
        if (this.VoiceAlerts && !this.FirstRun)  ; Don't speak on first run
            this.Speak(message)
    }

    DetermineConnectionStatus() {
        this.Log("Starting connectivity check", "INFO")

        ; First, check network interface status for immediate disconnection detection
        if (!this.CheckNetworkInterfaceStatus()) {
            this.Log("Network interface check failed - immediate offline detection", "INFO")
            return "OFFLINE"
        }

        local InternetStatus := this.PingParallel(this.PingTargets)

        if (InternetStatus) {
            ; Check for captive portal
            this.Log("Internet ping succeeded, checking for captive portal", "DEBUG")
            if (this.CheckCaptivePortal()) {
                this.Log("Captive portal detected", "WARN")
                return "ISSUES" ; Internet reachable but captive portal detected
            }
            this.Log("Connection status: ONLINE", "INFO")
            return "ONLINE"
        } else {
            this.Log("All pings failed, checking DNS", "DEBUG")
            if (!this.CheckDNS()) {
                this.Log("DNS check failed", "WARN")
                return "DNS_FAILURE"
            } else {
                this.Log("DNS check passed, checking gateway", "DEBUG")
                local GatewayIP := this.GetDefaultGateway()
                this.Log("Gateway IP: " . GatewayIP, "DEBUG")
                if (GatewayIP != "" && this.PingAsync(GatewayIP)) {
                    this.Log("Gateway ping succeeded, status: ISSUES", "INFO")
                    return "ISSUES" ; LAN works but no internet
                } else {
                    this.Log("Gateway ping failed, status: OFFLINE", "INFO")
                    return "OFFLINE" ; No LAN connection
                }
            }
        }
    }

    CheckConnection() {
        ; prevent overlapping checks (timer + manual)
        if (this.IsChecking)
            return
        this.IsChecking := true
        try {
            local CurrentStatus := this.DetermineConnectionStatus()

            ; Status changed
            if (CurrentStatus != this.LastStatus || this.FirstRun) {
                this.HandleStatusChange(CurrentStatus, this.LastStatus)
                this.LastStatus := CurrentStatus
                this.FirstRun := false
            }

            this.UpdateStatistics(CurrentStatus)

            this.UpdateTooltip()
        } catch as e {
            this.Log("Error in CheckConnection: " . e.Message, "ERROR")
        } finally {
            this.IsChecking := false
        }
    }

    CheckDNS() {
        this.Log("Checking DNS for " . this.DNSTestHost, "DEBUG")
        ; Try multiple methods for DNS check
        try {
            ; Method 1: Use nslookup command
            local cmd := 'nslookup ' . this.DNSTestHost
            local result := RunWait(cmd, , "Hide")
            if (result == 0) {
                this.Log("DNS check passed with nslookup", "DEBUG")
                return true
            } else {
                this.Log("DNS check failed with nslookup, exit code: " . result, "DEBUG")
            }
        } catch as e {
            this.Log("nslookup failed: " . e.Message, "WARN")
        }

        try {
            ; Method 2: Fallback to gethostbyname
            local hModule := DllCall("LoadLibrary", "Str", "ws2_32.dll", "Ptr")
            if (!hModule) {
                this.Log("Failed to load ws2_32.dll", "WARN")
                return false
            }
            local pHostent := DllCall("ws2_32.dll\gethostbyname", "AStr", this.DNSTestHost, "Ptr")
            DllCall("FreeLibrary", "Ptr", hModule)
            if (pHostent != 0) {
                this.Log("DNS check passed with gethostbyname", "DEBUG")
                return true
            } else {
                this.Log("DNS check failed with gethostbyname", "DEBUG")
                return false
            }
        } catch as e {
            this.Log("gethostbyname failed: " . e.Message, "WARN")
            return false
        }
    }

    GetDefaultGateway() {
        ; Check cache first
        if (this.GatewayCache != "" && (A_TickCount - this.GatewayLastFetch) < this.GatewayCacheTTL) {
            this.Log("Using cached gateway: " . this.GatewayCache, "DEBUG")
            return this.GatewayCache
        }

        try {
            objWMIService := ComObjGet("winmgmts:\\.\root\cimv2")
            colItems := objWMIService.ExecQuery("SELECT * FROM Win32_IP4RouteTable WHERE Destination='0.0.0.0'")

            for objItem in colItems {
                if (objItem.NextHop && objItem.NextHop != "0.0.0.0") {
                    this.GatewayCache := objItem.NextHop
                    this.GatewayLastFetch := A_TickCount
                    this.Log("Found gateway: " . this.GatewayCache, "DEBUG")
                    return this.GatewayCache
                }
            }
            this.Log("No default gateway found", "WARN")
        } catch as e {
            this.Log("Failed to get default gateway: " . e.Message, "ERROR")
        }
        return "" ; Return empty if no gateway is found
    }

    ; Check network interface status for immediate disconnection detection
    CheckNetworkInterfaceStatus() {
        try {
            objWMIService := ComObjGet("winmgmts:\\.\root\cimv2")
            colItems := objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapter WHERE NetConnectionStatus = 2")  ; 2 = Connected

            local connectedInterfaces := 0
            for objItem in colItems {
                if (objItem.NetEnabled && objItem.NetConnectionStatus == 2) {
                    connectedInterfaces++
                    this.Log("Network interface connected: " . objItem.Name, "DEBUG")
                }
            }

            if (connectedInterfaces == 0) {
                this.Log("No network interfaces connected", "WARN")
                return false
            }

            this.Log("Found " . connectedInterfaces . " connected network interface(s)", "DEBUG")
            return true
        } catch as e {
            this.Log("Failed to check network interface status: " . e.Message, "ERROR")
            return true  ; Assume connected if we can't check
        }
    }

    UpdateTooltip() {
        if (this.LastStatus == "ONLINE") {
            local ElapsedTime := (A_TickCount - this.OnlineTime) // 1000
            local Hours := ElapsedTime // 3600
            local Minutes := Mod(ElapsedTime, 3600) // 60
            local Seconds := Mod(ElapsedTime, 60)
            local uptime := Format("{:02}:{:02}:{:02}", Hours, Minutes, Seconds)
            local Latency := this.GetLastPingTime()
            local LocalIP := this.GetPublicIPCached()
            
            local Availability := this.TotalChecks > 0 ? (this.SuccessfulChecks / this.TotalChecks) * 100 : 0
            
            ; Ensure that if there have been disconnects, uptime never shows as 100%
            local DisplayAvailability := Round(Availability, 1)
            if (this.DisconnectsToday > 0 && DisplayAvailability = 100.0) {
                DisplayAvailability := 99.9
            }
            
            ; Keep tooltip concise due to Windows tooltip length limitations
            A_IconTip := (
                "IP:`t" . LocalIP . "`n"
                . "Uptime:`t" . uptime . "`n"
                . "Latency:`t" . Latency . "ms`n"
                . "Drops:`t" . this.DisconnectsToday . "`n"
                . "Up:`t" . DisplayAvailability . "%"
            )
        } else if (this.LastStatus == "ISSUES") {
            A_IconTip := ("NO INTERNET")
        } else if (this.LastStatus == "DNS_FAILURE") {
            A_IconTip := ("OFFLINE")
        } else {
            A_IconTip := ("OFFLINE")
        }
    }

    ShowSettings(*) {
        ; Destroy existing settings window if open
        if (this.SettingsGui && IsObject(this.SettingsGui))
            this.SettingsGui.Destroy()
        
        this.SettingsGui := Gui("-Resize", "Settings")
        this.SettingsGui.SetFont("s10", "Segoe UI")
        this.SettingsGui.MarginX := 15
        this.SettingsGui.MarginY := 15
        
        ; Settings controls
        this.SettingsGui.Add("Text", "Section", "Check Interval (seconds):")
        this.IntervalInput := this.SettingsGui.Add("Edit", "xs w200 Number", this.Interval // 1000)
        this.SettingsGui.Add("Text", "xs", "Range: " . (NetNotifierApp.MIN_INTERVAL // 1000) . "-" . (NetNotifierApp.MAX_INTERVAL // 1000) . " seconds")

        this.SettingsGui.Add("Text", "xs Section", "Ping Targets (comma-separated):")
        local PingTargetsStr := ""
        for i, target in this.PingTargets {
            PingTargetsStr .= target . (i == this.PingTargets.Length ? "" : ",")
        }
        this.PingTargetsInput := this.SettingsGui.Add("Edit", "xs w200", PingTargetsStr)
        this.SettingsGui.Add("Text", "xs", "Ex: google.com, 8.8.8.8, 1.1.1.1")

        this.SettingsGui.Add("Text", "xs Section", "Ping Timeout (milliseconds):")
        this.PingTimeoutInput := this.SettingsGui.Add("Edit", "xs w200 Number", this.PingTimeout)
        this.SettingsGui.Add("Text", "xs", "Range: " . NetNotifierApp.MIN_PING_TIMEOUT . "-" . NetNotifierApp.MAX_PING_TIMEOUT . " ms")

        this.SettingsGui.Add("Text", "xs Section", "DNS Test Host:")
        this.DNSTestHostInput := this.SettingsGui.Add("Edit", "xs w200", this.DNSTestHost)
        this.SettingsGui.Add("Text", "xs", "Example: www.google.com")
        
        this.VoiceAlertsInput := this.SettingsGui.Add("CheckBox", "xs Section", "Enable Voice Alerts")
        this.VoiceAlertsInput.Value := this.VoiceAlerts  ; 0 or 1
        
        ; Buttons
        local SaveBtn := this.SettingsGui.Add("Button", "xs Section w60 h30 Default", "&Save")
        local TestBtn := this.SettingsGui.Add("Button", "x+10 w60 h30", "&Test")
        local CancelBtn := this.SettingsGui.Add("Button", "x+10 w60 h30", "&Cancel")
        
        ; Add event handlers
        SaveBtn.OnEvent("Click", (*) => this.SettingsButtonSave())
        TestBtn.OnEvent("Click", (*) => this.TestConnectionFunc())
        CancelBtn.OnEvent("Click", (*) => this.SettingsGui.Destroy())
        
        ; Add event handlers for GUI
        this.SettingsGui.OnEvent("Close", (*) => this.SettingsGui.Destroy())
        this.SettingsGui.Show("w230")
    }

    SettingsButtonSave(*) {
        try {
            ; Get values directly from controls without Submit
            local NewInterval := Integer(this.IntervalInput.Text) * 1000  ; Convert to milliseconds
            local NewPingTargetsStr := Trim(this.PingTargetsInput.Text)
            local NewPingTimeout := Integer(this.PingTimeoutInput.Text)
            local NewDNSTestHost := Trim(this.DNSTestHostInput.Text)
            local NewVoiceAlerts := this.VoiceAlertsInput.Value  ; 0 or 1
            
            ; Validate interval
            if (NewInterval < NetNotifierApp.MIN_INTERVAL) {
                MsgBox("Interval must be at least " . (NetNotifierApp.MIN_INTERVAL // 1000) . " seconds!", "Invalid Input", "OK Icon!")
                this.IntervalInput.Focus()
                return
            }
            if (NewInterval > NetNotifierApp.MAX_INTERVAL) {
                MsgBox("Interval cannot exceed " . (NetNotifierApp.MAX_INTERVAL // 1000) . " seconds!", "Invalid Input", "OK Icon!")
                this.IntervalInput.Focus()
                return
            }

            ; Validate URL
            if (NewPingTargetsStr == "") {
                MsgBox("Ping targets cannot be empty!", "Invalid Input", "OK Icon!")
                this.PingTargetsInput.Focus()
                return
            }

            ; Validate timeout
            if (NewPingTimeout < NetNotifierApp.MIN_PING_TIMEOUT) {
                MsgBox("Ping timeout must be at least " . NetNotifierApp.MIN_PING_TIMEOUT . " milliseconds!", "Invalid Input", "OK Icon!")
                this.PingTimeoutInput.Focus()
                return
            }
            if (NewPingTimeout > NetNotifierApp.MAX_PING_TIMEOUT) {
                MsgBox("Ping timeout cannot exceed " . NetNotifierApp.MAX_PING_TIMEOUT . " milliseconds!", "Invalid Input", "OK Icon!")
                this.PingTimeoutInput.Focus()
                return
            }

            ; Validate DNS Test Host
            if (NewDNSTestHost == "") {
                MsgBox("DNS Test Host cannot be empty!", "Invalid Input", "OK Icon!")
                this.DNSTestHostInput.Focus()
                return
            }
            
            ; Normalize and trim ping targets
            local TargetsArr := []
            for t in StrSplit(NewPingTargetsStr, ",")
                TargetsArr.Push(Trim(t))
            ; Remove empty entries (if any)
            local CleanArr := []
            for t in TargetsArr {
                if (t != "")
                    CleanArr.Push(t)
            }
            local NormalizedTargetsStr := ""
            for i, t in CleanArr
                NormalizedTargetsStr .= t . (i == CleanArr.Length ? "" : ",")

            ; Write to INI file FIRST
            local SettingsFile := A_ScriptDir . "\Settings.ini"
            
            IniWrite(NewInterval, SettingsFile, "Settings", "Interval")
            IniWrite(NormalizedTargetsStr, SettingsFile, "Settings", "PingTargets")
            IniWrite(NewPingTimeout, SettingsFile, "Settings", "PingTimeout")
            IniWrite(NewDNSTestHost, SettingsFile, "Settings", "DNSTestHost")
            IniWrite(NewVoiceAlerts ? 1 : 0, SettingsFile, "Settings", "VoiceAlerts")
            
            ; Update properties AFTER successful file write
            this.Interval := NewInterval
            this.PingTargets := CleanArr
            this.PingTimeout := NewPingTimeout
            this.DNSTestHost := NewDNSTestHost
            this.VoiceAlerts := NewVoiceAlerts  ; Store as integer (0 or 1)
            
            ; Update timer with new interval
            SetTimer(() => this.CheckConnection(), this.Interval)
            
            ; Close settings window
            this.SettingsGui.Destroy()
            
        } catch as e {
            this.Log("Error saving settings: " . e.Message, "ERROR")
        }
    }

    TestConnectionFunc(*) {
        local TestTargetsStr := Trim(this.PingTargetsInput.Text)
        if (TestTargetsStr == "") {
            MsgBox("Please enter at least one ping target first!", "Test Connection", "OK Icon!")
            return
        }
        
        local TestTargets := StrSplit(TestTargetsStr, ",")
        local TestURL := Trim(TestTargets[1])
        
        local StartTime := A_TickCount
        local Result := this.PingAsync(TestURL)
        local Duration := A_TickCount - StartTime
        
        if (Result) {
            MsgBox("Connection to " . TestURL . " successful! (" . Duration . "ms)", "Test Result", "OK Iconi")
        } else {
            MsgBox("Connection to " . TestURL . " failed!", "Test Result", "OK Icon!")
        }
    }

    ResetStats(*) {
        if (MsgBox("Reset all statistics?", "Reset Statistics", "YesNo Icon?") == "Yes") {
            this.DisconnectsToday := 0
            this.TotalChecks := 0
            this.SuccessfulChecks := 0
            this.OnlineTime := A_TickCount  ; Reset uptime counter
            
            ; Update tooltip immediately
            this.UpdateTooltip()
        }
    }

    ; Connectivity check preferring ICMP (WMI) with HTTP fallback
    PingAsync(target, retries := 2, parallelMode := false) {
        target := Trim(target)
        this.Log("Starting ping to " . target . " with " . retries . " retries" . (parallelMode ? " (parallel)" : ""), "DEBUG")

        ; Skip IPv6 for now if it contains :
        if (InStr(target, ":")) {
            this.Log("Skipping IPv6 target " . target, "DEBUG")
            return false
        }

        ; Use shorter timeout in parallel mode
        local originalTimeout := this.PingTimeout
        if (parallelMode) {
            this.PingTimeout := Max(500, this.PingTimeout // 2)
        }

        try {
            Loop retries + 1 {
                ; Add jitter to avoid network congestion
                if (A_Index > 1) {
                    Sleep(Random(100, 500))
                }

                ; Prefer ICMP via WMI for reliability, fallback to WinINet HTTP check
                try {
                    objWMIService := ComObjGet("winmgmts:\\.\root\cimv2")
                    query := "SELECT * FROM Win32_PingStatus WHERE Address = '" . target . "' AND Timeout = " . this.PingTimeout
                    colPings := objWMIService.ExecQuery(query)
                    for objPing in colPings {
                        if (objPing.StatusCode == 0) {
                            this.Log("WMI ping succeeded for " . target . " on attempt " . A_Index, "DEBUG")
                            return true
                        }
                    }
                    this.Log("WMI ping failed for " . target . " on attempt " . A_Index . ", status: " . (objPing ? objPing.StatusCode : "unknown"), "DEBUG")
                } catch as e {
                    this.Log("WMI ping exception for " . target . " on attempt " . A_Index . ": " . e.Message, "WARN")
                }

                ; Fallback to HTTP check
                try {
                    result := DllCall("wininet\\InternetCheckConnection", "str", "http://" . target, "uint", 1, "uint", 0)
                    if (result != 0) {
                        this.Log("HTTP check succeeded for " . target . " on attempt " . A_Index, "DEBUG")
                        return true
                    } else {
                        this.Log("HTTP check failed for " . target . " on attempt " . A_Index, "DEBUG")
                    }
                } catch as e {
                    this.Log("HTTP check exception for " . target . " on attempt " . A_Index . ": " . e.Message, "WARN")
                }
            }
            this.Log("All ping attempts failed for " . target, "WARN")
            return false
        } finally {
            if (parallelMode) {
                this.PingTimeout := originalTimeout
            }
        }
    }

    ; Parallel ping implementation for faster detection
    PingParallel(targets) {
        this.Log("Starting parallel ping to " . targets.Length . " targets", "DEBUG")

        ; Use rapid succession pinging with reduced retries
        try {
            for target in targets {
                local cleanTarget := Trim(target)
                if (cleanTarget == "")
                    continue

                this.Log("Pinging " . cleanTarget . " (parallel mode)", "DEBUG")
                if (this.PingAsync(cleanTarget, 0, true)) {  ; No retries, parallel mode
                    this.Log("Parallel ping succeeded for " . cleanTarget, "DEBUG")
                    return true
                }
            }
        } catch as e {
            this.Log("Parallel ping exception: " . e.Message, "WARN")
        }

        this.Log("All parallel pings failed", "WARN")
        return false
    }

    ; TCP connectivity check
    CheckTCPConnection(host, port := 80) {
        try {
            ; Initialize Winsock
            local WSADATA := Buffer(400, 0)
            local result := DllCall("ws2_32\WSAStartup", "ushort", 0x0202, "ptr", WSADATA, "int")
            if (result != 0) {
                return false
            }
            
            ; Create socket
            local socket := DllCall("ws2_32\socket", "int", 2, "int", 1, "int", 6, "ptr") ; AF_INET, SOCK_STREAM, IPPROTO_TCP
            if (socket == -1) {
                DllCall("ws2_32\WSACleanup")
                return false
            }
            
            ; Connect
            local sockaddr := Buffer(16, 0)
            NumPut("short", 2, sockaddr, 0) ; AF_INET
            NumPut("short", DllCall("ws2_32\htons", "ushort", port, "ushort"), sockaddr, 2)
            NumPut("int", DllCall("ws2_32\inet_addr", "astr", host, "int"), sockaddr, 4)
            
            local connectResult := DllCall("ws2_32\connect", "ptr", socket, "ptr", sockaddr, "int", 16, "int")
            DllCall("ws2_32\closesocket", "ptr", socket)
            DllCall("ws2_32\WSACleanup")
            
            return connectResult == 0
        } catch as e {
            this.Log("TCP check failed for " . host . ":" . port . ": " . e.Message, "WARN")
            return false
        }
    }

    ; Captive portal detection
    CheckCaptivePortal() {
        this.Log("Starting captive portal check", "DEBUG")
        try {
            Http := ComObject("WinHttp.WinHttpRequest.5.1")
            Http.Open("GET", "http://www.google.com/", false)
            Http.Send()

            this.Log("Captive portal HTTP status: " . Http.Status, "DEBUG")

            if (Http.Status == 200) {
                ; Check if response contains typical captive portal content
                response := Http.ResponseText
                ; Only check for obvious captive portal indicators
                if (InStr(response, "login") && InStr(response, "password")) {
                    this.Log("Captive portal detected in response", "WARN")
                    return true ; Likely captive portal
                }
                this.Log("No captive portal detected", "DEBUG")
            } else if (Http.Status >= 300 && Http.Status < 400) {
                this.Log("HTTP redirect detected, possible captive portal", "WARN")
                return true ; Redirect, likely captive portal
            }
            return false
        } catch as e {
            this.Log("Captive portal check failed: " . e.Message, "WARN")
            return false
        }
    }

    GetLastPingTime() {
        try {
            local StartTime := A_TickCount
            local Result := this.PingAsync(Trim(this.PingTargets[1]))
            local Duration := A_TickCount - StartTime
            return Result ? (Duration > 0 ? Duration : "<1") : "N/A"
        } catch as e {
            this.Log("Error getting ping time: " . e.Message, "WARN")
            return "N/A"
        }
    }

    ; Basic speed test by downloading a small file
    GetDownloadSpeed() {
        try {
            local StartTime := A_TickCount
            local Http := ComObject("WinHttp.WinHttpRequest.5.1")
            Http.Open("GET", "http://speedtest.tele2.net/1MB.zip", false) ; Small test file
            Http.Send()
            local EndTime := A_TickCount
            if (Http.Status == 200) {
                local Bytes := Http.ResponseBody.MaxIndex() + 1
                local Duration := (EndTime - StartTime) / 1000  ; seconds
                local Speed := (Bytes * 8) / Duration / 1000  ; Kbps
                return Round(Speed, 1) . " Kbps"
            }
        } catch as e {
            this.Log("Speed test failed: " . e.Message, "WARN")
        }
        return "N/A"
    }

    GetPublicIPFetch() {
        try {
            Http := ComObject("WinHttp.WinHttpRequest.5.1")
            Http.Open("GET", "https://api.ipify.org/", false)
            Http.Send()
            if (Http.Status == 200) {
                this.PublicIPCache := Trim(Http.ResponseText)
                this.Log("Fetched public IP: " . this.PublicIPCache, "INFO")
            } else {
                throw Error("HTTP " . Http.Status)
            }
        } catch as e {
            this.Log("Failed to fetch IP from api.ipify.org: " . e.Message, "WARN")
            try {
                Http := ComObject("WinHttp.WinHttpRequest.5.1")
                Http.Open("GET", "https://icanhazip.com/", false)
                Http.Send()
                if (Http.Status == 200) {
                    this.PublicIPCache := Trim(Http.ResponseText)
                    this.Log("Fetched public IP from fallback: " . this.PublicIPCache, "INFO")
                } else {
                    this.PublicIPCache := "N/A (HTTP " . Http.Status . ")"
                }
            } catch as e2 {
                this.Log("Failed to fetch IP from icanhazip.com: " . e2.Message, "ERROR")
                this.PublicIPCache := "N/A (Error)"
            }
        }
        this.PublicIPLastFetch := A_TickCount
    }

    GetPublicIPCached() {
        if (!this.Online) {
            return "N/A"
        }
        if (this.PublicIPCache != "" && (A_TickCount - this.PublicIPLastFetch) < this.PublicIPCacheTTL) {
            return this.PublicIPCache
        }
        ; Trigger an async fetch, but return current cache or N/A immediately
        this.GetPublicIPFetch() ; This is now async
        return this.PublicIPCache != "" ? this.PublicIPCache : "Fetching..." ; Indicate that it's being fetched
    }

    Speak(text) {
        try {
            ; 1 = SVSFlagsAsync (non-blocking)
            this.gVoice.Speak(text, 1)
        } catch as e {
            ; Optional: quick debug
            ; MsgBox "Speak failed: " e.Message
        }
    }
}

; Create app instance
global app := NetNotifierApp()

; --- Tray Menu ---
A_TrayMenu.Delete()
A_TrayMenu.Add("Settings", (*) => app.ShowSettings())
A_TrayMenu.Add("Check Now", (*) => app.CheckConnection())
A_TrayMenu.Add("Reset Statistics", (*) => app.ResetStats())
A_TrayMenu.Add()  ; Separator
A_TrayMenu.Add("Exit", (*) => ExitApp())
A_TrayMenu.Default := "Settings"

; Initialize with proper status
TraySetIcon("green.ico", 1, true)

app.CheckConnection()
SetTimer(() => app.CheckConnection(), app.Interval)