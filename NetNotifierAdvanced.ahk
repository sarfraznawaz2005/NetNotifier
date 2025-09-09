#Requires AutoHotkey v2.0
#SingleInstance Force

class NetNotifierApp {
    ; Constants
    static DEFAULT_INTERVAL := 5000  ; 5 seconds for balanced detection
    static DEFAULT_PING_TIMEOUT := 5000  ; 5 seconds for reliable timeout
    static MIN_INTERVAL := 5000  ; 5 seconds minimum to reduce false positives
    static MAX_INTERVAL := 300000
    static MIN_PING_TIMEOUT := 3000  ; 3 seconds minimum timeout
    static MAX_PING_TIMEOUT := 10000
    static IP_CACHE_TTL := 5 * 60 * 1000  ; 5 minutes
    static GATEWAY_CACHE_TTL := 10 * 60 * 1000  ; 10 minutes for gateway
    static DEFAULT_DNS_HOST := "1.1.1.1"
    static DEFAULT_PING_TARGETS := ["1.1.1.1"]  ; Single reliable target for faster detection

    ; Properties (formerly global variables)
    Interval := NetNotifierApp.DEFAULT_INTERVAL
    PingTargets := NetNotifierApp.DEFAULT_PING_TARGETS.Clone()
    PingTimeout := NetNotifierApp.DEFAULT_PING_TIMEOUT
    DNSTestHost := NetNotifierApp.DEFAULT_DNS_HOST
    VoiceAlerts := 1
    Online := false
    OnlineTime := 0
    DisconnectsToday := 0
    TotalChecks := 0
    SuccessfulChecks := 0
    LastStatus := false
    FirstRun := true
    SettingsGui := ""
    IsChecking := false
    PublicIPCache := ""
    PublicIPLastFetch := 0
    PublicIPCacheTTL := NetNotifierApp.IP_CACHE_TTL
    GatewayCache := ""
    GatewayLastFetch := 0
    GatewayCacheTTL := NetNotifierApp.GATEWAY_CACHE_TTL
    gVoice := ""
    ; Async flags for connectivity checks
    DNSWorking := true  ; Assume true initially
    CaptivePortalDetected := false
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
            FileAppend(LogEntry . "`n", LogFile)
        } catch {
            ; Silent fail if logging fails
        }
    }

    _GetWMIService() {
        try {
            return ComObjGet("winmgmts:\\.\root\cimv2")
        } catch as e {
            this.Log("Failed to get WMI service: " . e.Message, "ERROR")
            return ""
        }
    }

    _BuildPingTargetsStr(targets) {
        local str := ""
        for i, target in targets {
            str .= target . (i == targets.Length ? "" : ",")
        }
        return str
    }

    LoadSettings() {
        local SettingsFile := A_ScriptDir . "\Settings.ini"
        
        try {
            ; Create settings file with defaults if it doesn't exist
            if (!FileExist(SettingsFile)) {
                this.CreateDefaultSettings()
            }
            
            this.Interval := Integer(IniRead(SettingsFile, "Settings", "Interval", NetNotifierApp.DEFAULT_INTERVAL))
            local PingTargetsStr := IniRead(SettingsFile, "Settings", "PingTargets", "1.1.1.1")
            this.PingTargets := StrSplit(PingTargetsStr, ",")
            this.PingTimeout := Integer(IniRead(SettingsFile, "Settings", "PingTimeout", NetNotifierApp.DEFAULT_PING_TIMEOUT))
            this.DNSTestHost := IniRead(SettingsFile, "Settings", "DNSTestHost", "1.1.1.1")
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
            local PingTargetsStr := this._BuildPingTargetsStr(this.DEFAULT_PING_TARGETS)
            FileAppend("[Settings]`nInterval=" . this.DEFAULT_INTERVAL . "`nPingTargets=" . PingTargetsStr . "`nPingTimeout=" . this.DEFAULT_PING_TIMEOUT . "`nDNSTestHost=" . this.DEFAULT_DNS_HOST . "`nVoiceAlerts=1`n", SettingsFile)
        } catch {
            ; Ignore if can't create file
        }
    }

    UpdateStatistics(currentStatus) {
        this.TotalChecks++
        if (currentStatus == "ONLINE")
            this.SuccessfulChecks++  ; Only count fully online as successful
    }

    HandleStatusChange(newStatus, oldStatus) {
        if (newStatus == "ONLINE") {
            this.Online := true
            if (this.VoiceAlerts && oldStatus != "ONLINE" && !this.FirstRun)  ; Speak only on change, not on first run
                this.Speak("Connection Restored")
            if (oldStatus != "ONLINE" || this.FirstRun) {  ; Starting or reconnecting
                this.OnlineTime := A_TickCount
                ; refresh public IP on becoming online
                this.GetPublicIPFetch() ; This is now async
            }
        } else if (newStatus == "ISSUES") {
            this.Online := false
            if (this.VoiceAlerts && oldStatus != "ISSUES" && !this.FirstRun)  ; Speak only on change, not on first run
                this.Speak("Connection Lost")
            if (oldStatus == "ONLINE" && !this.FirstRun)  ; Was online, now has issues
                this.DisconnectsToday++
        } else if (newStatus == "DNS_FAILURE") {
            this.Online := false
            if (this.VoiceAlerts && oldStatus != "DNS_FAILURE" && !this.FirstRun)  ; Speak only on change, not on first run
                this.Speak("Connection Lost")
            if (oldStatus == "ONLINE" && !this.FirstRun)
                this.DisconnectsToday++
        } else { ; OFFLINE
            this.Online := false
            if (this.VoiceAlerts && oldStatus != "OFFLINE" && !this.FirstRun)  ; Speak only on change, not on first run
                this.Speak("Connection Lost")
            if (oldStatus == "ONLINE" && !this.FirstRun)  ; Was online, now disconnected
                this.DisconnectsToday++
        }
    }



    DetermineConnectionStatus() {
        this.Log("Starting connectivity check", "INFO")

        ; First, check network interface status for immediate disconnection detection
        if (!this.CheckNetworkInterfaceStatus()) {
            this.Log("Network interface check failed - immediate offline detection", "INFO")
            return "OFFLINE"
        }

        local InternetStatus := false
        for target in this.PingTargets {
            local cleanTarget := Trim(target)
            if (cleanTarget == "")
                continue
            if (this.PingAsync(cleanTarget)) {
                InternetStatus := true
                break
            }
        }

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
        this.DNSWorking := false  ; Reset flag
        ; Async DNS check
        this._CheckDNSAsync()
        ; Wait up to 2 seconds for async result
        local start := A_TickCount
        while (A_TickCount - start < 2000) {
            if (this.DNSWorking !== false)  ; Flag has been set
                break
            Sleep(50)
        }
        return this.DNSWorking
    }

    _CheckDNSAsync() {
        try {
            ; Method 1: Use nslookup command async
            local outputFile := A_Temp . "\dns_check.txt"
            local cmd := 'nslookup ' . this.DNSTestHost . ' > "' . outputFile . '" 2>&1'
            Run(cmd, , "Hide")
            ; Poll for result
            this._DNSPollTimer := ObjBindMethod(this, "_DNSPoll", outputFile)
            SetTimer(this._DNSPollTimer, 100)
        } catch as e {
            this.Log("nslookup failed: " . e.Message, "WARN")
            this._CheckDNSFallback()
        }
    }

    _DNSPoll(outputFile) {
        if (FileExist(outputFile)) {
            local content := FileRead(outputFile)
            FileDelete(outputFile)
            SetTimer(this._DNSPollTimer, 0)
            if (InStr(content, "Name:")) {
                this.Log("DNS check passed with nslookup", "DEBUG")
                this.DNSWorking := true
            } else {
                this.Log("DNS check failed with nslookup", "DEBUG")
                this.DNSWorking := false
                this._CheckDNSFallback()
            }
        }
    }

    _CheckDNSFallback() {
        try {
            ; Method 2: Fallback to gethostbyname
            local hModule := DllCall("LoadLibrary", "Str", "ws2_32.dll", "Ptr")
            if (!hModule) {
                this.Log("Failed to load ws2_32.dll", "WARN")
                this.DNSWorking := false
                return
            }
            local pHostent := DllCall("ws2_32.dll\gethostbyname", "AStr", this.DNSTestHost, "Ptr")
            DllCall("FreeLibrary", "Ptr", hModule)
            if (pHostent != 0) {
                this.Log("DNS check passed with gethostbyname", "DEBUG")
                this.DNSWorking := true
            } else {
                this.Log("DNS check failed with gethostbyname", "DEBUG")
                this.DNSWorking := false
            }
        } catch as e {
            this.Log("gethostbyname failed: " . e.Message, "WARN")
            this.DNSWorking := false
        }
    }

    GetDefaultGateway() {
        ; Check cache first
        if (this.GatewayCache != "" && (A_TickCount - this.GatewayLastFetch) < this.GatewayCacheTTL) {
            this.Log("Using cached gateway: " . this.GatewayCache, "DEBUG")
            return this.GatewayCache
        }

        try {
            objWMIService := this._GetWMIService()
            if (!objWMIService)
                return ""
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
            objWMIService := this._GetWMIService()
            if (!objWMIService)
                return true  ; Assume connected if can't check
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
        ; Set icon based on current status
        if (this.LastStatus == "ONLINE") {
            TraySetIcon("green.ico", 1, true)
            local ElapsedTime := (A_TickCount - this.OnlineTime) // 1000
            local Hours := ElapsedTime // 3600
            local Minutes := Mod(ElapsedTime, 3600) // 60
            local Seconds := Mod(ElapsedTime, 60)
            local uptime := Format("{:02}:{:02}:{:02}", Hours, Minutes, Seconds)
            local Latency := this.GetLastPingTime()
            local LocalIP := this.GetPublicIPCached()

            local Availability := this.TotalChecks > 0 ? (this.SuccessfulChecks / this.TotalChecks) * 100 : 0

            ; Format to exactly one decimal place to avoid floating point precision issues
            local DisplayAvailability := Format("{:.1f}", Availability)

            ; Keep tooltip concise due to Windows tooltip length limitations
            A_IconTip := (
                "IP:`t" . LocalIP . "`n"
                . "Uptime:`t" . uptime . "`n"
                . "Latency:`t" . Latency . "ms`n"
                . "Drops:`t" . this.DisconnectsToday . "`n"
                . "Up:`t" . DisplayAvailability . "%"
            )
        } else if (this.LastStatus == "ISSUES") {
            TraySetIcon("issues.ico", 1, true)
            A_IconTip := ("NO INTERNET")
        } else if (this.LastStatus == "DNS_FAILURE") {
            TraySetIcon("red.ico", 1, true)
            A_IconTip := ("OFFLINE")
        } else {
            TraySetIcon("red.ico", 1, true)
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
        local PingTargetsStr := this._BuildPingTargetsStr(this.PingTargets)
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
    PingAsync(target, retries := 2) {
        target := Trim(target)
        this.Log("Starting ping to " . target . " with " . retries . " retries", "DEBUG")

        ; Skip IPv6 for now if it contains :
        if (InStr(target, ":")) {
            this.Log("Skipping IPv6 target " . target, "DEBUG")
            return false
        }

        try {
            Loop retries + 1 {
                ; Jitter removed to avoid blocking

                ; Prefer ICMP via WMI for reliability, fallback to WinINet HTTP check
                try {
                    objWMIService := ComObjGet("winmgmts:\\.\root\cimv2")
                    query := "SELECT * FROM Win32_PingStatus WHERE Address = '" . target . "' AND Timeout = " . this.PingTimeout
                    colPings := objWMIService.ExecQuery(query)
                    local objPing := ""
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
                local Http := ComObject("MSXML2.ServerXMLHTTP")
                Http.Open("GET", "https://icanhazip.com/", false)
                Http.Send()
            if (Http.status == 200) {
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
        }
    }





    ; Captive portal detection
    CheckCaptivePortal() {
        this.Log("Starting captive portal check", "DEBUG")
        this.CaptivePortalDetected := false  ; Reset flag
        ; Check if we can access a known URL without redirection
        url := "http://www.msftconnecttest.com/connecttest.txt"
        callback := ObjBindMethod(this, "_HandleCaptiveResponse")
        this.HttpRequestAsync("GET", url, Map(), "", callback)
        ; Wait up to 2 seconds for async result
        local start := A_TickCount
        while (A_TickCount - start < 2000) {
            if (this.CaptivePortalDetected !== false)  ; Flag has been set
                break
            Sleep(50)
        }
        return this.CaptivePortalDetected
    }

    _HandleCaptiveResponse(response, status, headers) {
        if (status == 200 && InStr(response, "Microsoft Connect Test")) {
            this.Log("No captive portal detected", "DEBUG")
            this.CaptivePortalDetected := false
        } else {
            this.Log("Captive portal detected or connection issue", "WARN")
            this.CaptivePortalDetected := true
        }
    }

    ; Async HTTP request with callback
    HttpRequestAsync(method, url, headersObj, bodyText, callback) {
        req := ComObject("MSXML2.ServerXMLHTTP")
        req.open(method, url, true)
        if IsSet(headersObj) {
            for k, v in headersObj {
                if (k != "__timeoutMs")
                    req.setRequestHeader(k, v)
            }
        }
        deadline := 0
        try {
            if IsSet(headersObj) && headersObj.HasOwnProp("__timeoutMs")
                deadline := A_TickCount + Integer(headersObj.__timeoutMs)
        } catch {
        }
        try {
            req.send(IsSet(bodyText) ? bodyText : "")
        } catch as e {
            try callback("", -1, "")
            return
        }
        poll := 0
        poll := () => (
            (deadline && A_TickCount > deadline)
                ? ( SetTimer(poll, 0), callback ? callback("", -1, "") : 0 )
            : ( req.readyState = 4
                ? ( SetTimer(poll, 0), callback ? callback(req.responseText, req.status, req.getAllResponseHeaders()) : 0 )
                : 0 )
        )
        SetTimer(poll, 30)
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

    GetPublicIPFetch() {
        ; Async fetch with retry
        this._FetchIPWithRetry("https://api.ipify.org/")
    }

    _FetchIPWithRetry(url, attempt := 1) {
        maxAttempts := 2
        fallbackUrl := "https://icanhazip.com/"
        callback := ObjBindMethod(this, "_HandleIPResponse", url, attempt, maxAttempts, fallbackUrl)
        this.HttpRequestAsync("GET", url, Map(), "", callback)
    }

    _HandleIPResponse(url, attempt, maxAttempts, fallbackUrl, response, status, headers) {
        if (status == 200 && response != "") {
            this.PublicIPCache := Trim(response)
            this.PublicIPLastFetch := A_TickCount
            this.Log("Fetched public IP: " . this.PublicIPCache, "INFO")
        } else {
            this.Log("Failed to fetch IP from " . url . " (attempt " . attempt . "): HTTP " . status, "WARN")
            if (attempt < maxAttempts) {
                if (attempt == 1) {
                    this._FetchIPWithRetry(fallbackUrl, attempt + 1)
                } else {
                    this.PublicIPCache := "N/A (HTTP " . status . ")"
                    this.PublicIPLastFetch := A_TickCount
                }
            } else {
                this.PublicIPCache := "N/A (Error)"
                this.PublicIPLastFetch := A_TickCount
            }
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

    Speak(text) {
        try {
            ; 1 = SVSFlagsAsync (non-blocking)
            this.gVoice.Speak(text, 1)
        } catch as e {
            ; Silent fail
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