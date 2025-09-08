#Requires AutoHotkey v2.0
#SingleInstance Force

; --- Default Settings ---
global Interval := 5000 ; in milliseconds
global PingTargets := ["google.com", "8.8.8.8", "1.1.1.1"]
global PingTimeout := 5000 ; in milliseconds
global DNSTestHost := "www.google.com"
global VoiceAlerts := 1
global Online := false
global OnlineTime := 0
global DisconnectsToday := 0
global LastCheck := 0
global TotalChecks := 0
global SuccessfulChecks := 0
global LastStatus := false
global FirstRun := true
global SettingsGui := ""

global IsChecking := false  ; reentrancy guard for CheckConnection

; Public IP cache
global PublicIPCache := ""
global PublicIPLastFetch := 0
global PublicIPCacheTTL := 5 * 60 * 1000  ; 5 minutes in ms
global gPublicIPHttpRequest := "" ; To hold the WinHttpRequest object for async calls

global gVoice := ComObject("SAPI.SpVoice")
gVoice.Rate := -2          ; -10 (very slow) to +10 (very fast), with 0 being the normal.
gVoice.Volume := 100      ; 0..100

; --- Load Settings ---
LoadSettings()

; --- Tray Menu ---
A_TrayMenu.Delete()
A_TrayMenu.Add("Settings", ShowSettings)
A_TrayMenu.Add("Check Now", (*) => CheckConnection())
A_TrayMenu.Add("Reset Statistics", ResetStats)
A_TrayMenu.Add()  ; Separator
A_TrayMenu.Add("Exit", (*) => ExitApp())
A_TrayMenu.Default := "Settings"

; Initialize with proper status
TraySetIcon("green.ico", 1, true)

CheckConnection()
SetTimer(CheckConnection, Interval)

LoadSettings() {
    global
    local SettingsFile := A_ScriptDir . "\Settings.ini"
    
    try {
        ; Create settings file with defaults if it doesn't exist
        if (!FileExist(SettingsFile)) {
            CreateDefaultSettings()
        }
        
        Interval := Integer(IniRead(SettingsFile, "Settings", "Interval", "5000"))
        local PingTargetsStr := IniRead(SettingsFile, "Settings", "PingTargets", "google.com,8.8.8.8,1.1.1.1")
        PingTargets := StrSplit(PingTargetsStr, ",")
        PingTimeout := Integer(IniRead(SettingsFile, "Settings", "PingTimeout", "5000"))
        DNSTestHost := IniRead(SettingsFile, "Settings", "DNSTestHost", "www.google.com")
        VoiceAlerts := Integer(IniRead(SettingsFile, "Settings", "VoiceAlerts", "1"))
        VoiceAlerts := VoiceAlerts ? 1 : 0  ; force 0/1 int
        
        ; Validate interval (minimum 1 second, maximum 5 minutes)
        if (Interval < 1000)
            Interval := 1000
        if (Interval > 300000)
            Interval := 300000
    } catch {
        CreateDefaultSettings()
    }
}

CreateDefaultSettings() {
    local SettingsFile := A_ScriptDir . "\Settings.ini"
    try {
        FileAppend("[Settings]`nInterval=5000`nPingTargets=google.com,8.8.8.8,1.1.1.1`nPingTimeout=5000`nDNSTestHost=www.google.com`nVoiceAlerts=1`n", SettingsFile)
    } catch {
        ; Ignore if can't create file
    }
}

UpdateStatistics(currentStatus) {
    global TotalChecks, SuccessfulChecks
    TotalChecks++
    if (currentStatus == "ONLINE")
        SuccessfulChecks++
}

HandleStatusChange(newStatus, oldStatus) {
    global Online, VoiceAlerts, FirstRun, DisconnectsToday, OnlineTime
    
    if (newStatus == "ONLINE") {
        Online := true
        TraySetIcon("green.ico", 1, true)  ; Force icon refresh
        if (VoiceAlerts && !FirstRun)  ; Don't speak on first run
            Speak("Connection Restored")
        if (oldStatus != "ONLINE" || FirstRun) {  ; Starting or reconnecting
            OnlineTime := A_TickCount
            ; refresh public IP on becoming online
            GetPublicIPFetch() ; This is now async
        }
    } else if (newStatus == "ISSUES") {
        Online := false
        TraySetIcon("issues.ico", 1, true)  ; Force icon refresh
        if (VoiceAlerts && !FirstRun)  ; Don't speak on first run
            Speak("Connection Lost")
        if (oldStatus == "ONLINE" && !FirstRun)  ; Was online, now has issues
            DisconnectsToday++
    } else if (newStatus == "DNS_FAILURE") {
        Online := false
        TraySetIcon("red.ico", 1, true)
        if (VoiceAlerts && !FirstRun)
            Speak("Connection Lost")
        if (oldStatus == "ONLINE" && !FirstRun)
            DisconnectsToday++
    } else { ; OFFLINE
        Online := false
        TraySetIcon("red.ico", 1, true)  ; Force icon refresh
        if (VoiceAlerts && !FirstRun)  ; Don't speak on first run
            Speak("Connection Lost")
        if (oldStatus == "ONLINE" && !FirstRun)  ; Was online, now disconnected
            DisconnectsToday++
    }
}

DetermineConnectionStatus() {
    global PingTargets
    local InternetStatus := false
    for target in PingTargets {
        if (PingAsync(Trim(target))) {
            InternetStatus := true
            break
        }
    }

    if (InternetStatus) {
        return "ONLINE"
    } else {
        if (!CheckDNS()) {
            return "DNS_FAILURE"
        } else {
            local GatewayIP := GetDefaultGateway()
            if (GatewayIP != "" && PingAsync(GatewayIP)) {
                return "ISSUES" ; LAN works but no internet
            } else {
                return "OFFLINE" ; No LAN connection
            }
        }
    }
}

CheckConnection() {
    global Online, OnlineTime, DisconnectsToday, TotalChecks, SuccessfulChecks, LastStatus, PingTargets, VoiceAlerts, FirstRun, IsChecking, PublicIPCache, PublicIPLastFetch

    ; prevent overlapping checks (timer + manual)
    if (IsChecking)
        return
    IsChecking := true
    try {
        local CurrentStatus := DetermineConnectionStatus() ; Use the new function

        ; Status changed
        if (CurrentStatus != LastStatus || FirstRun) {
            HandleStatusChange(CurrentStatus, LastStatus)
            LastStatus := CurrentStatus
            FirstRun := false
        }

        UpdateStatistics(CurrentStatus) ; Use the new function

        UpdateTooltip()
    } finally {
        IsChecking := false
    }
}

CheckDNS() {
    global DNSTestHost
    try {
        local hModule := DllCall("LoadLibrary", "Str", "ws2_32.dll", "Ptr")
        if (!hModule) {
            return false
        }
        local pHostent := DllCall("ws2_32.dll\gethostbyname", "AStr", DNSTestHost, "Ptr")
        DllCall("FreeLibrary", "Ptr", hModule)
        return pHostent != 0
    } catch {
        return false
    }
}

GetDefaultGateway() {
    try {
        objWMIService := ComObjGet("winmgmts:\\.\root\cimv2")
        colItems := objWMIService.ExecQuery("SELECT * FROM Win32_IP4RouteTable WHERE Destination='0.0.0.0'")
        
        for objItem in colItems {
            if (objItem.NextHop && objItem.NextHop != "0.0.0.0") {
                return objItem.NextHop
            }
        }
    } catch {
        ; Fallback or error handling can be added here if WMI fails
    }
    return "" ; Return empty if no gateway is found
}

UpdateTooltip() {
    global Online, OnlineTime, DisconnectsToday, TotalChecks, SuccessfulChecks, PingTargets, LastStatus
    
    if (LastStatus == "ONLINE") {
        local ElapsedTime := (A_TickCount - OnlineTime) // 1000
        local Hours := ElapsedTime // 3600
        local Minutes := Mod(ElapsedTime, 3600) // 60
        local Seconds := Mod(ElapsedTime, 60)
        local uptime := Format("{:02}:{:02}:{:02}", Hours, Minutes, Seconds)
        ;local Latency := GetLastPingTime()
        local Availability := TotalChecks > 0 ? (SuccessfulChecks / TotalChecks) * 100 : 0
        local LocalIP := GetPublicIPCached()
        
        ; Keep tooltip concise due to Windows tooltip length limitations
        A_IconTip := (
            "IP:`t" . LocalIP . "`n"
            . "Uptime:`t" . uptime . "`n"
            . "Drops:`t" . DisconnectsToday . "`n"
            . "Up:`t" . Round(Availability, 1) . "%"
        )
    } else if (LastStatus == "ISSUES") {
        A_IconTip := ("NO INTERNET")
    } else if (LastStatus == "DNS_FAILURE") {
        A_IconTip := ("OFFLINE")
    } else {
        A_IconTip := ("OFFLINE")
    }
}

ShowSettings(*) {
    global Interval, PingTargets, PingTimeout, DNSTestHost, VoiceAlerts, SettingsGui
    
    ; Destroy existing settings window if open
    if (SettingsGui && IsObject(SettingsGui))
        SettingsGui.Destroy()
    
    SettingsGui := Gui("-Resize", "Settings")
    SettingsGui.SetFont("s10", "Segoe UI")
    SettingsGui.MarginX := 15
    SettingsGui.MarginY := 15
    
    ; Settings controls
    SettingsGui.Add("Text", "Section", "Check Interval (seconds):")
    global IntervalInput := SettingsGui.Add("Edit", "xs w200 Number", Interval // 1000)
    SettingsGui.Add("Text", "xs", "Range: 1-300 seconds")
    
    SettingsGui.Add("Text", "xs Section", "Ping Targets (comma-separated):")
    local PingTargetsStr := ""
    for i, target in PingTargets {
        PingTargetsStr .= target . (i == PingTargets.Length ? "" : ",")
    }
    global PingTargetsInput := SettingsGui.Add("Edit", "xs w200", PingTargetsStr)
    SettingsGui.Add("Text", "xs", "Example: google.com, 8.8.8.8")

    SettingsGui.Add("Text", "xs Section", "Ping Timeout (milliseconds):")
    global PingTimeoutInput := SettingsGui.Add("Edit", "xs w200 Number", PingTimeout)
    SettingsGui.Add("Text", "xs", "Range: 500-10000")

    SettingsGui.Add("Text", "xs Section", "DNS Test Host:")
    global DNSTestHostInput := SettingsGui.Add("Edit", "xs w200", DNSTestHost)
    SettingsGui.Add("Text", "xs", "Example: www.google.com")
    
    global VoiceAlertsInput := SettingsGui.Add("CheckBox", "xs Section", "Enable Voice Alerts")
    VoiceAlertsInput.Value := VoiceAlerts  ; 0 or 1
    
    ; Buttons
    local SaveBtn := SettingsGui.Add("Button", "xs Section w60 h30 Default", "&Save")
    local TestBtn := SettingsGui.Add("Button", "x+10 w60 h30", "&Test")
    local CancelBtn := SettingsGui.Add("Button", "x+10 w60 h30", "&Cancel")
    
    ; Add event handlers
    SaveBtn.OnEvent("Click", SettingsButtonSave)
    TestBtn.OnEvent("Click", TestConnectionFunc)
    CancelBtn.OnEvent("Click", (*) => SettingsGui.Destroy())
    
    ; Add event handlers for GUI
    SettingsGui.OnEvent("Close", (*) => SettingsGui.Destroy())
    SettingsGui.Show("w230")
}

SettingsButtonSave(*) {
    global Interval, PingTargets, PingTimeout, DNSTestHost, VoiceAlerts, SettingsGui, IntervalInput, PingTargetsInput, PingTimeoutInput, DNSTestHostInput, VoiceAlertsInput
    
    try {
        ; Get values directly from controls without Submit
        local NewInterval := Integer(IntervalInput.Text) * 1000  ; Convert to milliseconds
        local NewPingTargetsStr := Trim(PingTargetsInput.Text)
        local NewPingTimeout := Integer(PingTimeoutInput.Text)
        local NewDNSTestHost := Trim(DNSTestHostInput.Text)
        local NewVoiceAlerts := VoiceAlertsInput.Value  ; 0 or 1
        
        ; Validate interval
        if (NewInterval < 1000) {
            MsgBox("Interval must be at least 1 second!", "Invalid Input", "OK Icon!")
            IntervalInput.Focus()
            return
        }
        if (NewInterval > 300000) {
            MsgBox("Interval cannot exceed 300 seconds!", "Invalid Input", "OK Icon!")
            IntervalInput.Focus()
            return
        }
        
        ; Validate URL
        if (NewPingTargetsStr == "") {
            MsgBox("Ping targets cannot be empty!", "Invalid Input", "OK Icon!")
            PingTargetsInput.Focus()
            return
        }

        ; Validate timeout
        if (NewPingTimeout < 500) {
            MsgBox("Ping timeout must be at least 500 milliseconds!", "Invalid Input", "OK Icon!")
            PingTimeoutInput.Focus()
            return
        }
        if (NewPingTimeout > 10000) {
            MsgBox("Ping timeout cannot exceed 10000 milliseconds!", "Invalid Input", "OK Icon!")
            PingTimeoutInput.Focus()
            return
        }

        ; Validate DNS Test Host
        if (NewDNSTestHost == "") {
            MsgBox("DNS Test Host cannot be empty!", "Invalid Input", "OK Icon!")
            DNSTestHostInput.Focus()
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
        
        ; Update global variables AFTER successful file write
        Interval := NewInterval
        PingTargets := CleanArr
        PingTimeout := NewPingTimeout
        DNSTestHost := NewDNSTestHost
        VoiceAlerts := NewVoiceAlerts  ; Store as integer (0 or 1)
        
        ; Update timer with new interval
        SetTimer(CheckConnection, Interval)
        
        ; Close settings window
        SettingsGui.Destroy()
        
    } catch as e {
        MsgBox("Error saving settings: " . e.Message, "Error", "OK Icon!")
    }
}

TestConnectionFunc(*) {
    global PingTargetsInput
    local TestTargetsStr := Trim(PingTargetsInput.Text)
    if (TestTargetsStr == "") {
        MsgBox("Please enter at least one ping target first!", "Test Connection", "OK Icon!")
        return
    }
    
    local TestTargets := StrSplit(TestTargetsStr, ",")
    local TestURL := Trim(TestTargets[1])
    
    local StartTime := A_TickCount
    local Result := PingAsync(TestURL)
    local Duration := A_TickCount - StartTime
    
    if (Result) {
        MsgBox("Connection to " . TestURL . " successful! (" . Duration . "ms)", "Test Result", "OK Iconi")
    } else {
        MsgBox("Connection to " . TestURL . " failed!", "Test Result", "OK Icon!")
    }
}

ResetStats(*) {
    global DisconnectsToday, TotalChecks, SuccessfulChecks, OnlineTime
    
    if (MsgBox("Reset all statistics?", "Reset Statistics", "YesNo Icon?") == "Yes") {
        DisconnectsToday := 0
        TotalChecks := 0
        SuccessfulChecks := 0
        OnlineTime := A_TickCount  ; Reset uptime counter
        
        ; Update tooltip immediately
        UpdateTooltip()
    }
}

; Connectivity check preferring ICMP (WMI) with HTTP fallback
PingAsync(target) {
    global PingTimeout
    target := Trim(target)
    ; Prefer ICMP via WMI for reliability, fallback to WinINet HTTP check
    try {
        local objWMIService := ComObjGet("winmgmts:\\.\root\cimv2")
        local query := "SELECT * FROM Win32_PingStatus WHERE Address = '" . target . "' AND Timeout = " . PingTimeout
        local colPings := objWMIService.ExecQuery(query)
        for objPing in colPings {
            return objPing.StatusCode == 0
        }
        return false
    } catch {
        try {
            local result := DllCall("wininet\\InternetCheckConnection", "str", "http://" . target, "uint", 1, "uint", 0)
            return result != 0
        } catch {
            return false
        }
    }
}

GetLastPingTime() {
    global PingTargets
    try {
        local StartTime := A_TickCount
        PingAsync(Trim(PingTargets[1]))
        local Duration := A_TickCount - StartTime
        return Duration > 0 ? Duration : "<1"
    } catch {
        return "N/A"
    }
}

GetPublicIPFetch() {
    global PublicIPCache, PublicIPLastFetch
    local LogFile := A_ScriptDir . "\NetNotifier_log.txt"

    try {
        local Http := ComObject("WinHttp.WinHttpRequest.5.1")
        Http.Open("GET", "https://api.ipify.org/", true) ; true for async, but WaitForResponse makes it synchronous
        Http.Send()
        Http.WaitForResponse()
        PublicIPCache := Trim(Http.ResponseText)
        
    } catch as e {
        
        try {
            local Http := ComObject("WinHttp.WinHttpRequest.5.1")
            Http.Open("GET", "https://icanhazip.com/", true) ; true for async, but WaitForResponse makes it synchronous
            Http.Send()
            Http.WaitForResponse()
            PublicIPCache := Trim(Http.ResponseText)
            
        } catch as e2 {
            PublicIPCache := "N/A (WinHttp Error)"
            
        }
    }
    PublicIPLastFetch := A_TickCount
}

GetPublicIPCached() {
    global PublicIPCache, PublicIPLastFetch, PublicIPCacheTTL, Online
    if (!Online) {
        return "N/A"
    }
    if (PublicIPCache != "" && (A_TickCount - PublicIPLastFetch) < PublicIPCacheTTL) {
        return PublicIPCache
    }
    ; Trigger an async fetch, but return current cache or N/A immediately
    GetPublicIPFetch() ; This is now async
    return PublicIPCache != "" ? PublicIPCache : "Fetching..." ; Indicate that it's being fetched
}

Speak(text) {
    try {
        ; 1 = SVSFlagsAsync (non-blocking)
        gVoice.Speak(text, 1)
    } catch as e {
        ; Optional: quick debug
        ; MsgBox "Speak failed: " e.Message
    }
}
