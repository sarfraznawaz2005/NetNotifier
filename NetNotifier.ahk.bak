#Requires AutoHotkey v2.0
#SingleInstance Force

; --- Default Settings ---
global Interval := 3000 ; in milliseconds
global PingURL := "google.com"
global VoiceAlerts := 0
global Online := false
global OnlineTime := 0
global DisconnectsToday := 0
global LastCheck := 0
global TotalChecks := 0
global SuccessfulChecks := 0
global LastStatus := false
global FirstRun := true
global SettingsGui := ""

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
        
        Interval := Integer(IniRead(SettingsFile, "Settings", "Interval", "3000"))
        PingURL := IniRead(SettingsFile, "Settings", "PingURL", "google.com")
        VoiceAlerts := Integer(IniRead(SettingsFile, "Settings", "VoiceAlerts", "0"))
        VoiceAlerts := VoiceAlerts = "1" ? 1 : 0  ; force 0/1 int
        
        ; Validate interval (minimum 1 second, maximum 5 minutes)
        if (Interval < 1000)
            Interval := 1000
        if (Interval > 300000)
            Interval := 300000
    } catch {
        ; Use defaults if reading fails
        Interval := 3000
        PingURL := "google.com"
        VoiceAlerts := 0
        CreateDefaultSettings()
    }
}

CreateDefaultSettings() {
    local SettingsFile := A_ScriptDir . "\Settings.ini"
    try {
        FileAppend("[Settings]`nInterval=3000`nPingURL=google.com`nVoiceAlerts=0`n", SettingsFile)
    } catch {
        ; Ignore if can't create file
    }
}

CheckConnection() {
    global Online, OnlineTime, DisconnectsToday, TotalChecks, SuccessfulChecks, LastStatus, PingURL, VoiceAlerts, FirstRun
    
    ; Use async ping to prevent UI freezing
    local InternetStatus := PingAsync(PingURL)
    local CurrentStatus := ""
    
    if (InternetStatus) {
        ; Internet is working - ONLINE
        CurrentStatus := "ONLINE"
    } else {
        ; Internet failed, check local gateway
        local GatewayIP := GetDefaultGateway()
        if (GatewayIP != "" && PingAsync(GatewayIP)) {
            ; LAN works but no internet - ISSUES
            CurrentStatus := "ISSUES"
        } else {
            ; No LAN connection - OFFLINE
            CurrentStatus := "OFFLINE"
        }
    }
    
    ; Status changed
    if (CurrentStatus != LastStatus || FirstRun) {
        if (CurrentStatus == "ONLINE") {
            Online := true
            TraySetIcon("green.ico", 1, true)  ; Force icon refresh
            if (VoiceAlerts && !FirstRun)  ; Don't speak on first run
                Speak("Connection Restored")
            if (LastStatus != "ONLINE" || FirstRun)  ; Starting or reconnecting
                OnlineTime := A_TickCount
        } else if (CurrentStatus == "ISSUES") {
            Online := false
            ;TraySetIcon("issues.ico", 1, true)  ; Force icon refresh
            TraySetIcon("red.ico", 1, true)  ; Force icon refresh
            if (VoiceAlerts && !FirstRun)  ; Don't speak on first run
                Speak("Connection Lost")
            if (LastStatus == "ONLINE" && !FirstRun)  ; Was online, now has issues
                DisconnectsToday++
        } else {
            Online := false
            TraySetIcon("red.ico", 1, true)  ; Force icon refresh
            if (VoiceAlerts && !FirstRun)  ; Don't speak on first run
                Speak("Connection Lost")
            if (LastStatus == "ONLINE" && !FirstRun)  ; Was online, now disconnected
                DisconnectsToday++
        }
        LastStatus := CurrentStatus
        FirstRun := false
    }

    TotalChecks++
    if (CurrentStatus == "ONLINE")
        SuccessfulChecks++

    UpdateTooltip()
}

GetDefaultGateway() {
    try {
        ; Method 1: Try using WMI (faster and cleaner)
        local objWMIService := ComObjGet("winmgmts:\\.\\root\cimv2")
        local colItems := objWMIService.ExecQuery("SELECT * FROM Win32_IP4RouteTable WHERE Destination='0.0.0.0'")
        
        for objItem in colItems {
            if (objItem.NextHop && objItem.NextHop != "0.0.0.0") {
                return objItem.NextHop
            }
        }
        
        ; Method 2: Fallback using GetIpForwardTable via DLL
        return GetGatewayViaDLL()
    } catch {
        ; Method 3: Last resort fallback
        return GetGatewayViaDLL()
    }
}

GetGatewayViaDLL() {
    try {
        ; Get the size needed for the IP forward table
        local dwSize := 0
        local result := DllCall("iphlpapi\GetIpForwardTable", "ptr", 0, "uint*", &dwSize, "int", 0)
        
        if (dwSize > 0) {
            ; Allocate buffer
            local pIpForwardTable := Buffer(dwSize, 0)
            
            ; Get the actual table
            result := DllCall("iphlpapi\GetIpForwardTable", "ptr", pIpForwardTable, "uint*", &dwSize, "int", 0)
            
            if (result == 0) {  ; NO_ERROR
                ; Read number of entries (first DWORD)
                local dwNumEntries := NumGet(pIpForwardTable, 0, "UInt")
                
                ; Each entry is 56 bytes, starting at offset 4
                local entrySize := 56
                local offset := 4
                
                ; Look for default route (destination 0.0.0.0)
                Loop dwNumEntries {
                    local dwForwardDest := NumGet(pIpForwardTable, offset, "UInt")
                    
                    ; Check if this is default route (0.0.0.0)
                    if (dwForwardDest == 0) {
                        ; Get next hop (gateway) - offset +12 from entry start
                        local dwForwardNextHop := NumGet(pIpForwardTable, offset + 12, "UInt")
                        
                        ; Convert to IP string
                        local ip := ((dwForwardNextHop & 0xFF)) . "." 
                               . ((dwForwardNextHop >> 8) & 0xFF) . "." 
                               . ((dwForwardNextHop >> 16) & 0xFF) . "." 
                               . ((dwForwardNextHop >> 24) & 0xFF)
                        
                        if (ip != "0.0.0.0") {
                            return ip
                        }
                    }
                    offset += entrySize
                }
            }
        }
        return ""
    } catch {
        return ""
    }
}

UpdateTooltip() {
    global Online, OnlineTime, DisconnectsToday, TotalChecks, SuccessfulChecks, PingURL, LastStatus
    
    if (LastStatus == "ONLINE") {
        local ElapsedTime := (A_TickCount - OnlineTime) // 1000
        local Hours := ElapsedTime // 3600
        local Minutes := Mod(ElapsedTime, 3600) // 60
        local Seconds := Mod(ElapsedTime, 60)
        local uptime := Format("{:02}:{:02}:{:02}", Hours, Minutes, Seconds)
        ;local Latency := GetLastPingTime()
        local Availability := TotalChecks > 0 ? (SuccessfulChecks / TotalChecks) * 100 : 0
        local LocalIP := GetIP()
        
        ; Keep tooltip concise due to Windows tooltip length limitations
        A_IconTip := (
            "IP:`t" . LocalIP . "`n"
            . "Uptime:`t" . uptime . "`n"
            . "Drops:`t" . DisconnectsToday . "`n"
            . "Up:`t" . Round(Availability, 1) . "%"
        )
    } else if (LastStatus == "ISSUES") {
        A_IconTip := ("NO INTERNET")
    } else {
        A_IconTip := ("OFFLINE")
    }
}

ShowSettings(*) {
    global Interval, PingURL, VoiceAlerts, SettingsGui
    
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
    
    SettingsGui.Add("Text", "xs Section", "Ping Target:")
    global PingURLInput := SettingsGui.Add("Edit", "xs w200", PingURL)
    SettingsGui.Add("Text", "xs", "Example: google.com, 8.8.8.8")
    
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
    SettingsGui.Show("w230 h310")
}

SettingsButtonSave(*) {
    global Interval, PingURL, VoiceAlerts, SettingsGui, IntervalInput, PingURLInput, VoiceAlertsInput
    
    try {
        ; Get values directly from controls without Submit
        local NewInterval := Integer(IntervalInput.Text) * 1000  ; Convert to milliseconds
        local NewPingURL := Trim(PingURLInput.Text)
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
        if (NewPingURL == "") {
            MsgBox("Ping target cannot be empty!", "Invalid Input", "OK Icon!")
            PingURLInput.Focus()
            return
        }
        
        ; Write to INI file FIRST
        local SettingsFile := A_ScriptDir . "\Settings.ini"
        
        ; Delete and recreate the file to ensure clean write
        try {
            FileDelete(SettingsFile)
        } catch {
            ; File might not exist, continue
        }
        
        ; Create new settings file - convert checkbox boolean to integer for INI
        local VoiceAlertsValue := NewVoiceAlerts ? 1 : 0
        local SettingsContent := "[Settings]`nInterval=" . NewInterval . "`nPingURL=" . NewPingURL . "`nVoiceAlerts=" . VoiceAlertsValue . "`n"
        FileAppend(SettingsContent, SettingsFile)
        
        ; Update global variables AFTER successful file write
        Interval := NewInterval
        PingURL := NewPingURL
        VoiceAlerts := VoiceAlertsValue  ; Store as integer (0 or 1)
        
        ; Update timer with new interval
        SetTimer(CheckConnection, Interval)
        
        ; Test new connection immediately
        ; CheckConnection()
        
        ; Close settings window
        SettingsGui.Destroy()
        
    } catch as e {
        MsgBox("Error saving settings: " . e.Message, "Error", "OK Icon!")
    }
}

TestConnectionFunc(*) {
    global PingURL, PingURLInput
    local TestURL := Trim(PingURLInput.Text)
    if (TestURL == "") {
        MsgBox("Please enter a ping target first!", "Test Connection", "OK Icon!")
        return
    }
    
    local StartTime := A_TickCount
    local Result := PingAsync(TestURL)
    local Duration := A_TickCount - StartTime
    
    if (Result) {
        MsgBox("Connection successful! (" . Duration . "ms)", "Test Result", "OK Iconi")
    } else {
        MsgBox("Connection failed!", "Test Result", "OK Icon!")
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

; Async ping function to prevent UI freezing
PingAsync(url) {
    try {
        ; Method 1: Try WinINet API (faster)
        local result := DllCall("wininet\InternetCheckConnection", "str", "http://" . url, "uint", 1, "uint", 0)
        return result != 0
    } catch {
        ; Method 2: Fallback to WMI ping (more reliable but slower)
        try {
            ; Use a timeout to prevent hanging
            local objWMIService := ComObjGet("winmgmts:\\.\\root\cimv2")
            local colPings := objWMIService.ExecQuery("SELECT * FROM Win32_PingStatus WHERE Address = '" . url . "' AND Timeout = 3000")
            
            for objPing in colPings {
                return objPing.StatusCode == 0
            }
            return false
        } catch {
            return false
        }
    }
}

GetLastPingTime() {
    global PingURL
    try {
        local StartTime := A_TickCount
        PingAsync(PingURL)
        local Duration := A_TickCount - StartTime
        return Duration > 0 ? Duration : "<1"
    } catch {
        return "N/A"
    }
}

GetIP() {
    try {
        Http := ComObject("WinHttp.WinHttpRequest.5.1")
        Http.Open("GET", "https://api.ipify.org/", true)
        Http.Send()
        Http.WaitForResponse()
        return Http.ResponseText
    } catch {
        try {
            Http := ComObject("WinHttp.WinHttpRequest.5.1")
            Http.Open("GET", "https://icanhazip.com/", true)
            Http.Send()
            Http.WaitForResponse()
            return Http.ResponseText
        } catch {
            return "N/A"
        }
    }
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