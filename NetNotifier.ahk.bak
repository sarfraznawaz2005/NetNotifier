#Requires AutoHotkey v2.0
#SingleInstance Force

class NetNotifierApp {
    ; Constants
    static DEFAULT_INTERVAL := 10000  ; 10 seconds for balanced detection
    static MIN_INTERVAL := 5000  ; 5 seconds minimum to reduce false positives
    static MAX_INTERVAL := 300000
    static IP_CACHE_TTL := 5 * 60 * 1000  ; 5 minutes
    static HTTP_TIMEOUT_DEFAULT := 20000  ; 20 seconds default timeout
    static HTTP_TIMEOUT_MIN := 5000  ; 5 seconds minimum
    static HTTP_TIMEOUT_MAX := 30000  ; 30 seconds maximum

    Interval := NetNotifierApp.DEFAULT_INTERVAL
    TestURL := "https://www.google.com"
    VoiceAlerts := 1
    HTTPTimeout := NetNotifierApp.HTTP_TIMEOUT_DEFAULT
    Online := false
    OnlineTime := 0
    DisconnectsToday := 0
    TotalChecks := 0
    SuccessfulChecks := 0
    LastStatus := false
    FirstRun := true
    StatusChangeCount := 0  ; Track rapid status changes
    LastStatusChangeTime := 0
    SettingsGui := ""
    IsChecking := false
    PublicIPCache := ""
    PublicIPLastFetch := 0
    PublicIPCacheTTL := NetNotifierApp.IP_CACHE_TTL
    gVoice := ""
    ; Async flags for connectivity checks
    HTTPWorking := false  ; Flag for async HTTP check
    ; GUI input controls
    IntervalInput := ""
    TestURLInput := ""
    VoiceAlertsInput := ""
    HTTPTimeoutInput := ""

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

    _GetWMIService() {
        try {
            return ComObjGet("winmgmts:\\.\root\cimv2")
        } catch as e {
            this.Log("Failed to get WMI service: " . e.Message, "ERROR")
            return ""
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
            this.TestURL := IniRead(SettingsFile, "Settings", "TestURL", "https://www.google.com")
            this.VoiceAlerts := Integer(IniRead(SettingsFile, "Settings", "VoiceAlerts", "1"))
            this.VoiceAlerts := this.VoiceAlerts ? 1 : 0  ; force 0/1 int
            this.HTTPTimeout := Integer(IniRead(SettingsFile, "Settings", "HTTPTimeout", NetNotifierApp.HTTP_TIMEOUT_DEFAULT))

            ; Validate interval
            if (this.Interval < NetNotifierApp.MIN_INTERVAL)
                this.Interval := NetNotifierApp.MIN_INTERVAL
            if (this.Interval > NetNotifierApp.MAX_INTERVAL)
                this.Interval := NetNotifierApp.MAX_INTERVAL

            ; Validate HTTP timeout
            if (this.HTTPTimeout < NetNotifierApp.HTTP_TIMEOUT_MIN)
                this.HTTPTimeout := NetNotifierApp.HTTP_TIMEOUT_MIN
            if (this.HTTPTimeout > NetNotifierApp.HTTP_TIMEOUT_MAX)
                this.HTTPTimeout := NetNotifierApp.HTTP_TIMEOUT_MAX
        } catch {
            this.CreateDefaultSettings()
        }
    }

    CreateDefaultSettings() {
        local SettingsFile := A_ScriptDir . "\Settings.ini"
        try {
            FileAppend("[Settings]`nInterval=" . this.DEFAULT_INTERVAL . "`nTestURL=https://www.google.com`nVoiceAlerts=1`nHTTPTimeout=" . this.HTTP_TIMEOUT_DEFAULT . "`n", SettingsFile)
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
        ; Add hysteresis to prevent rapid status changes
        local currentTime := A_TickCount
        if (newStatus != oldStatus) {
            this.StatusChangeCount++
            ; If status changed too recently, ignore unless it's a significant change
            if (currentTime - this.LastStatusChangeTime < this.Interval && this.StatusChangeCount > 2) {
                this.Log("Ignoring rapid status change to prevent flapping", "WARN")
                return
            }
            this.LastStatusChangeTime := currentTime
            this.StatusChangeCount := 0
        } else {
            this.StatusChangeCount := 0  ; Reset counter on stable status
        }

        if (newStatus == "ONLINE") {
            this.Online := true
            if (this.VoiceAlerts && oldStatus != "ONLINE" && !this.FirstRun)  ; Speak only on change, not on first run
                this.Speak("Connection Restored")
            if (oldStatus != "ONLINE" || this.FirstRun) {  ; Starting or reconnecting
                this.OnlineTime := A_TickCount
                ; refresh public IP on becoming online
                ;this.GetPublicIPFetch() ; async
                this.GetPublicIPSync()
            }
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

        ; Check internet via HTTP to google.com
        if (this.CheckInternetHTTP()) {
            this.Log("HTTP check succeeded - ONLINE", "INFO")
            return "ONLINE"
        } else {
            this.Log("HTTP check failed - OFFLINE", "INFO")
            return "OFFLINE"
        }
    }

    CheckInternetHTTP() {
        this.Log("Starting async HTTP check to " . this.TestURL, "DEBUG")
        this.HTTPWorking := false  ; Reset flag
        ; Async HTTP check
        this._CheckHTTPAsync()
        ; Wait up to HTTPTimeout for async result
        local start := A_TickCount
        while (A_TickCount - start < this.HTTPTimeout) {
            if (this.HTTPWorking !== false)  ; Flag has been set
                break
            Sleep(50)
        }
        ; Add small buffer to prevent rapid fluctuations
        if (!this.HTTPWorking) {
            Sleep(1000)  ; Wait 1 second before confirming offline
        }
        return this.HTTPWorking
    }

    _CheckHTTPAsync() {
        callback := ObjBindMethod(this, "_HandleHTTPResponse")
        headersObj := Map("__timeoutMs", this.HTTPTimeout)
        this.HttpRequestAsync("GET", this.TestURL, headersObj, "", callback)
    }

    _HandleHTTPResponse(response, status, headers) {
        if (status == 200) {
            this.Log("Async HTTP check succeeded", "DEBUG")
            this.HTTPWorking := true
        } else {
            this.Log("Async HTTP check failed: HTTP " . status, "DEBUG")
            this.HTTPWorking := false
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
            TraySetIcon("online.ico", 1, true)
            
            local ElapsedTime := (A_TickCount - this.OnlineTime) // 1000
            local Hours := ElapsedTime // 3600
            local Minutes := Mod(ElapsedTime, 3600) // 60
            local Seconds := Mod(ElapsedTime, 60)
            local uptime := Format("{:02}:{:02}:{:02}", Hours, Minutes, Seconds)
            local LocalIP := this.GetPublicIPCached()

            local Availability := this.TotalChecks > 0 ? (this.SuccessfulChecks / this.TotalChecks) * 100 : 0

            ; Format to exactly one decimal place to avoid floating point precision issues
            local DisplayAvailability := Format("{:.1f}", Availability)

            ; Keep tooltip concise due to Windows tooltip length limitations
            A_IconTip := (
                "IP:`t" . LocalIP . "`n"
                . "Uptime:`t" . uptime . "`n"
                . "Drops:`t" . this.DisconnectsToday . "`n"
                . "Up:`t" . DisplayAvailability . "%"
            )
        } else {
            TraySetIcon("offline.ico", 1, true)
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
        this.SettingsGui.Add("Text", "Section", "Test URL:")
        this.TestURLInput := this.SettingsGui.Add("Edit", "xs w200", this.TestURL)

        this.SettingsGui.Add("Text", "xs Section", "Check Interval (seconds):")
        this.IntervalInput := this.SettingsGui.Add("Edit", "xs w200 Number", this.Interval // 1000)

        this.SettingsGui.Add("Text", "xs Section", "HTTP Timeout (ms):")
        this.HTTPTimeoutInput := this.SettingsGui.Add("Edit", "xs w200 Number", this.HTTPTimeout)

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
        this.SettingsGui.Show("")
    }

    SettingsButtonSave(*) {
        try {
            ; Get values directly from controls without Submit
            local NewInterval := Integer(this.IntervalInput.Text) * 1000  ; Convert to milliseconds
            local NewTestURL := Trim(this.TestURLInput.Text)
            local NewVoiceAlerts := this.VoiceAlertsInput.Value  ; 0 or 1
            local NewHTTPTimeout := Integer(this.HTTPTimeoutInput.Text)

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

            ; Validate Test URL
            if (NewTestURL == "") {
                MsgBox("Test URL cannot be empty!", "Invalid Input", "OK Icon!")
                this.TestURLInput.Focus()
                return
            }

            ; Validate HTTP Timeout
            if (NewHTTPTimeout < NetNotifierApp.HTTP_TIMEOUT_MIN) {
                MsgBox("HTTP Timeout must be at least " . NetNotifierApp.HTTP_TIMEOUT_MIN . " ms!", "Invalid Input", "OK Icon!")
                this.HTTPTimeoutInput.Focus()
                return
            }
            if (NewHTTPTimeout > NetNotifierApp.HTTP_TIMEOUT_MAX) {
                MsgBox("HTTP Timeout cannot exceed " . NetNotifierApp.HTTP_TIMEOUT_MAX . " ms!", "Invalid Input", "OK Icon!")
                this.HTTPTimeoutInput.Focus()
                return
            }

            ; Write to INI file FIRST
            local SettingsFile := A_ScriptDir . "\Settings.ini"

            IniWrite(NewInterval, SettingsFile, "Settings", "Interval")
            IniWrite(NewTestURL, SettingsFile, "Settings", "TestURL")
            IniWrite(NewVoiceAlerts ? 1 : 0, SettingsFile, "Settings", "VoiceAlerts")
            IniWrite(NewHTTPTimeout, SettingsFile, "Settings", "HTTPTimeout")

            ; Update properties AFTER successful file write
            this.Interval := NewInterval
            this.TestURL := NewTestURL
            this.VoiceAlerts := NewVoiceAlerts  ; Store as integer (0 or 1)
            this.HTTPTimeout := NewHTTPTimeout

            ; Update timer with new interval
            SetTimer(() => this.CheckConnection(), this.Interval)

            ; Close settings window
            this.SettingsGui.Destroy()

        } catch as e {
            this.Log("Error saving settings: " . e.Message, "ERROR")
        }
    }

    TestConnectionFunc(*) {
        local StartTime := A_TickCount
        local Result := this.CheckInternetHTTP()
        local Duration := A_TickCount - StartTime

        if (Result) {
            MsgBox("Internet connection successful! (" . Duration . "ms)", "Test Result", "OK Iconi")
        } else {
            MsgBox("Internet connection failed!", "Test Result", "OK Icon!")
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
        ; Fetch sync and update cache
        ;this.PublicIPCache := this.GetPublicIPFetch() ; async
        this.PublicIPCache := this.GetPublicIPSync()
        this.PublicIPLastFetch := A_TickCount
        return this.PublicIPCache
    }

    GetPublicIPSync() {
        try {
            req := ComObject("MSXML2.ServerXMLHTTP")
            req.open("GET", "https://api.ipify.org/", false)
            req.send()
            if (req.status == 200) {
                return Trim(req.responseText)
            } else {
                return "N/A"
            }
        } catch {
            return "N/A"
        }
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

    Speak(text) {
        try {
            ; 0 = SVSFlagsSync (blocking)
            this.gVoice.Speak(text, 0)
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
A_TrayMenu.Add("Reload", (*) => Reload())
A_TrayMenu.Add("Exit", (*) => ExitApp())
A_TrayMenu.Default := "Settings"

; Initialize with proper status
TraySetIcon("no-internet.ico", 1, true)

app.CheckConnection()
; Use a more robust timer that checks if previous check is still running
SetTimer(CheckConnectionTimer, app.Interval)

CheckConnectionTimer() {
    if (!app.IsChecking) {
        app.CheckConnection()
    } else {
        app.Log("Skipping scheduled check - previous check still running", "DEBUG")
    }
}
