<#
.SYNOPSIS
KPopListener.ps1 v26.0 PA8 — PS5.1 + PS7.5+ FULLY COMPATIBLE

.DESCRIPTION
- Named pipe (byte-mode with ACK / optional raw stream)
- RaiseEventMode support (async event-based pipe listener)
- File watcher fallback
- WinRT native toasts + BurntToast fallback
- BurntToast used for startup progress & Ready link toast
- HTTP dashboard with live controls for UseRawStream & RaiseEventMode
- Custom URI protocol registration (kpop:) for callback buttons
#>

[CmdletBinding()]
param(
[switch]$Background,
[switch]$NoStartupToast,
[switch]$NoDashboard,
[int]$Port = 0,
[int]$WebSocketPort = 8080,
[switch]$EnableWebSocket,
[switch]$RegisterProtocol
)

# ---- Constants (can be overridden by dashboard) ----

$global:UseRawStream = $true
$global:RaiseEventMode = $false

# ---- Directories / Temp Files ----

$ScriptVersion    = "26.0 PA8"
$AppID            = "KPop.Pop"
$AppDisplayName   = "KPop Pop!"
$TempDir          = "$env:TEMP\KPopListener"
$PidFile          = "$TempDir\KPopListener.pid"
$StatusFile       = "$TempDir\ListenerStatus.json"
$RequestLog       = "$TempDir\ToastRequests.log"
$StartConfirmFile = "$TempDir\start.confirmed"
$ForceBTFile      = "$TempDir\forcebt.flag"

if (-not (Test-Path $TempDir)) { New-Item -Path $TempDir -ItemType Directory -Force | Out-Null }

# ---- Global State ----

$global:Stats = @{ StartTime = Get-Date; TotalToasts=0; PipeToasts=0; FileToasts=0; Failures=0; LastToastTitle="None"; LastToastTime=$null; Version=$ScriptVersion }
$global:IsRunning = $true
$global:MessageQueue = [System.Collections.Concurrent.ConcurrentQueue[object]]::new()
$global:BasePipeName = "KPopListenerPipe_$PID"
$global:WinRTAvailable = $false
$global:BurntToastAvailable = $false
$global:EngineSync = [hashtable]::Synchronized(@{ UseWinRT = $true })
$global:UseWinRT = $true

# ---- Load Common Functions ----

$commonModule = Join-Path $PSScriptRoot "KPopCommon.psm1"
if (Test-Path $commonModule) { Import-Module $commonModule -Force }

# ---- Load Pipes Module ----

$pipesModule = Join-Path $PSScriptRoot "KPopPipes.psm1"
if (Test-Path $pipesModule) { Import-Module $pipesModule -Force }

# ---- Load Dashboard ----

$dashboardModule = Join-Path $PSScriptRoot "KPopDashboard.psm1"
if (-not $NoDashboard -and Test-Path $dashboardModule) { Import-Module $dashboardModule -Force }

# ---- Toast Initialization ----

$global:WinRTAvailable = try { $null = [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime]; $true } catch { $false }
try { Import-Module BurntToast -ErrorAction Stop; $global:BurntToastAvailable = $true } catch {}

# ---- Protocol registration ----

if ($RegisterProtocol.IsPresent) { Register-KPopProtocol }

# ---- Register AppID ----

Register-KPopAppID

# ---- Dashboard helpers ----

function Apply-DashboardSettings {
# Called at startup and after live toggle
if ($global:RaiseEventMode) {
if (Get-Command Start-PipeListenerAsync -ErrorAction SilentlyContinue) {
$global:PipeHandle = Start-PipeListenerAsync -PipeName $global:BasePipeName -UseRawStream:$global:UseRawStream
Write-Log "Pipe listener started in event mode (RaiseEventMode=True)" "SUCCESS"
}
} else {
if (Get-Command Start-PipeListener -ErrorAction SilentlyContinue) {
$global:PipeHandle = Start-PipeListener -PipeName $global:BasePipeName -UseRawStream:$global:UseRawStream
Write-Log "Pipe listener started in runspace loop mode (RaiseEventMode=False)" "SUCCESS"
}
}
}

# ---- Start listener based on initial constants ----

Apply-DashboardSettings

# ---- File watcher (optional fallback) ----

if (Get-Command Start-FileWatcher -ErrorAction SilentlyContinue) {
$global:FileWatcher = Start-FileWatcher
Write-Log "File watcher started" "SUCCESS"
}

# ---- Dashboard GUI ----

if (-not $NoDashboard) { Start-KPopDashboard -AutoRefresh }

# ---- Show startup progress ----

if (-not $NoStartupToast -and $global:BurntToastAvailable) {
Show-StartupProgress -Steps 4 -DelayMs 400
if ($global:WinRTAvailable -and $global:EngineSync.UseWinRT) { Show-WinRTSpinner -DurationSec 2 }
}

# ---- Ready toast ----

if (-not $NoStartupToast) {
if (-not $NoDashboard -and $global:DashboardPort -ne "Disabled" -and $global:BurntToastAvailable) {
Send-BurntToastNotif -Title "$AppDisplayName Ready!" -Message "Click to open dashboard." -ButtonText "Open Dashboard" -ButtonUrl $global:DashboardUrl
} else { Send-Toast -Title "$AppDisplayName Ready!" -Message "All systems go!" }
}

# ---- Main loop with spinner ----

$spin = @("⣾","⣽","⣻","⢿","⡿","⣟","⣯","⣷")
$i = 0
$last = Get-Date
$ControlFile = Join-Path $TempDir "KPopControl.txt"

while ($global:IsRunning) {
# Process queued messages
$m = $null
while ($global:MessageQueue.TryDequeue([ref]$m)) { try { Process-Message -msg $m } catch {} }

```
if (-not $Background) {
    $e = if ($global:EngineSync.UseWinRT) { "WinRT" } else { "BT" }
    $txt = "Listening... [$e] (T:$($global:Stats.TotalToasts)|P:$($global:Stats.PipeToasts)|F:$($global:Stats.FileToasts)) $($spin[$i])"
    Write-Host "`r$txt " -NoNewline -ForegroundColor Cyan
    $i = ($i + 1) % $spin.Count
}

# Poll control file for protocol commands
if (Test-Path $ControlFile) {
    try {
        $cmd = (Get-Content $ControlFile -Raw).Trim().ToLower()
        switch ($cmd) {
            'pause'  { $global:IsRunning = $false; Write-Log "Runtime command: pause" "INFO" }
            'resume' { $global:IsRunning = $true; Write-Log "Runtime command: resume" "INFO" }
            'exit'   { Write-Log "Runtime command: exit" "INFO"; break }
        }
    } catch {}
    try { Remove-Item $ControlFile -Force } catch {}
}

# Refresh dashboard settings every 5 sec
if (((Get-Date) - $last).TotalSeconds -ge 5) {
    Apply-DashboardSettings
    Update-ListenerStatus
    $last = Get-Date
}

Start-Sleep -Milliseconds 200
```

}

# ---- Cleanup ----

try { if ($global:PipeHandle -and ($global:PipeHandle -is [System.IDisposable])) { $global:PipeHandle.Dispose() } } catch {}
try { if ($global:FileWatcher -and ($global:FileWatcher -is [System.Management.Automation.PSCustomObject])) {} } catch {}

Write-Log "Listener stopped." "INFO"
