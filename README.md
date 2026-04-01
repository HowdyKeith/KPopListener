# KPop Notification System v3.0 🎯

A comprehensive, enterprise-grade notification system for PowerShell with advanced features including plugin architecture, multiple transport methods, and rich toast notifications.

## 📋 Table of Contents

- [Features](#features)
- [Architecture](#architecture)
- [Installation](#installation)
- [Quick Start](#quick-start)
- [Modules](#modules)
- [Plugin System](#plugin-system)
- [Examples](#examples)
- [Configuration](#configuration)
- [API Reference](#api-reference)

---

## ✨ Features

### Core Features
- ✅ **Multi-engine toast notifications** (WinRT, BurntToast, Custom)
- ✅ **Plugin architecture** for extensibility
- ✅ **Multiple transport methods** (Named Pipes, File Watcher, WebSocket)
- ✅ **Advanced logging** with rotation and multiple targets
- ✅ **Real-time dashboard** with live statistics
- ✅ **Event-driven architecture** with custom event handlers
- ✅ **Performance tracking** and monitoring
- ✅ **Thread-safe operations** throughout

### Advanced Features
- 🔌 **Plugin Types**: Transport, Renderer, Integration
- 📊 **Statistics & Metrics**: Comprehensive tracking and reporting
- 🎨 **Toast Templates**: Pre-defined and custom templates
- 🔒 **Message validation**: CRC checksums, compression
- 🔄 **Auto-reconnection**: Resilient pipe connections
- 📝 **Structured logging**: Multiple log levels and targets
- 🎛️ **Live configuration**: Change settings at runtime

---

## 🏗️ Architecture

```
┌─────────────────────────────────────────────────────────┐
│                   KPopListener (Main)                    │
├─────────────────────────────────────────────────────────┤
│                                                           │
│  ┌──────────────┐  ┌──────────────┐  ┌──────────────┐  │
│  │  KPopCore    │  │  KPopLog     │  │ KPopToast    │  │
│  │  (State)     │  │  (Logging)   │  │ (Notify)     │  │
│  └──────────────┘  └──────────────┘  └──────────────┘  │
│                                                           │
│  ┌──────────────┐  ┌──────────────┐  ┌──────────────┐  │
│  │  KPopPipes   │  │ KPopPlugins  │  │ KPopDashboard│  │
│  │  (Transport) │  │ (Extensible) │  │ (GUI)        │  │
│  └──────────────┘  └──────────────┘  └──────────────┘  │
│                                                           │
└─────────────────────────────────────────────────────────┘
                            │
                ┌───────────┼───────────┐
                │           │           │
          ┌─────▼────┐ ┌───▼────┐ ┌───▼────┐
          │  Pipes   │ │  File  │ │WebSocket│
          └──────────┘ └────────┘ └─────────┘
```

---

## 📦 Installation

### Prerequisites
- **PowerShell 5.1** or **PowerShell 7+**
- **Windows 10/11** (for WinRT notifications)
- **BurntToast module** (optional, auto-installed)

### Steps

1. **Clone or download** the repository:
```powershell
git clone https://github.com/yourrepo/kpop-system.git
cd kpop-system
```

2. **Import modules**:
```powershell
Import-Module .\KPopCore.psm1
Import-Module .\KPopLog.psm1
Import-Module .\KPopToast.psm1
Import-Module .\KPopPipes.psm1
Import-Module .\KPopPlugins.psm1
Import-Module .\KPopDashboard.psm1
```

3. **Register AppID** (for proper branding):
```powershell
Register-AppID -AppId "KPop.Pop" -DisplayName "KPop Pop!" -CreateShortcut
```

---

## 🚀 Quick Start

### Basic Usage

```powershell
# 1. Initialize core
Initialize-KPopCore

# 2. Start logging
Start-KPopLog -Path "C:\Logs\kpop.log" -EnableConsole

# 3. Send a notification
Show-KPopNotification -Title "Hello!" -Message "KPop is running" -Type Info

# 4. Show dashboard
Show-KPopDashboard -AutoRefresh
```

### Listener Mode

```powershell
# Start the listener
.\KPopListener.ps1

# From another script/process
.\KPopSender.ps1 -Title "Build Complete" -Message "Success!" -Type SUCCESS
```

---

## 📚 Modules

### KPopCore.psm1 v3.0
**Core system with state management and events**

```powershell
# Initialize
Initialize-KPopCore

# Configuration
Set-KPopConfig -Name "MaxQueueSize" -Value 1000
Set-KPopConfig -Name "LogLevel" -Value "DEBUG"

# Statistics
$stats = Get-KPopStatistics
Write-Host "Total toasts: $($stats.TotalToasts)"

# Events
Register-KPopEventHandler -EventName "OnToastSent" -ScriptBlock {
    param($Data)
    Write-Host "Toast sent: $($Data.Title)"
}
```

**Key Functions:**
- `Initialize-KPopCore` / `Stop-KPopCore`
- `Set-KPopConfig` / `Get-KPopConfig`
- `Get-KPopStatistics` / `Reset-KPopStatistics`
- `Register-KPopEventHandler` / `Invoke-KPopEvent`
- `Add-KPopMessage` / `Get-KPopQueueStatus`

---

### KPopLog.psm1 v4.0
**Advanced logging with rotation and multiple targets**

```powershell
# Start logging
Start-KPopLog -Path "C:\Logs\app.log" -MinLevel INFO -EnableConsole

# Write logs
Write-KPopLog "Application started" -Level INFO
Write-KPopError "Something went wrong" -Category "Error"
Write-KPopDebug "Debug information" -Context @{ User = $env:USERNAME }

# Set context
Set-KPopLogContext -Context @{ Session = [guid]::NewGuid() }

# Query logs
Get-KPopLogEntries -Last 50 -Level ERROR
```

**Features:**
- **Multiple targets**: File, Console, Event Log, Custom
- **Log rotation**: Automatic file rotation with archival
- **Async mode**: Buffered writing for performance
- **Structured logging**: Attach context to entries
- **Log levels**: DEBUG, INFO, WARN, ERROR, FATAL

**Key Functions:**
- `Start-KPopLog` / `Stop-KPopLog`
- `Write-KPopLog` (and convenience: `Write-KPopInfo`, `Write-KPopError`, etc.)
- `Add-KPopLogTarget` / `Remove-KPopLogTarget`
- `Set-KPopLogContext` / `Clear-KPopLogContext`
- `Get-KPopLogStatistics` / `Get-KPopLogEntries`

---

### KPopToast.psm1 v5.0
**Rich toast notifications with templates**

```powershell
# Simple notification
Show-KPopNotification -Title "Success" -Message "Operation complete" -Type Success

# Progress notification
Update-KPopProgressToast -Tag "download" -Progress 75 -Status "Downloading..."

# Custom configuration
$config = [ToastConfig]::new()
$config.Title = "Custom Toast"
$config.Message = "With buttons"
$config.Buttons = @(
    @{ Content = "Action"; Arguments = "action=do" }
)
Send-KPopToast -Config $config

# Templates
Register-ToastTemplate -Name "MyTemplate" -Template $config
$template = Get-ToastTemplate -Name "MyTemplate"
```

**Features:**
- **Multi-engine**: WinRT (native) + BurntToast fallback
- **Templates**: Pre-defined and custom templates
- **Rich content**: Hero images, inline images, attribution
- **Actions**: Buttons, inputs, callbacks
- **Progress bars**: Real-time progress updates
- **Groups & tags**: Organize and update toasts

**Key Functions:**
- `Show-KPopNotification` (quick notifications)
- `Send-KPopToast` (full control)
- `Update-KPopProgressToast` / `Clear-KPopToast`
- `Register-ToastTemplate` / `Get-ToastTemplate`
- `Get-KPopToastHistory`

---

### KPopPipes.psm1 v3.0
**Advanced named pipe communication**

```powershell
# Server
$handler = {
    param($Message)
    Write-Host "Received: $($Message.Data)"
}

$config = [PipeConfig]::new()
$config.Protocol = [PipeProtocol]::Framed
$config.EnableCRC = $true

$server = Start-KPopPipeServer -Name "MyPipe" -MessageHandler $handler -Config $config

# Client
$client = New-KPopPipeClient -PipeName "MyPipe" -Config $config
$client.Connect()
$client.Send(@{ Title = "Hello"; Message = "World" })
$client.Disconnect()

# Statistics
Get-KPopPipeStatistics
Get-KPopPipeServers
```

**Features:**
- **Multiple protocols**: Raw, Framed (with headers), Chunked
- **Validation**: CRC32 checksums, magic headers
- **Compression**: Optional GZip compression
- **Auto-reconnect**: Resilient connections
- **Bidirectional**: Full duplex with ACK

**Key Functions:**
- `Start-KPopPipeServer` / `Stop-KPopPipeServer`
- `New-KPopPipeClient`
- `Get-KPopPipeStatistics` / `Get-KPopPipeServers`

---

### KPopPlugins.psm1 v1.0
**Extensible plugin architecture**

```powershell
# Create transport plugin
$myTransport = [TransportPlugin]@{
    Name = "CustomTransport"
    Version = "1.0"
}

$myTransport | Add-Member -MemberType ScriptMethod -Name Send -Value {
    param([hashtable]$Message)
    # Custom send logic
    return $true
} -Force

# Register
Register-KPopPlugin -Type Transport -Plugin $myTransport

# Use
Invoke-KPopTransport -TransportName "CustomTransport" -Message @{ Title = "Test" }

# Built-in integrations
$slack = New-SlackIntegrationPlugin -WebhookUrl "https://hooks.slack.com/..."
Register-KPopPlugin -Type Integration -Plugin $slack
Invoke-KPopIntegration -IntegrationName "Slack" -Data @{ Title = "Alert" }
```

**Plugin Types:**
1. **TransportPlugin**: Custom delivery mechanisms
2. **RendererPlugin**: Custom notification formats
3. **IntegrationPlugin**: External service integrations (Slack, Discord, Teams)

**Built-in Integrations:**
- Slack
- Discord
- Microsoft Teams

**Key Functions:**
- `Register-KPopPlugin` / `Unregister-KPopPlugin`
- `Get-KPopPlugin` / `Get-KPopPlugins`
- `Invoke-KPopTransport` / `Invoke-KPopRenderer` / `Invoke-KPopIntegration`
- `Import-KPopPluginFromFile` / `Import-KPopPluginsFromDirectory`

---

### KPopDashboard.psm1 v5.0
**Real-time monitoring and control GUI**

```powershell
# Launch dashboard
Show-KPopDashboard -AutoRefresh -RefreshInterval 1000
```

**Features:**
- 📊 **Real-time statistics**: Toasts, performance, memory
- ⚙️ **Live configuration**: Toggle settings on-the-fly
- 🎨 **Theme support**: Dark/Light modes
- 📬 **Test toasts**: Quick testing interface
- 💾 **Export data**: JSON/CSV exports
- 📋 **Activity log**: Recent events viewer
- 🔄 **Auto-refresh**: Configurable refresh rates

---

## 🔌 Plugin System

### Creating a Custom Transport Plugin

```powershell
# Define plugin
$customTransport = [TransportPlugin]@{
    Name = "EmailTransport"
    Version = "1.0.0"
    Author = "YourName"
    Description = "Send notifications via email"
}

# Implement Initialize
$customTransport | Add-Member -MemberType ScriptMethod -Name Initialize -Value {
    Write-Host "Email transport initialized"
} -Force

# Implement Send
$customTransport | Add-Member -MemberType ScriptMethod -Name Send -Value {
    param([hashtable]$Message)
    
    Send-MailMessage -To "admin@example.com" `
                     -Subject $Message.Title `
                     -Body $Message.Message `
                     -SmtpServer "smtp.example.com"
    
    return $true
} -Force

# Register
Register-KPopPlugin -Type Transport -Plugin $customTransport
```

### Creating an Integration Plugin

```powershell
$customIntegration = [IntegrationPlugin]@{
    Name = "CustomAPI"
    Version = "1.0.0"
    ServiceName = "My API"
    RequiresAuth = $true
}

$customIntegration.Authentication = @{
    ApiKey = "your-api-key"
    Endpoint = "https://api.example.com/notify"
}

$customIntegration | Add-Member -MemberType ScriptMethod -Name SendNotification -Value {
    param([hashtable]$Data)
    
    $headers = @{ 'X-API-Key' = $this.Authentication.ApiKey }
    $body = $Data | ConvertTo-Json
    
    Invoke-RestMethod -Uri $this.Authentication.Endpoint `
                     -Method Post `
                     -Headers $headers `
                     -Body $body `
                     -ContentType 'application/json'
    
    return $true
} -Force

Register-KPopPlugin -Type Integration -Plugin $customIntegration
```

### Loading Plugins from Files

Create a plugin file (`MyPlugin.plugin.ps1`):
```powershell
# Return plugin instance
[TransportPlugin]@{
    Name = "FileBasedTransport"
    Version = "1.0"
}
```

Load it:
```powershell
Import-KPopPluginFromFile -Path ".\MyPlugin.plugin.ps1"
# Or load all from directory
Import-KPopPluginsFromDirectory -Path ".\Plugins" -Recurse
```

---

## 📖 Examples

### Example 1: Simple Notification Script

```powershell
Import-Module .\KPopCore.psm1
Import-Module .\KPopToast.psm1

Initialize-KPopCore
Show-KPopNotification -Title "Backup Complete" -Message "Your data has been backed up" -Type Success
```

### Example 2: Progress Tracking

```powershell
$tasks = 1..100

foreach ($i in $tasks) {
    # Do work
    Start-Sleep -Milliseconds 50
    
    # Update progress
    $progress = ($i / $tasks.Count) * 100
    Update-KPopProgressToast -Tag "work-progress" -Progress $progress -Status "Processing task $i"
}

Clear-KPopToast -Tag "work-progress"
Show-KPopNotification -Title "Complete" -Message "All tasks finished" -Type Success
```

### Example 3: Full System with Logging

```powershell
# Initialize
Initialize-KPopCore
Start-KPopLog -Path "C:\Logs\app.log" -EnableConsole

# Configure
Set-KPopConfig -Name "LogLevel" -Value "DEBUG"

# Add event handler
Register-KPopEventHandler -EventName "OnError" -ScriptBlock {
    param($Data)
    Write-KPopError "Error occurred: $($Data.Message)"
}

# Your application logic
try {
    Write-KPopInfo "Starting process"
    
    # ... do work ...
    
    Show-KPopNotification -Title "Success" -Message "Process complete" -Type Success
    Write-KPopInfo "Process completed successfully"
    
} catch {
    Invoke-KPopEvent -EventName "OnError" -Data @{ Message = $_.Exception.Message }
    Show-KPopNotification -Title "Error" -Message $_.Exception.Message -Type Error
} finally {
    Stop-KPopLog
    Stop-KPopCore
}
```

### Example 4: Multi-Channel Notifications

```powershell
# Setup integrations
$slack = New-SlackIntegrationPlugin -WebhookUrl $env:SLACK_WEBHOOK
$discord = New-DiscordIntegrationPlugin -WebhookUrl $env:DISCORD_WEBHOOK

Register-KPopPlugin -Type Integration -Plugin $slack
Register-KPopPlugin -Type Integration -Plugin $discord

# Send to all channels
$message = @{
    Title = "🚨 Alert"
    Message = "Critical system event detected"
    Type = "ERROR"
}

# Windows toast
Show-KPopNotification -Title $message.Title -Message $message.Message -Type Error

# Slack
Invoke-KPopIntegration -IntegrationName "Slack" -Data $message

# Discord
Invoke-KPopIntegration -IntegrationName "Discord" -Data $message
```

---

## ⚙️ Configuration

### Core Configuration Options

```powershell
Set-KPopConfig -Name "LogLevel" -Value "DEBUG"              # Log verbosity
Set-KPopConfig -Name "MaxQueueSize" -Value 1000             # Message queue size
Set-KPopConfig -Name "EnablePerformanceTracking" -Value $true  # Track metrics
Set-KPopConfig -Name "AutoCleanup" -Value $true             # Auto GC
Set-KPopConfig -Name "CleanupThresholdMB" -Value 100        # GC threshold
```

### Pipe Configuration

```powershell
$pipeConfig = [PipeConfig]::new()
$pipeConfig.Protocol = [PipeProtocol]::Framed  # Raw, Framed, or Chunked
$pipeConfig.EnableCRC = $true                   # Checksum validation
$pipeConfig.EnableCompression = $true           # GZip compression
$pipeConfig.AutoReconnect = $true               # Auto-reconnect on failure
$pipeConfig.MaxReconnectAttempts = 3            # Max reconnect tries
$pipeConfig.TimeoutMs = 5000                    # Connection timeout
```

### Log Configuration

```powershell
$script:LogConfig.MinLevel = 'INFO'             # Minimum log level
$script:LogConfig.MaxBufferSize = 1000          # Buffer size for async mode
$script:LogConfig.AsyncMode = $true             # Enable async logging
$script:LogConfig.MaxFileSize = 10MB            # Max size before rotation
$script:LogConfig.MaxArchiveFiles = 10          # Number of archives to keep
```

---

## 🔍 API Reference

### Core Functions

| Function | Description |
|----------|-------------|
| `Initialize-KPopCore` | Initialize the core system |
| `Stop-KPopCore` | Stop and cleanup core |
| `Set-KPopConfig` | Update configuration |
| `Get-KPopConfig` | Get configuration value |
| `Get-KPopStatistics` | Get system statistics |
| `Register-KPopEventHandler` | Register event callback |
| `Invoke-KPopEvent` | Fire an event |

### Logging Functions

| Function | Description |
|----------|-------------|
| `Start-KPopLog` | Start logging |
| `Stop-KPopLog` | Stop logging |
| `Write-KPopLog` | Write log entry |
| `Write-KPopInfo/Debug/Warn/Error/Fatal` | Convenience functions |
| `Get-KPopLogStatistics` | Get logging stats |
| `Add-KPopLogTarget` | Add custom log target |

### Toast Functions

| Function | Description |
|----------|-------------|
| `Show-KPopNotification` | Quick notification |
| `Send-KPopToast` | Send with full config |
| `Update-KPopProgressToast` | Update progress |
| `Clear-KPopToast` | Clear notification |
| `Register-ToastTemplate` | Register template |
| `Get-KPopToastHistory` | Get notification history |

### Pipe Functions

| Function | Description |
|----------|-------------|
| `Start-KPopPipeServer` | Start pipe server |
| `Stop-KPopPipeServer` | Stop pipe server |
| `New-KPopPipeClient` | Create pipe client |
| `Get-KPopPipeStatistics` | Get pipe stats |

### Plugin Functions

| Function | Description |
|----------|-------------|
| `Register-KPopPlugin` | Register plugin |
| `Unregister-KPopPlugin` | Unregister plugin |
| `Get-KPopPlugin` | Get plugin by name |
| `Get-KPopPlugins` | List all plugins |
| `Invoke-KPopTransport` | Use transport plugin |
| `Invoke-KPopIntegration` | Use integration plugin |

---

## 🎯 Best Practices

1. **Always initialize core first**: Call `Initialize-KPopCore` before other operations
2. **Use structured logging**: Attach context to log entries with `Set-KPopLogContext`
3. **Handle cleanup**: Always call `Stop-KPopCore` and `Stop-KPopLog` on exit
4. **Use templates**: Define reusable toast templates for consistency
5. **Monitor performance**: Enable performance tracking and review statistics
6. **Plugin isolation**: Keep plugin logic separate and handle errors gracefully
7. **Dashboard for debugging**: Use the dashboard during development

---

## 🐛 Troubleshooting

### Toasts Not Showing
- Verify AppID registration: `Test-AppIDRegistered -AppId "KPop.Pop"`
- Check toast engine: `Get-PreferredToastEngine`
- Review logs for errors

### Pipe Connection Failures
- Verify pipe name matches between server/client
- Check for existing pipe with same name: `Get-NamedPipe`
- Enable auto-reconnect in PipeConfig

### Memory Issues
- Lower `MaxQueueSize` configuration
- Enable `AutoCleanup` and adjust `CleanupThresholdMB`
- Review log rotation settings

---

## 📄 License

MIT License - See LICENSE file for details

---

## 🤝 Contributing

Contributions welcome! Please:
1. Fork the repository
2. Create a feature branch
3. Submit a pull request

---

## 📞 Support

- **Issues**: GitHub Issues
- **Documentation**: This README
- **Examples**: See `Examples` directory

---

**Made with ❤️ for PowerShell enthusiasts**