# KPopListener

**KPop Notification System v3.0**  
A comprehensive, enterprise-grade notification toolkit for **PowerShell** with strong **VBA (Excel/Access)** integration.

It delivers rich toast notifications, multiple backends, advanced logging, a real-time dashboard, plugin support, and fast inter-process communication via **named pipes**.

## Features

### Core Features
- Multi-engine toast notifications (WinRT, BurntToast, Custom templates)
- Plugin architecture for extensibility (Transport, Renderer, Integration)
- Multiple transport methods: **Named Pipes**, File Watcher, WebSocket
- Advanced logging with rotation and multiple targets
- Real-time dashboard with live statistics
- Event-driven architecture and callbacks
- Performance tracking and thread-safe operations

### VBA Integration Highlights
- Unified message box system with many backends (Toast, WinRT, MSHTA, Python, PushBullet, PowerShell)
- Toast sender, progress toasts, queue management, analytics
- Setup, logging, credentials, Ribbon UI, and Push2Run support

## Architecture Overview

The system is modular:
- **KPopCore** – State and events
- **KPopLog** – Advanced logging
- **KPopToast** – Notification rendering
- **KPopPipes** – Named pipe communication
- **KPopPlugins** & **KPopDashboard** – Extensibility and monitoring

## Installation & Setup

### PowerShell
1. Clone or download the repository.
2. Import the required modules:
   ```powershell
   Import-Module .\KPopCommon.psm1
   Import-Module .\KPopLog.psm1
   Import-Module .\KPopToast.psm1
   Import-Module .\KPopPipes.psm1
   Import-Module .\PipeManager.psm1

(Optional) Register AppID for branded toasts:PowerShell.\AppIDTestHarness.ps1

VBA (Excel / Access)

Open your file → Alt + F11
Import key modules: Setup.bas, MsgBoxUnified.bas, ToastSender.bas, ToastMasterDemo.bas, DiagnosePowerShellListener.bas, and relevant .cls files (e.g., clsToastProgress.cls, clsCallbacks.cls).
Run Setup once to configure registry and initial settings.

Detailed VBA Examples
1. Simple Toast
vbaSub ShowSimpleToast()
    Call SendToast("KPopListener", "Hello from VBA!", "This is a basic notification.")
End Sub
2. Unified Message Box (Recommended)
vbaSub TestUnifiedMsgBox()
    Dim Response As VbMsgBoxResult
    Response = MsgBoxUnified("Do you want to continue?", vbYesNo + vbQuestion, "KPopListener", "Toast")
    
    If Response = vbYes Then
        Call MsgBoxUnified("You selected Yes!", vbInformation)
    End If
End Sub
3. Progress Toast
vbaSub ShowProgressExample()
    Dim prog As clsToastProgress
    Set prog = New clsToastProgress
    prog.Title = "Processing Data"
    prog.Message = "Please wait..."
    prog.Show
    
    Dim i As Integer
    For i = 10 To 100 Step 10
        prog.UpdateProgress i, "Step " & i & "% complete"
        Application.Wait Now + TimeValue("0:00:01")
    Next i
    
    prog.Complete "All tasks finished successfully!"
End Sub
4. Run Demo Files
Import and run ToastMasterDemo.bas or ToastExample.bas for ready-made examples.
PowerShell Listener (Named Pipes)
The PowerShell Listener provides fast, persistent, bidirectional communication between VBA and PowerShell using named pipes. Instead of launching a new PowerShell process for every toast, the listener runs in the background and receives commands efficiently.
Key Files

KPopListener.ps1 – Starts the named pipe server
PipeManager.psm1 & KPopPipes.psm1 – Pipe handling with framed protocol, CRC validation, compression, and auto-reconnect
NamedPipeSender.ps1 & KPopSender.ps1 – Client tools to send messages
DiagnosePowerShellListener.bas – VBA diagnostic and integration module
clsCallbacks.cls – Handles callbacks and events from PowerShell back to VBA

How It Works

Start the listener (KPopListener.ps1).
VBA (or another client) connects to the named pipe and sends structured messages.
The listener processes the request (shows toast, logs, etc.) and can send responses/callbacks back.

PowerShell Examples
Start the Listener:
PowerShell.\KPopListener.ps1
Send a Toast via Sender:
PowerShell.\KPopSender.ps1 -Title "Build Complete" -Message "Success!" -Type SUCCESS
Advanced Pipe Usage (from KPopPipes.psm1):
PowerShell$config = [PipeConfig]::new()
$config.Protocol = [PipeProtocol]::Framed
$config.EnableCRC = $true
$config.EnableCompression = $true

$server = Start-KPopPipeServer -Name "KPopListener" -MessageHandler $handler -Config $config
Improved VBA Integration with the Listener
vba' 1. Start the listener from VBA (hidden)
Sub StartKPopListener()
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")
    shell.Run "powershell.exe -ExecutionPolicy Bypass -File ""KPopListener.ps1"" -WindowStyle Hidden", 0, False
End Sub

' 2. Diagnose / Test connection
Sub DiagnoseListener()
    Call DiagnosePowerShellListener   ' Uses DiagnosePowerShellListener.bas
End Sub

' 3. Send a toast via named pipe (example using helper logic)
Sub SendToastViaListener()
    ' You can extend ToastSender.bas or MsgBoxToastsPS.bas
    ' to open the pipe and send JSON-like messages such as:
    ' { "command": "showToast", "title": "KPopListener", "message": "Sent via named pipe!", "duration": "Long" }
    
    Call MsgBoxUnified("Toast sent via listener!", vbInformation)
End Sub

' 4. Example with callback handling (using clsCallbacks.cls)
Sub SetupCallback()
    Dim cb As clsCallbacks
    Set cb = New clsCallbacks
    ' Register callback for toast events, user actions, etc.
End Sub
Tip: Use DiagnosePowerShellListener.bas first to check status and automatically start the listener if needed.
Quick Start (PowerShell)
PowerShellInitialize-KPopCore
Start-KPopLog -Path "C:\Logs\kpop.log"
Show-KPopNotification -Title "Hello!" -Message "KPop is running" -Type Info
Show-KPopDashboard -AutoRefresh
Contributing
Contributions welcome! Improve VBA integration, add new plugins, enhance pipe reliability, or expand documentation.
License
MIT License (see LICENSE if present).
