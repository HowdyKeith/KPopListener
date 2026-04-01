# File: v7/KPopupClipSaver.ps1
# Path: v7/KPopupClipSaver.ps1
# ScriptVersion: 7.66
# ModuleVersion: 7.66

<#
.SYNOPSIS
KPopupClipSaver standalone — monitors clipboard and auto-saves scripts/assets with version-aware overwrite,
tray icon, dashboard (WPF), TTS feedback, shortcuts, and backups.

.DESCRIPTION
- Detects # File: filename.ext or # Path: vX\filename.ext header lines in clipboard text.
- Valid extensions: ps1, psm1, py, js, css, vbs, bat, bas, frm, cls, png, jpg, jpeg, mp4
- Saves files into version folder (default latest v* folder) under this project's base folder.
- For psm1 saves, uses Modules/ subfolder automatically.
- Does version-aware overwrite: reads # ModuleVersion: or # ScriptVersion: from existing file (first 10 lines)
  and only writes if incoming version is greater.
- Creates backups in Backup/ folder (timestamped) before overwriting.
- Optional creation of .lnk shortcuts for configured extensions.
- Plays TTS/audio feedback on saves, and shows Tray icon + Dashboard.

.NOTES
- Requires PowerShell 5.1+ (desktop) for WPF/WinForms UI and System.Speech.
- Run: powershell -ExecutionPolicy Bypass -File .\KPopupClipSaver.ps1
#>

[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Continue'

# =====================================================================
# Configuration / State
# =====================================================================

$Script:ScriptVersion = '7.66'
$ScriptLocation = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$ConfigFile = Join-Path $ScriptLocation 'KPopupClipSaver.ini'
$DefaultCheckIntervalMs = 700
$MinSaveThrottleSeconds = 10 # minimum seconds between successive saves of the same target file
$MaxHistory = 100
$ValidExtensions = @('ps1','psm1','py','js','css','vbs','bat','bas','frm','cls','png','jpg','jpeg','mp4')
$ShortcutExtsDefault = @('ps1','bat') # create shortcuts for scripts by default

# State containers
$Script:LastClipHash = ''
$Script:LastImageHash = ''
$Script:LastSavedTimes = @{} # path -> DateTime of last save (throttling)
$Script:SavedHistory = [System.Collections.ObjectModel.ObservableCollection[psobject]]::new()
$Script:TrayIcon = $null
$Script:DashboardWindow = $null
$Script:PlaySounds = $true
$Script:SelectedVoice = $null
$Script:SaveCount = 0

# =====================================================================
# Read/Write Config
# =====================================================================

function Read-Config {
    if (-not (Test-Path $ConfigFile)) { 
        return @{ BaseFolder = $ScriptLocation; PlaySounds = 'True'; CreateShortcuts = 'ps1,bat' } 
    }
    $cfg = @{}
    Get-Content $ConfigFile | ForEach-Object {
        $l = $_.Trim()
        if ($l -and $l -notmatch '^(#|;)' -and $l -match '^(.+?)=(.+)$') { 
            $cfg[$matches[1].Trim()] = $matches[2].Trim() 
        }
    }
    if (-not $cfg.BaseFolder) { $cfg.BaseFolder = $ScriptLocation }
    return $cfg
}

$cfg = Read-Config
$BaseFolder = if ($cfg.BaseFolder -and (Test-Path $cfg.BaseFolder)) { $cfg.BaseFolder } else { $ScriptLocation }
$Script:PlaySounds = ($cfg.PlaySounds -eq 'True')
$ShortcutExts = if ($cfg.ContainsKey('CreateShortcuts') -and $cfg.CreateShortcuts) { 
    ($cfg.CreateShortcuts -split ',') | ForEach-Object { $_.Trim().ToLower() } 
} else { 
    $ShortcutExtsDefault 
}

# Ensure at least one version folder exists
function Get-LatestVersionFolderName {
    $versions = Get-ChildItem -Path $BaseFolder -Directory -ErrorAction SilentlyContinue | 
        Where-Object { $_.Name -match '^v(\d+)(\.\d+)?$' }
    if (-not $versions) { return 'v1' }
    $sorted = $versions | Sort-Object { 
        if ($_.Name -match '^v(\d+)') { [int]$matches[1] } else { 0 }
    } -Descending
    return $sorted[0].Name
}
$DefaultVersion = Get-LatestVersionFolderName

# =====================================================================
# Utility helpers
# =====================================================================

function Get-ClipboardHash([string]$Text) {
    if (-not $Text) { return '' }
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($Text)
        $ms = New-Object System.IO.MemoryStream(,$bytes)
        $h = (Get-FileHash -InputStream $ms -Algorithm SHA256).Hash
        $ms.Dispose()
        return $h
    } catch { return '' }
}

function Get-ImageHash([System.Drawing.Image]$Img) {
    if (-not $Img) { return '' }
    try {
        $ms = New-Object System.IO.MemoryStream
        $Img.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
        $ms.Position = 0
        $h = (Get-FileHash -InputStream $ms -Algorithm SHA256).Hash
        $ms.Close()
        return $h
    } catch { return '' }
}

function Normalize-PathForSave([string]$root, [string]$ver, [string]$file) {
    $root = Join-Path $root $ver
    if ($file -match '[/\\]') {
        $full = Join-Path $root $file
        $dir = Split-Path $full -Parent
        if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        return $full
    } else {
        $ext = ([IO.Path]::GetExtension($file) -replace '^\.' , '').ToLower()
        if ($ext -eq 'psm1') {
            $moddir = Join-Path $root 'Modules'
            if (-not (Test-Path $moddir)) { New-Item -ItemType Directory -Path $moddir -Force | Out-Null }
            return Join-Path $moddir $file
        } else {
            if (-not (Test-Path $root)) { New-Item -ItemType Directory -Path $root -Force | Out-Null }
            return Join-Path $root $file
        }
    }
}

function Read-Top-Versions([string]$FilePath) {
    if (-not (Test-Path $FilePath)) { return $null }
    try {
        $lines = Get-Content -Path $FilePath -TotalCount 16 -ErrorAction SilentlyContinue
        foreach ($ln in $lines) {
            if ($ln -match '#\s*(ModuleVersion|ScriptVersion)\s*[:=]\s*"?([0-9a-zA-Z.-]+)"?') { 
                return $matches[2] 
            }
        }
    } catch {}
    return $null
}

function Backup-File([string]$Path) {
    try {
        if (-not (Test-Path $Path)) { return }
        $bakdir = Join-Path (Split-Path $Path -Parent) 'Backup'
        if (-not (Test-Path $bakdir)) { New-Item -ItemType Directory -Path $bakdir -Force | Out-Null }
        $name = [IO.Path]::GetFileNameWithoutExtension($Path)
        $ext = [IO.Path]::GetExtension($Path)
        $stamp = (Get-Date).ToString('yyyyMMdd_HHmmss')
        $bakname = "{0}-BAK_{1}{2}" -f $name, $stamp, $ext
        Copy-Item -Path $Path -Destination (Join-Path $bakdir $bakname) -Force
        return (Join-Path $bakdir $bakname)
    } catch { 
        Write-Warning "Backup failed: $_"
        return $null 
    }
}

function New-FileShortcut([string]$Path) {
    try {
        $ext = ([IO.Path]::GetExtension($Path) -replace '^\.' , '').ToLower()
        if ($ShortcutExts -notcontains $ext) { return }
        $sh = New-Object -ComObject WScript.Shell
        $lnk = "$Path.lnk"
        $s = $sh.CreateShortcut($lnk)
        $s.TargetPath = $Path
        $s.WorkingDirectory = Split-Path $Path
        $s.Save()
        [Runtime.InteropServices.Marshal]::ReleaseComObject($sh) | Out-Null
    } catch { 
        Write-Warning "Shortcut creation failed: $_" 
    }
}

# =====================================================================
# TTS / Sound
# =====================================================================

try {
    Add-Type -AssemblyName System.Speech -ErrorAction SilentlyContinue
    $Script:TTS = [System.Speech.Synthesis.SpeechSynthesizer]::new()
    $Script:TTS.Volume = 100
    $Script:TTS.Rate = 0
} catch {
    $Script:TTS = $null
}

function Speak-Message([string]$Text, [switch]$Async) {
    if (-not $Script:PlaySounds -or -not $Script:TTS) { return }
    try {
        if ($Async) { 
            $Script:TTS.SpeakAsync($Text) | Out-Null 
        } else { 
            $Script:TTS.Speak($Text) 
        }
    } catch {}
}

function Play-SystemSound([string]$Type='Asterisk') {
    if (-not $Script:PlaySounds) { return }
    try {
        switch ($Type) {
            'Asterisk' { [System.Media.SystemSounds]::Asterisk.Play() }
            'Beep' { [System.Media.SystemSounds]::Beep.Play() }
            'Hand' { [System.Media.SystemSounds]::Hand.Play() }
            'Exclamation' { [System.Media.SystemSounds]::Exclamation.Play() }
            default { [System.Media.SystemSounds]::Beep.Play() }
        }
    } catch {}
}

# =====================================================================
# Tray Icon & Dashboard (WinForms + WPF)
# =====================================================================

Add-Type -AssemblyName System.Windows.Forms, System.Drawing -ErrorAction SilentlyContinue
Add-Type -AssemblyName PresentationCore, PresentationFramework, WindowsBase -ErrorAction SilentlyContinue

function Initialize-TrayIcon {
    # Creates NotifyIcon and context menu
    $notify = New-Object System.Windows.Forms.NotifyIcon
    $bmp = New-Object System.Drawing.Bitmap 32,32
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.Clear([System.Drawing.Color]::Transparent)
    $brush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(180,76,175,80))
    $g.FillEllipse($brush,4,4,24,24)
    $g.Dispose()
    $notify.Icon = [System.Drawing.Icon]::FromHandle($bmp.GetHicon())
    $notify.Visible = $true
    $notify.Text = "KPopupClipSaver v$($Script:ScriptVersion)"
    
    $menu = New-Object System.Windows.Forms.ContextMenuStrip
    $miDashboard = $menu.Items.Add('Open Dashboard')
    $miDashboard.add_Click({ Show-Dashboard })
    $miConfig = $menu.Items.Add('Open Base Folder')
    $miConfig.add_Click({ Start-Process explorer.exe -ArgumentList $BaseFolder })
    $miQuit = $menu.Items.Add('Exit')
    $miQuit.add_Click({ Stop-ClipSaver })
    $notify.ContextMenuStrip = $menu
    $notify.add_DoubleClick({ Show-Dashboard })
    $Script:TrayIcon = $notify
}

function Update-TrayTooltip {
    if ($Script:TrayIcon) { 
        $Script:TrayIcon.Text = "KPopupClipSaver v$($Script:ScriptVersion)`nSaved: $($Script:SaveCount) files" 
    }
}

# =====================================================================
# Dashboard (simple WPF window)
# =====================================================================

function Show-Dashboard {
    if ($Script:DashboardWindow -and $Script:DashboardWindow.IsVisible) {
        $Script:DashboardWindow.Activate()
        return
    }

    $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" 
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="KPopupClipSaver" 
        Height="640" Width="900" 
        WindowStartupLocation="CenterScreen" 
        FontFamily="Segoe UI" 
        Background="#121212">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>

        <StackPanel Orientation="Horizontal" Grid.Row="0" HorizontalAlignment="Left" Margin="6">
            <TextBlock Text="KPopupClipSaver" Foreground="White" FontSize="18" FontWeight="SemiBold" Margin="6"/>
            <TextBlock Text="v$($Script:ScriptVersion)" Foreground="#B0B0B0" VerticalAlignment="Center" Margin="4,0"/>
        </StackPanel>

        <StackPanel Orientation="Horizontal" Grid.Row="1" Margin="6" HorizontalAlignment="Left">
            <Button x:Name="BtnClear" Width="110" Margin="4">Clear History</Button>
            <Button x:Name="BtnOpenBase" Width="140" Margin="4">Open Base Folder</Button>
            <Button x:Name="BtnPlay" Width="120" Margin="4">Toggle Sounds</Button>
            <Button x:Name="BtnStop" Width="110" Margin="4">Stop</Button>
        </StackPanel>

        <ListBox x:Name="History" Grid.Row="2" Margin="6" Background="#1E1E1E" Foreground="White">
            <ListBox.ItemTemplate>
                <DataTemplate>
                    <Grid Margin="6">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="120"/>
                        </Grid.ColumnDefinitions>
                        <TextBlock Text="{Binding File}" Foreground="White" FontWeight="SemiBold"/>
                        <TextBlock Grid.Column="1" Text="{Binding Time}" Foreground="#B0B0B0" HorizontalAlignment="Right"/>
                    </Grid>
                </DataTemplate>
            </ListBox.ItemTemplate>
        </ListBox>
    </Grid>
</Window>
"@

    $reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]$xaml)
    $win = [Windows.Markup.XamlReader]::Load($reader)
    $Script:DashboardWindow = $win

    # Get controls
    $btnClear = $win.FindName('BtnClear')
    $btnOpen = $win.FindName('BtnOpenBase')
    $btnPlay = $win.FindName('BtnPlay')
    $btnStop = $win.FindName('BtnStop')
    $history = $win.FindName('History')

    $history.ItemsSource = $Script:SavedHistory

    $btnClear.Add_Click({
        $Script:SavedHistory.Clear()
    })
    $btnOpen.Add_Click({
        Start-Process explorer.exe -ArgumentList $BaseFolder
    })
    $btnPlay.Add_Click({
        $Script:PlaySounds = -not $Script:PlaySounds
        $btnPlay.Content = if ($Script:PlaySounds) { 'Sounds: ON' } else { 'Sounds: OFF' }
    })
    $btnStop.Add_Click({
        Stop-ClipSaver
    })

    $win.Show()
}

# =====================================================================
# Core Save logic
# =====================================================================

function Save-ClipboardFile([string]$Content) {
    if (-not $Content) { return }

    try {
        # Look for # File: or # Path:
        if ($Content -notmatch "(?m)^[#;]\s*(File|Path)\s*[:=]\s*(.+)$") { return }
        $pathHeader = $matches[2].Trim()
        $verFolder = $DefaultVersion
        $fileName = $pathHeader

        if ($pathHeader -match '^(v[\d\.]+)[/\\](.+)$') {
            $verFolder = $matches[1]
            $fileName = $matches[2]
        } elseif ($pathHeader -match '^(v[\d\.]+)$') {
            $verFolder = $matches[1]
            Write-Verbose "Header only contained folder $verFolder - skipping save"
            return
        }

        # Safely get extension
        $ext = ''
        try {
            $ext = ([IO.Path]::GetExtension($fileName) -replace '^\.' , '').ToLower()
        } catch {
            Write-Verbose "Invalid filename format: $fileName"
            return
        }
        
        if (-not $ext -or $ValidExtensions -notcontains $ext) { return }

        $targetPath = Normalize-PathForSave -root $BaseFolder -ver $verFolder -file $fileName

        # Throttle: if we've saved this path recently, skip
        if ($Script:LastSavedTimes.ContainsKey($targetPath)) {
            $last = $Script:LastSavedTimes[$targetPath]
            if (((Get-Date) - $last).TotalSeconds -lt $MinSaveThrottleSeconds) {
                Write-Verbose "Throttling save for $targetPath"
                return
            }
        }

        # Check version-aware overwrite
        $incomingVer = $null
        if ($Content -match "(?m)^[#;]\s*(ModuleVersion|ScriptVersion)\s*[:=]\s*([0-9a-zA-Z\.\-]+)") { 
            $incomingVer = $matches[2].Trim() 
        }

        $existingVer = Read-Top-Versions -FilePath $targetPath

        if ($existingVer -and $incomingVer) {
            try {
                if ([version]$existingVer -ge [version]$incomingVer) {
                    Write-Host "Skip: existing version $existingVer >= incoming $incomingVer -> $targetPath" -ForegroundColor Yellow
                    return
                }
            } catch {}
        }

        # Backup & save
        if (Test-Path $targetPath) { Backup-File -Path $targetPath | Out-Null }

        $dir = Split-Path $targetPath -Parent
        if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        Set-Content -Path $targetPath -Value $Content -Encoding UTF8 -Force
        $Script:LastSavedTimes[$targetPath] = Get-Date
        $Script:SaveCount++
        Update-TrayTooltip

        # Create shortcut if desired
        New-FileShortcut -Path $targetPath

        # Update history for dashboard
        $obj = [pscustomobject]@{
            Version = $verFolder
            File    = $fileName
            Path    = $targetPath
            Time    = (Get-Date).ToString('HH:mm:ss')
        }
        
        # Update on UI thread
        if ($Script:DashboardWindow) {
            $Script:DashboardWindow.Dispatcher.Invoke([action]{
                [void]$Script:SavedHistory.Insert(0, $obj)
                while ($Script:SavedHistory.Count -gt $MaxHistory) { 
                    $Script:SavedHistory.RemoveAt($Script:SavedHistory.Count - 1) 
                }
            })
        }

        # Feedback
        Play-SystemSound -Type 'Asterisk'
        Speak-Message -Text ("Saved: {0}" -f $fileName) -Async
        Write-Host "SAVED: $verFolder\$fileName" -ForegroundColor Green
    } catch {
        Write-Verbose "Error processing clipboard content: $_"
    }
}

# =====================================================================
# Clipboard Image Handler
# =====================================================================

function Save-ClipboardImage {
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
        $img = [System.Windows.Forms.Clipboard]::GetImage()
        if (-not $img) { return }
        
        $hash = Get-ImageHash -Img $img
        if ($hash -and $hash -eq $Script:LastImageHash) { 
            $img.Dispose()
            return 
        }
        $Script:LastImageHash = $hash

        $ver = $DefaultVersion
        $timestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
        $file = "Image_$timestamp.png"
        $dir = Join-Path $BaseFolder $ver
        $imgDir = Join-Path $dir 'Images'
        if (-not (Test-Path $imgDir)) { New-Item -ItemType Directory -Path $imgDir -Force | Out-Null }
        $path = Join-Path $imgDir $file
        $img.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
        $img.Dispose()
        
        $Script:SaveCount++
        Update-TrayTooltip
        New-FileShortcut -Path $path
        $obj = [pscustomobject]@{ 
            Version = $ver
            File = "Images\$file"
            Path = $path
            Time = (Get-Date).ToString("HH:mm:ss")
        }
        
        # Update on UI thread
        if ($Script:DashboardWindow) {
            $Script:DashboardWindow.Dispatcher.Invoke([action]{
                [void]$Script:SavedHistory.Insert(0,$obj)
            })
        }
        
        Speak-Message -Text "Image saved!" -Async
        Write-Host "SAVED IMAGE: $ver\Images\$file" -ForegroundColor Green
    } catch { 
        Write-Verbose "Image save skipped: $_"
    }
}

# =====================================================================
# Clipboard monitoring main loop
# =====================================================================

function Start-ClipboardMonitor {
    Write-Host ""
    Write-Host "KPopupClipSaver v$($Script:ScriptVersion)" -ForegroundColor Yellow
    Write-Host "Monitoring clipboard (base: $BaseFolder)" -ForegroundColor Gray
    Write-Host ""
    
    Initialize-TrayIcon
    Show-Dashboard

    $checkInterval = $DefaultCheckIntervalMs
    while ($true) {
        Start-Sleep -Milliseconds $checkInterval
        try {
            # Text clipboard read
            $txt = $null
            try { $txt = Get-Clipboard -Format Text -ErrorAction Stop } catch {}
            if ($txt) {
                $txtStr = $txt -as [string]
                $hash = Get-ClipboardHash -Text $txtStr
                if ($hash -and $hash -ne $Script:LastClipHash) {
                    $Script:LastClipHash = $hash
                    Save-ClipboardFile -Content $txtStr
                }
            }

            # Image handling
            Save-ClipboardImage
            
            # Allow GUI dispatchers to update
            if ($Script:DashboardWindow) { 
                try { 
                    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke(
                        [action]{}, 
                        [System.Windows.Threading.DispatcherPriority]::Background
                    ) 
                } catch {}
            }
        } catch { 
            Write-Verbose "Monitoring error: $_" 
        }
    }
}

# =====================================================================
# Stop routine
# =====================================================================

function Stop-ClipSaver {
    Write-Host "Stopping KPopupClipSaver..." -ForegroundColor Yellow
    if ($Script:TrayIcon) { 
        $Script:TrayIcon.Visible = $false
        $Script:TrayIcon.Dispose() 
    }
    if ($Script:DashboardWindow) { 
        try { $Script:DashboardWindow.Close() } catch {} 
    }
    exit 0
}

# =====================================================================
# Start up / Entry
# =====================================================================

try {
    # Ensure base folder exists
    if (-not (Test-Path $BaseFolder)) { 
        New-Item -ItemType Directory -Path $BaseFolder -Force | Out-Null 
    }
    
    # Ensure default version folder exists
    $vpath = Join-Path $BaseFolder $DefaultVersion
    if (-not (Test-Path $vpath)) { 
        New-Item -ItemType Directory -Path $vpath -Force | Out-Null 
    }

    # Warm up TTS voice list if available
    if ($Script:TTS) {
        try {
            $voices = $Script:TTS.GetInstalledVoices() | ForEach-Object { $_.VoiceInfo.Name }
            if ($voices) { 
                $Script:TTS.SelectVoice($voices[0])
                $Script:SelectedVoice = $voices[0] 
            }
        } catch {}
    }

    Start-ClipboardMonitor

} catch {
    Write-Error "Fatal: $_"
    Stop-ClipSaver
} finally {
    # never reached normally
    if ($Script:TrayIcon) { 
        $Script:TrayIcon.Visible = $false
        $Script:TrayIcon.Dispose() 
    }
}