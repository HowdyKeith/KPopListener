# File: v7/KPopupClipSaver.ps1
# ScriptVersion: 7.81
# SINGLE FILE — NO DEPENDENCIES — WORKS IMMEDIATELY

Add-Type @'
using System;
using System.Runtime.InteropServices;
public class Win32 {
    [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
}
'@
[Win32]::ShowWindow([Win32]::GetConsoleWindow(), 0) | Out-Null

$BaseFolder = Split-Path -Parent $PSScriptRoot
$DefaultVersion = (Get-ChildItem $BaseFolder -Directory | ? Name -match '^v\d+' | Sort Name -Desc | Select -First 1).Name
if (!$DefaultVersion) { $DefaultVersion = 'v1' }

$Script:History = [System.Collections.ObjectModel.ObservableCollection[object]]::new()
$Script:PlaySounds = $true

# Tray icon
Add-Type -AssemblyName System.Windows.Forms, System.Drawing
$icon = New-Object System.Windows.Forms.NotifyIcon
$icon.Icon = [System.Drawing.SystemIcons]::Application
$icon.Visible = $true
$icon.Text = "KPopupClipSaver v7.81"

$menu = New-Object System.Windows.Forms.ContextMenuStrip
$menu.Items.Add("Dashboard") | % { $_.add_Click({ & "$PSScriptRoot\KPopupClipSaver.ps1" dashboard }) }
$menu.Items.Add("Exit")      | % { $_.add_Click({ $icon.Visible = $false; exit }) }
$icon.ContextMenuStrip = $menu
$icon.add_DoubleClick({ & "$PSScriptRoot\KPopupClipSaver.ps1" dashboard })

# Dashboard
Add-Type -AssemblyName PresentationFramework
function Show-Dashboard {
    if ($script:win -and $script:win.IsVisible) { $script:win.Activate(); return }
    $xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="KPopupClipSaver v7.81" Height="600" Width="900" Background="#121212" Foreground="White">
  <Grid Margin="10">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
    </Grid.RowDefinitions>
    <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
      <Button Content="Clear" Width="80" Margin="5"/>
      <Button Content="Sounds ON" Name="snd" Width="100" Margin="5"/>
    </StackPanel>
    <ListBox Grid.Row="1" Name="list" Background="#1E1E1E">
      <ListBox.ItemTemplate>
        <DataTemplate>
          <StackPanel Orientation="Horizontal" Margin="5">
            <TextBlock Text="{Binding Version}" Width="60" Foreground="#4CAF50" FontWeight="Bold"/>
            <TextBlock Text="{Binding File}" Margin="10,0"/>
            <TextBlock Text="{Binding Time}" Foreground="#888" Margin="20,0,0,0"/>
          </StackPanel>
        </DataTemplate>
      </ListBox.ItemTemplate>
    </ListBox>
  </Grid>
</Window>
'@
    $script:win = [Windows.Markup.XamlReader]::Load([System.Xml.XmlReader]::Create([System.IO.StringReader]$xaml))
    $script:win.FindName('list').ItemsSource = $Script:History
    $script:win.FindName('snd').Add_Click({
        $Script:PlaySounds = -not $Script:PlaySounds
        $this.Content = "Sounds " + $(if($Script:PlaySounds){"ON"}else{"OFF"})
    })
    $script:win.FindName('Clear').Add_Click({ $Script:History.Clear() })
    $script:win.Show()
}

# Save with auto-increment
function Save-Clip([string]$c) {
    if ($c -notmatch "(?m)^[#;]\s*(File|Path)\s*[:=]\s*(.+)$") { return }
    $path = $matches[2].Trim()
    $ver = $DefaultVersion
    $file = $path
    if ($path -match '^(v\d+)[/\\](.+)$') { $ver = $matches[1]; $file = $matches[2] }
    $full = Join-Path $BaseFolder "$ver\$file"
    $dir = Split-Path $full -Parent
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

    # Auto-increment version
    if ($c -match "(?m)^([#;]\s*(ModuleVersion|ScriptVersion)\s*[:=]\s*)([0-9]+\.[0-9]+)") {
        if (Test-Path $full) {
            $old = (Select-String -Path $full -Pattern $matches[1]).Line -replace '.*([0-9]+\.[0-9]+).*','$1'
            if ($old -and [version]$old -ge [version]$matches[3]) {
                $v = [version]$old
                $new = "{0}.{1}" -f $v.Major, ($v.Minor + 1)
                $c = $c -replace [regex]::Escape($matches[0]), "$($matches[1])$new"
                Write-Host "Bumped $old → $new" -ForegroundColor Cyan
            }
        }
    }

    Set-Content -Path $full -Value $c -Encoding UTF8
    $item = [pscustomobject]@{ Version=$ver; File=$file; Time=(Get-Date -f 'HH:mm:ss') }
    if ($script:win) { $script:win.Dispatcher.Invoke({ $Script:History.Insert(0,$item) }) }
    if ($Script:PlaySounds) { [System.Media.SystemSounds]::Asterisk.Play() }
    Write-Host "SAVED: $ver\$file" -ForegroundColor Green
}

# Image save
function Save-Image {
    $img = [System.Windows.Forms.Clipboard]::GetImage()
    if (!$img) { return }
    $ts = Get-Date -Format "yyyyMMdd_HHmmss"
    $p = Join-Path $BaseFolder "$DefaultVersion\Images\img_$ts.png"
    $d = Split-Path $p -Parent
    if (-not (Test-Path $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
    $img.Save($p)
    $img.Dispose()
    $item = [pscustomobject]@{ Version=$DefaultVersion; File="Images\img_$ts.png"; Time=(Get-Date -f 'HH:mm:ss') }
    if ($script:win) { $script:win.Dispatcher.Invoke({ $Script:History.Insert(0,$item) }) }
}

# Dashboard shortcut
if ($args[0] -eq "dashboard") { Show-Dashboard; return }

# Main loop
Show-Dashboard
$last = ''
while ($true) {
    Start-Sleep -Milliseconds 700
    $text = Get-Clipboard -Format Text -ErrorAction SilentlyContinue
    if ($text -and (Get-FileHash -InputStream ([System.IO.MemoryStream]::new([Text.Encoding]::UTF8.GetBytes($text))) -Algorithm SHA256).Hash -ne $last) {
        $last = (Get-FileHash -InputStream ([System.IO.MemoryStream]::new([Text.Encoding]::UTF8.GetBytes($text))) -Algorithm SHA256).Hash
        Save-Clip $text
    }
    Save-Image
}