# File: KPopupClipSaver.ps1
# Path: KPopupSuite/KPopupClipSaver/KPopupClipSaver.ps1
# Standalone version with tray + dashboard; can also be used as module via KPopupClipSaverAdvanced.psm1

Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) '..\KPopupSuite\Modules\KPopupClipSaverAdvanced.psm1') -ErrorAction SilentlyContinue
Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) '..\KPopupSuite\Modules\KPopupTrayIcon.psm1') -ErrorAction SilentlyContinue
Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) '..\KPopupSuite\Modules\KPopupToastAdvanced.psm1') -ErrorAction SilentlyContinue

$root = (Split-Path -Parent $PSScriptRoot)
Start-KPopupClipSaverAdvanced -BaseFolder (Join-Path $root 'v1') -IntervalMs 800 -RunInBackground:$true
$tray = Start-KPopupTrayIcon -Tooltip 'KPopupClipSaver (standalone)' -EnableMenu
Write-Host "KPopupClipSaver running (standalone). Use tray menu to exit."
while ($true) { Start-Sleep -Seconds 60 }
