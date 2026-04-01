# File: KPopupTrayIcon.psm1
# Path: KPopupSuite/Modules/KPopupTrayIcon.psm1
Add-Type -AssemblyName System.Windows.Forms, System.Drawing -ErrorAction SilentlyContinue

function Start-KPopupTrayIcon {
    param([string]$Tooltip='KPopup',[switch]$EnableMenu)

    $ni = New-Object System.Windows.Forms.NotifyIcon
    $bmp = [System.Drawing.Bitmap]::new(32,32)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.Clear([System.Drawing.Color]::Transparent)
    $g.FillEllipse([System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(180,30,140,60)),4,4,24,24)
    $g.Dispose()
    $ni.Icon = [System.Drawing.Icon]::FromHandle($bmp.GetHicon())
    $ni.Visible = $true
    $ni.Text = $Tooltip
    if ($EnableMenu) {
        $menu = New-Object System.Windows.Forms.ContextMenuStrip
        $m1 = $menu.Items.Add('Open Dashboard'); $m1.add_Click({ if (Get-Command Show-KPopupDashboard -ErrorAction SilentlyContinue) { Show-KPopupDashboard } })
        $m2 = $menu.Items.Add('Open v1 Folder'); $m2.add_Click({ Start-Process explorer.exe -ArgumentList (Join-Path (Split-Path -Parent $PSScriptRoot) 'v1') })
        $m3 = $menu.Items.Add('Exit'); $m3.add_Click({ $ni.Visible=$false; exit })
        $ni.ContextMenuStrip = $menu
    }
    return $ni
}

Export-ModuleMember -Function Start-KPopupTrayIcon
