# Module: KPopDashboard.psm1

# Live control panel for KPopListener settings

function Show-KPopDashboard {
param(
[switch]$AutoRefresh
)

```
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = "KPopListener Dashboard"
$form.Size = New-Object System.Drawing.Size(400,200)
$form.StartPosition = "CenterScreen"

# UseRawStream checkbox
$chkRawStream = New-Object System.Windows.Forms.CheckBox
$chkRawStream.Text = "Use Raw Stream"
$chkRawStream.Location = New-Object System.Drawing.Point(20,20)
$chkRawStream.Checked = $global:UseRawStream
$chkRawStream.Add_CheckedChanged({
    $global:UseRawStream = $chkRawStream.Checked
    Write-Log "Dashboard: UseRawStream set to $($global:UseRawStream)"
    Update-ListenerStatus
})
$form.Controls.Add($chkRawStream)

# RaiseEventMode checkbox
$chkRaiseEvent = New-Object System.Windows.Forms.CheckBox
$chkRaiseEvent.Text = "Raise Event Mode"
$chkRaiseEvent.Location = New-Object System.Drawing.Point(20,50)
$chkRaiseEvent.Checked = $global:RaiseEventMode
$chkRaiseEvent.Add_CheckedChanged({
    $global:RaiseEventMode = $chkRaiseEvent.Checked
    Write-Log "Dashboard: RaiseEventMode set to $($global:RaiseEventMode)"
    Update-ListenerStatus
})
$form.Controls.Add($chkRaiseEvent)

# Refresh button for status
$btnRefresh = New-Object System.Windows.Forms.Button
$btnRefresh.Text = "Refresh Status"
$btnRefresh.Location = New-Object System.Drawing.Point(20,80)
$btnRefresh.Add_Click({
    Update-ListenerStatus
    Write-Log "Dashboard: Status refreshed"
})
$form.Controls.Add($btnRefresh)

# Timer for auto-refresh
if ($AutoRefresh) {
    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 2000
    $timer.Add_Tick({
        Update-ListenerStatus
    })
    $timer.Start()
}

$form.Topmost = $true
$form.Add_Shown({$form.Activate()})
[void]$form.ShowDialog()
```

}

function Start-KPopDashboard {
param([switch]$AutoRefresh)
Show-KPopDashboard -AutoRefresh:$AutoRefresh
}
