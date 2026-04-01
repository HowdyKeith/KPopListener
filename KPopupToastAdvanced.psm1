# File: KPopupToastAdvanced.psm1
# Path: KPopupSuite/Modules/KPopupToastAdvanced.psm1
# Version: 2.0.0
Add-Type -AssemblyName System.Windows.Forms, System.Drawing -ErrorAction SilentlyContinue

if (-not $global:KPopupToastQueue) { $global:KPopupToastQueue = [System.Collections.Concurrent.ConcurrentQueue[object]]::new() }
if (-not $global:KPopupToastCallbacks) { $global:KPopupToastCallbacks = [System.Collections.Concurrent.ConcurrentDictionary[string, scriptblock]]::new() }

function Show-KPopupToastAdvanced {
    param(
        [string]$Title='KPopup',
        [string]$Message='Hello',
        [int]$Seconds=4,
        [hashtable]$Actions = $null,
        [switch]$Progress,
        [int]$ProgressPercent = 0,
        [string]$Id
    )
    if (-not $Id) { $Id = [guid]::NewGuid().ToString() }
    if ($Actions) { foreach ($k in $Actions.Keys) { $global:KPopupToastCallbacks["$Id|$k"] = $Actions[$k] } }
    $entry = [pscustomobject]@{ Title=$Title; Message=$Message; Seconds=$Seconds; Actions=$Actions; Progress=$Progress; Percent=$ProgressPercent; Id=$Id }
    $global:KPopupToastQueue.Enqueue($entry)
    # Start processor job if not running
    if (-not $global:KPopupToastProcessorRunning) {
        $global:KPopupToastProcessorRunning = $true
        Start-Job -Name KPopupToastProcessor -ScriptBlock {
            Add-Type -AssemblyName System.Windows.Forms, System.Drawing -ErrorAction SilentlyContinue
            while ($true) {
                $q = $using:global:KPopupToastQueue
                if ($q.TryDequeue([ref]$it)) {
                    try {
                        $f = New-Object System.Windows.Forms.Form
                        $f.FormBorderStyle = 'None'; $f.TopMost = $true; $f.StartPosition = 'Manual'
                        $screen = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
                        $f.Size = New-Object System.Drawing.Size(360,100)
                        $f.Location = New-Object System.Drawing.Point($screen.Width - 380, $screen.Height - 140)
                        $lblT = New-Object System.Windows.Forms.Label; $lblT.Text = $it.Title; $lblT.Font = New-Object System.Drawing.Font('Segoe UI',10,[System.Drawing.FontStyle]::Bold); $lblT.Location = New-Object System.Drawing.Point(12,6); $lblT.AutoSize=$true
                        $lblM = New-Object System.Windows.Forms.Label; $lblM.Text=$it.Message; $lblM.Location = New-Object System.Drawing.Point(12,34); $lblM.Size = New-Object System.Drawing.Size(336,34)
                        $f.Controls.AddRange(@($lblT,$lblM))
                        if ($it.Progress) {
                            $pb = New-Object System.Windows.Forms.ProgressBar; $pb.Size = New-Object System.Drawing.Size(336,14); $pb.Location = New-Object System.Drawing.Point(12,70); $pb.Value = [int]$it.Percent
                            $f.Controls.Add($pb)
                        }
                        if ($it.Actions) {
                            $x = 12
                            foreach ($a in $it.Actions.Keys) {
                                $btn = New-Object System.Windows.Forms.Button; $btn.Text=$a; $btn.Size=New-Object System.Drawing.Size(80,26); $btn.Location=New-Object System.Drawing.Point($x,70)
                                $btn.Add_Click({
                                    param($s,$e)
                                    $key = "$($it.Id)|$($s.Text)"
                                    if ($using:global:KPopupToastCallbacks.ContainsKey($key)) {
                                        $cb = $using:global:KPopupToastCallbacks[$key]
                                        & $cb
                                    }
                                })
                                $f.Controls.Add($btn); $x += 90
                            }
                        }
                        $f.Show()
                        Start-Sleep -Seconds $it.Seconds
                        $f.Close()
                    } catch {}
                } else { Start-Sleep -Milliseconds 250 }
            }
        } | Out-Null
    }
    return $Id
}

Export-ModuleMember -Function Show-KPopupToastAdvanced
