# File: KPopupClipSaverAdvanced.psm1
# Path: KPopupSuite/Modules/KPopupClipSaverAdvanced.psm1
# Version: 2.1.0
function Get-StringHashSHA512 { param([string]$Text) $b=[System.Text.Encoding]::UTF8.GetBytes($Text); (Get-FileHash -InputStream ([IO.MemoryStream]::new($b)) -Algorithm SHA512).Hash }

function Split-MultiFileClipboard {
    param([string]$Content)
    $parts = @()
    $lines = $Content -split "`r?`n"
    $current = @()
    $currentHeader = $null
    foreach ($l in $lines) {
        if ($l -match '^\s*#\s*(File|Path)\s*:\s*(.+)$') {
            if ($currentHeader) { $parts += [pscustomobject]@{ Header=$currentHeader; Body=($current -join "`n") } }
            $currentHeader = $matches[2].Trim()
            $current = @()
        } else { $current += $l }
    }
    if ($currentHeader) { $parts += [pscustomobject]@{ Header=$currentHeader; Body=($current -join "`n") } }
    if ($parts.Count -gt 0) { return $parts }
    return ,([pscustomobject]@{ Header = $null; Body = $Content })
}

function Save-ClipContentAdvanced {
    param([string]$BaseFolder,[hashtable]$Payload)

    $content = $Payload.content
    $createShortcut = [bool]($Payload.createShortcut -ne $false)

    $parts = Split-MultiFileClipboard -Content $content
    foreach ($p in $parts) {
        $dest = $p.Header
        $body = $p.Body
        if (-not $dest) {
            $cls = 'text'
            if ($body -match '(?m)^\s*#\s*(File|Path)\s*:') { $cls = 'script' }
            $dest = "v1\snippets\auto_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
        }
        $ver='v1'; $file=$dest
        if ($dest -match '^(v[0-9\.]+)[\\/](.+)$') { $ver=$matches[1]; $file=$matches[2] }

        $root = Join-Path $BaseFolder $ver
        $modulesDir = Join-Path $root 'Modules'
        $backupDir = Join-Path $root 'Backup'
        @($root,$modulesDir,$backupDir) | ForEach-Object { if (-not (Test-Path $_)) { New-Item -ItemType Directory -Path $_ -Force | Out-Null } }

        $target = if ([IO.Path]::GetExtension($file).TrimStart('.') -eq 'psm1') { Join-Path $modulesDir $file } else { Join-Path $root $file }
        $incomingHash = Get-StringHashSHA512 -Text $body
        if (Test-Path $target) {
            $existing = Get-Content -Path $target -Raw -ErrorAction SilentlyContinue
            $existingHash = Get-StringHashSHA512 -Text $existing
            if ($existingHash -eq $incomingHash) { Write-Host "No change: $target"; continue }
            $bakName = "{0}-BAK_{1}{2}" -f ([IO.Path]::GetFileNameWithoutExtension($target)), (Get-Date -f 'yyyyMMdd_HHmmss'), [IO.Path]::GetExtension($target)
            Copy-Item -Path $target -Destination (Join-Path $backupDir $bakName) -Force
        }
        Set-Content -Path $target -Value $body -Encoding UTF8 -Force
        Write-Host "Saved: $target"
        if ($createShortcut) {
            try {
                $wsh = New-Object -ComObject WScript.Shell
                $lnk = "$target.lnk"
                $s = $wsh.CreateShortcut($lnk)
                $s.TargetPath = $target
                $s.WorkingDirectory = Split-Path $target -Parent
                $s.Save()
                [Runtime.InteropServices.Marshal]::ReleaseComObject($wsh) | Out-Null
            } catch { Write-Warning "Shortcut failed: $_" }
        }
    }
    return $true
}

function Start-KPopupClipSaverAdvanced {
    param(
        [string]$BaseFolder = (Join-Path (Split-Path -Parent $PSScriptRoot) 'v1'),
        [int]$IntervalMs = 500,
        [switch]$RunInBackground
    )
    if ($RunInBackground) {
        Start-Job -Name KPopupClipSaverAdvanced -ScriptBlock {
            param($base,$interval)
            Import-Module (Join-Path (Split-Path -Parent $PSScriptRoot) 'KPopupClipSaverAdvanced.psm1') -ErrorAction SilentlyContinue
            while ($true) {
                $txt = Get-Clipboard -ErrorAction SilentlyContinue
                if ($txt) {
                    $hash = (Get-FileHash -InputStream ([IO.MemoryStream]::new([Text.Encoding]::UTF8.GetBytes($txt))) -Algorithm SHA256).Hash
                    if ($env:KPopup_LastClip -ne $hash) {
                        $env:KPopup_LastClip = $hash
                        Save-ClipContentAdvanced -BaseFolder $base -Payload @{ content = $txt; createShortcut = $true }
                    }
                }
                Start-Sleep -Milliseconds $interval
            }
        } -ArgumentList $BaseFolder,$IntervalMs | Out-Null
        Write-Host "ClipSaver started as job."
        return
    }
    while ($true) {
        $txt = Get-Clipboard -ErrorAction SilentlyContinue
        if ($txt) {
            $hash = (Get-FileHash -InputStream ([IO.MemoryStream]::new([Text.Encoding]::UTF8.GetBytes($txt))) -Algorithm SHA256).Hash
            if ($script:LastClipHash -ne $hash) {
                $script:LastClipHash = $hash
                Save-ClipContentAdvanced -BaseFolder $BaseFolder -Payload @{ content = $txt; createShortcut = $true }
            }
        }
        Start-Sleep -Milliseconds $IntervalMs
    }
}

Export-ModuleMember -Function Save-ClipContentAdvanced, Start-KPopupClipSaverAdvanced, Split-MultiFileClipboard
