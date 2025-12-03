# ===========================================
# 🎥 Zoom Installer & Shortcut Cleanup (Fixed Edition)
# ===========================================

Write-Host "Installing Zoom..." -ForegroundColor Cyan
winget settings --enable InstallerHashOverride
# Run the Winget installation (silent)
Start-Process "winget.exe" -ArgumentList "install --id Zoom.Zoom --source winget --exact --accept-package-agreements --accept-source-agreements --ignore-security-hash" -NoNewWindow -Wait

Write-Host "⏳ Waiting for Zoom installation to complete..." -ForegroundColor Cyan

# Shortcut filenames created by modern Zoom versions
$shortcutNames = @("Zoom.lnk", "Zoom Workplace.lnk", "Zoom Meetings.lnk")
$desktopPaths = @(
    "$env:PUBLIC\Desktop",
    [Environment]::GetFolderPath('Desktop')
)

# Wait loop (max 45 seconds total)
$foundShortcut = $false
for ($i = 0; $i -lt 45; $i++) {
    foreach ($desktop in $desktopPaths) {
        foreach ($name in $shortcutNames) {
            $shortcut = Join-Path $desktop $name
            if (Test-Path $shortcut) {
                Write-Host "🧩 Detected shortcut: $shortcut" -ForegroundColor DarkGray
                $foundShortcut = $true
            }
        }
    }

    if ($foundShortcut) { break }
    Start-Sleep -Seconds 1
}

if ($foundShortcut) {
    Write-Host "🧩 Zoom shortcut detected — proceeding to remove..." -ForegroundColor Yellow
} else {
    Write-Host "⚠️ No Zoom shortcut detected after waiting — continuing anyway." -ForegroundColor DarkYellow
}

# 🗑️ Attempt to remove all possible Zoom shortcuts (with retries)
foreach ($desktop in $desktopPaths) {
    foreach ($name in $shortcutNames) {
        $shortcut = Join-Path $desktop $name
        if (Test-Path $shortcut) {
            for ($retry = 0; $retry -lt 3; $retry++) {
                try {
                    Remove-Item $shortcut -Force -ErrorAction Stop
                    Write-Host "🗑️ Removed desktop shortcut: $shortcut" -ForegroundColor Green
                    break
                } catch {
                    Write-Host "⚠️ Failed to remove shortcut (attempt $($retry + 1)): $($_.Exception.Message)" -ForegroundColor Red
                    Start-Sleep -Seconds 1
                }
            }
        }
    }
}

Write-Host "✅ Zoom desktop shortcut cleanup completed." -ForegroundColor Cyan
