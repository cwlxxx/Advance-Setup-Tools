Write-Host "Installing IrfanView 64..." -ForegroundColor Cyan
winget settings --enable InstallerHashOverride
# Run the Winget installation (silent)
Start-Process "winget.exe" -ArgumentList "install --id=IrfanSkiljan.IrfanView --source winget --exact --accept-package-agreements --accept-source-agreements --ignore-security-hash" -NoNewWindow -Wait

Write-Host "⏳ Waiting for IrfanView installation to complete..." -ForegroundColor Cyan

# Shortcut filenames that IrfanView creates
$shortcutNames = @("IrfanView.lnk", "IrfanView 64.lnk")
$desktopPaths = @(
    "$env:PUBLIC\Desktop",
    [Environment]::GetFolderPath('Desktop')
)

# Wait up to 45 seconds (longer than before)
$foundShortcut = $false
for ($i = 0; $i -lt 45; $i++) {
    foreach ($desktop in $desktopPaths) {
        foreach ($name in $shortcutNames) {
            $shortcut = Join-Path $desktop $name
            if (Test-Path $shortcut) {
                $foundShortcut = $true
                Write-Host "🧩 Detected shortcut: $shortcut" -ForegroundColor DarkGray
            }
        }
    }

    if ($foundShortcut) { break }
    Start-Sleep -Seconds 1
}

if ($foundShortcut) {
    Write-Host "🧩 IrfanView shortcut detected — proceeding to remove..." -ForegroundColor Yellow
} else {
    Write-Host "⚠️ No IrfanView shortcut detected after waiting — continuing anyway." -ForegroundColor DarkYellow
}

# 🗑️ Attempt to remove all shortcuts (with retry)
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

Write-Host "✅ IrfanView desktop shortcut cleanup completed." -ForegroundColor Cyan
