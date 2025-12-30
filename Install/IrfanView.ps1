Start-Process "winget.exe" -ArgumentList "install --id=IrfanSkiljan.IrfanView --source winget --exact --accept-package-agreements --accept-source-agreements" -NoNewWindow -Wait

Write-Host "Waiting for IrfanView installation to complete..." -ForegroundColor Cyan

$shortcutNames = @("IrfanView.lnk", "IrfanView 64.lnk")
$desktopPaths = @(
    "$env:PUBLIC\Desktop",
    [Environment]::GetFolderPath('Desktop')
)

$foundShortcut = $false
for ($i = 0; $i -lt 30; $i++) {
    foreach ($desktop in $desktopPaths) {
        foreach ($name in $shortcutNames) {
            $shortcut = Join-Path $desktop $name
            if (Test-Path $shortcut) {
                $foundShortcut = $true
                Write-Host "Detected shortcut: $shortcut" -ForegroundColor DarkGray
            }
        }
    }

    if ($foundShortcut) { break }
    Start-Sleep -Seconds 1
}

if ($foundShortcut) {
    Write-Host "IrfanView shortcut detected — proceeding to remove..." -ForegroundColor Yellow
} else {
    Write-Host "No IrfanView shortcut detected after waiting — continuing anyway." -ForegroundColor DarkYellow
}


foreach ($desktop in $desktopPaths) {
    foreach ($name in $shortcutNames) {
        $shortcut = Join-Path $desktop $name
        if (Test-Path $shortcut) {
            for ($retry = 0; $retry -lt 3; $retry++) {
                try {
                    Remove-Item $shortcut -Force -ErrorAction Stop
                    Write-Host "Removed desktop shortcut: $shortcut" -ForegroundColor Green
                    break
                } catch {
                    Write-Host "Failed to remove shortcut (attempt $($retry + 1)): $($_.Exception.Message)" -ForegroundColor Red
                    Start-Sleep -Seconds 1
                }
            }
        }
    }
}

Write-Host "IrfanView desktop shortcut cleanup completed." -ForegroundColor Cyan
