Write-Host "=== Setting Timezone and Regional Format (MY) ===" -ForegroundColor Cyan

try {
    $tz = "Singapore Standard Time"
    Set-TimeZone -Id $tz
    Write-Host "Time zone set to: $tz" -ForegroundColor Green
} catch {
    Write-Host "Failed to set time zone. Error: $_" -ForegroundColor Red
}

try {
    $regPath = "HKCU:\Control Panel\International"

    Set-ItemProperty -Path $regPath -Name "sShortDate" -Value "dd-MM-yyyy"
    Set-ItemProperty -Path $regPath -Name "sLongDate" -Value "dddd, dd MMMM yyyy"

    Set-ItemProperty -Path $regPath -Name "sShortTime" -Value "h:mm tt"
    Set-ItemProperty -Path $regPath -Name "sTimeFormat" -Value "hh:mm:ss tt"

    Write-Host "Date and time formats set successfully (dd-MM-yyyy, 12-hour AM/PM)" -ForegroundColor Green
} catch {
    Write-Host "Failed to set regional formats. Error: $_" -ForegroundColor Red
}

try {
    Write-Host "`nApplying settings immediately..." -ForegroundColor Cyan

    rundll32.exe user32.dll,UpdatePerUserSystemParameters

    $explorers = Get-Process explorer -ErrorAction SilentlyContinue
    if ($explorers) {
        Write-Host "Restarting Explorer..." -ForegroundColor Yellow
        Stop-Process -Name explorer -Force
        Start-Sleep -Seconds 2
        Start-Process explorer.exe
        Write-Host "Explorer restarted successfully." -ForegroundColor Green
    } else {
        Write-Host "Explorer not running — skipping restart." -ForegroundColor Yellow
    }
} catch {
    Write-Host "Failed to refresh Explorer. Error: $_" -ForegroundColor Red
}

Write-Host ""
Write-Host "Current time zone:" -ForegroundColor Yellow
Get-TimeZone | Format-Table -AutoSize

Write-Host ""
Write-Host "Current date/time preview:" -ForegroundColor Yellow
Get-Date -Format "dd-MM-yyyy hh:mm:ss tt"

Write-Host ""
Write-Host "=== Configuration Complete ===" -ForegroundColor Cyan
