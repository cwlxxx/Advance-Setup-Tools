# =====================================================
# Set-ClockAndRegion.ps1
# Purpose: Configure system time zone and date/time format
# Compatible: Windows 10 / Windows 11
# Requires: Administrator privileges
# =====================================================

Write-Host "=== Setting Timezone and Regional Format (MY) ===" -ForegroundColor Cyan

# ----- 1. Set Time Zone -----
try {
    $tz = "Singapore Standard Time"  # UTC+08:00 Kuala Lumpur, Singapore
    Set-TimeZone -Id $tz
    Write-Host "Time zone set to: $tz" -ForegroundColor Green
} catch {
    Write-Host "Failed to set time zone. Error: $_" -ForegroundColor Red
}

# ----- 2. Set Date/Time Format -----
try {
    $regPath = "HKCU:\Control Panel\International"

    # Date formats
    Set-ItemProperty -Path $regPath -Name "sShortDate" -Value "dd-MM-yyyy"
    Set-ItemProperty -Path $regPath -Name "sLongDate" -Value "dddd, dd MMMM yyyy"

    # Time formats (12-hour with AM/PM)
    # h:mm tt → 1:45 PM   |   hh:mm:ss tt → 01:45:30 PM
    Set-ItemProperty -Path $regPath -Name "sShortTime" -Value "h:mm tt"
    Set-ItemProperty -Path $regPath -Name "sTimeFormat" -Value "hh:mm:ss tt"

    Write-Host "Date and time formats set successfully (dd-MM-yyyy, 12-hour AM/PM)" -ForegroundColor Green
} catch {
    Write-Host "Failed to set regional formats. Error: $_" -ForegroundColor Red
}

# ----- 3. Apply Changes Instantly -----
try {
    Write-Host "`nApplying settings immediately..." -ForegroundColor Cyan

    # Refresh registry values in current session
    rundll32.exe user32.dll,UpdatePerUserSystemParameters

    # Restart Explorer for visible clock update (safe)
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

# ----- 4. Confirm Result -----
Write-Host ""
Write-Host "Current time zone:" -ForegroundColor Yellow
Get-TimeZone | Format-Table -AutoSize

Write-Host ""
Write-Host "Current date/time preview:" -ForegroundColor Yellow
Get-Date -Format "dd-MM-yyyy hh:mm:ss tt"

Write-Host ""
Write-Host "=== Configuration Complete ===" -ForegroundColor Cyan
