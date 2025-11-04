# =====================================================
# Set-ClockAndRegion.ps1
# Purpose: Configure system time zone and date format
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

# ----- 2. Set Date Format -----
try {
    $culture = Get-Culture
    $regPath = "HKCU:\Control Panel\International"

    # Set short date (for system clock and file explorer)
    Set-ItemProperty -Path $regPath -Name "sShortDate" -Value "dd-MM-yyyy"

    # Optional: long date (for full written dates)
    Set-ItemProperty -Path $regPath -Name "sLongDate" -Value "dddd, dd MMMM yyyy"

    # Optional: short time / long time formats (24-hour style)
    Set-ItemProperty -Path $regPath -Name "sShortTime" -Value "HH:mm"
    Set-ItemProperty -Path $regPath -Name "sTimeFormat" -Value "HH:mm:ss"

    Write-Host "Date and time formats set successfully (dd-MM-yyyy, 24-hour)" -ForegroundColor Green
} catch {
    Write-Host "Failed to set regional formats. Error: $_" -ForegroundColor Red
}

# ----- 3. Confirm Result -----
Write-Host ""
Write-Host "Current time zone:" -ForegroundColor Yellow
Get-TimeZone | Format-Table -AutoSize

Write-Host ""
Write-Host "Current date/time preview:" -ForegroundColor Yellow
Get-Date -Format "dd-MM-yyyy HH:mm:ss"

Write-Host ""
Write-Host "=== Configuration Complete ===" -ForegroundColor Cyan
