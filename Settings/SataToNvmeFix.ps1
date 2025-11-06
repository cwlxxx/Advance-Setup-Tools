# ============================================================
#  Advance Setup Tools - SATA to NVMe Clone Fix
#  Version 1.0  |  Enables NVMe Driver Before or After Clone
#  Compatible: Windows 10 / 11  |  PowerShell 7+
# ============================================================

# --- Ensure running as Administrator ---
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "[!] Please run this script as Administrator." -ForegroundColor Yellow
    pause
    exit
}

Clear-Host
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "   SATA to NVMe Boot Fix Utility" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "[🔧] Enabling NVMe boot driver..." -ForegroundColor Yellow

# --- Enable NVMe driver (stornvme) ---
try {
    reg add "HKLM\SYSTEM\CurrentControlSet\Services\stornvme" /v Start /t REG_DWORD /d 0 /f | Out-Null
    Write-Host "[✅] NVMe driver successfully enabled." -ForegroundColor Green
}
catch {
    Write-Host "[❌] Failed to modify registry: $($_.Exception.Message)" -ForegroundColor Red
    pause
    exit
}

# --- Confirmation popup ---
Add-Type -AssemblyName PresentationFramework
[System.Windows.MessageBox]::Show(
    "✅ SATA to NVMe Clone Fix applied successfully.`n`nYou can now clone and boot from NVMe safely.",
    "SATA → NVMe Fix Complete",
    [System.Windows.MessageBoxButton]::OK,
    [System.Windows.MessageBoxImage]::Information
) | Out-Null

Write-Host "`nAll done. You may close this window." -ForegroundColor Cyan
pause
exit
