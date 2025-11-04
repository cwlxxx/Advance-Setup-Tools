# ============================================================
# ⚙️ PowerShell 7 Script - Disable Hibernation and Fast Startup
# ============================================================

# ------------------------------------------------------------
# 🧩 Section : Script Metadata - Start
# ------------------------------------------------------------
$ScriptName    = "Disable Hibernation & Fast Startup"
$ScriptVersion = "1.0"
Write-Host "`n=== $ScriptName (v$ScriptVersion) ===`n"
# ------------------------------------------------------------
# 🧩 Section : Script Metadata - End
# ------------------------------------------------------------


# ------------------------------------------------------------
# ⚙️ Section : Ensure Administrator - Start
# ------------------------------------------------------------
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "⛔ This script must be run as Administrator." -ForegroundColor Red
    exit
}
# ------------------------------------------------------------
# ⚙️ Section : Ensure Administrator - End
# ------------------------------------------------------------


# ------------------------------------------------------------
# 🔋 Section : Disable Hibernation - Start
# ------------------------------------------------------------
Write-Host "Disabling hibernation..." -ForegroundColor Cyan
try {
    powercfg /hibernate off | Out-Null
    Write-Host "✔ Hibernation has been disabled."
} catch {
    Write-Host "❌ Failed to disable hibernation: $_" -ForegroundColor Red
}
# ------------------------------------------------------------
# 🔋 Section : Disable Hibernation - End
# ------------------------------------------------------------


# ------------------------------------------------------------
# 🧹 Section : Remove Hibernation File - Start
# ------------------------------------------------------------
$hiberFile = "C:\hiberfil.sys"
if (Test-Path $hiberFile) {
    Write-Host "Removing hibernation file..." -ForegroundColor Cyan
    try {
        Remove-Item -Path $hiberFile -Force
        Write-Host "✔ Hibernation file removed."
    } catch {
        Write-Host "❌ Could not remove hibernation file: $_" -ForegroundColor Red
    }
} else {
    Write-Host "No hibernation file found."
}
# ------------------------------------------------------------
# 🧹 Section : Remove Hibernation File - End
# ------------------------------------------------------------


# ------------------------------------------------------------
# ⚡ Section : Disable Fast Startup - Start
# ------------------------------------------------------------
Write-Host "Disabling Windows Fast Startup..." -ForegroundColor Cyan
try {
    # Registry path for Fast Startup
    $regPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power"
    Set-ItemProperty -Path $regPath -Name HiberbootEnabled -Value 0
    Write-Host "✔ Fast Startup has been disabled."
} catch {
    Write-Host "❌ Failed to disable Fast Startup: $_" -ForegroundColor Red
}
# ------------------------------------------------------------
# ⚡ Section : Disable Fast Startup - End
# ------------------------------------------------------------


# ------------------------------------------------------------
# ✅ Section : Summary - Start
# ------------------------------------------------------------
Write-Host "`nAll tasks completed successfully." -ForegroundColor Green
Write-Host "You may need to restart your system for changes to take full effect."
# ------------------------------------------------------------
# ✅ Section : Summary - End
# ------------------------------------------------------------
