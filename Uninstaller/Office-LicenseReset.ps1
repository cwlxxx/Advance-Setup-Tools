# ============================================================
#  Advance Setup Tools - Office License Reset / Remove Tool
#  Version 1.0  |  Full License Wipe (Safe for Cloud Re-Login)
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
Write-Host "   Microsoft Office License Reset Tool      " -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

# --- User Confirmation ---
Write-Host "[⚠️] This will remove ALL local Office license and activation data." -ForegroundColor Yellow
Write-Host "    Use only if Outlook or activation is broken." -ForegroundColor Yellow
Write-Host "    After reset, you must open Word or Excel and sign in again." -ForegroundColor Yellow
$confirm = Read-Host "`nType YES to continue"
if ($confirm -ne "YES") {
    Write-Host "`n[!] Cancelled by user. No changes made." -ForegroundColor Yellow
    pause
    exit
}

# --- Stop Office and ClickToRun services ---
Write-Host "`n[⏹️] Stopping Office processes and services..." -ForegroundColor Cyan
Get-Process WINWORD,EXCEL,OUTLOOK,POWERPNT,ONENOTE,MSACCESS,OfficeClickToRun -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Stop-Service -Name ClickToRunSvc -Force -ErrorAction SilentlyContinue
Stop-Service -Name OfficeSvc -Force -ErrorAction SilentlyContinue

Start-Sleep -Seconds 2

# --- Define target paths ---
$pathsToWipe = @(
    "C:\ProgramData\Microsoft\Office\Licenses",
    "C:\ProgramData\Microsoft\OfficeSoftwareProtectionPlatform",
    "$env:LOCALAPPDATA\Microsoft\Office\Licensing",
    "$env:LOCALAPPDATA\Microsoft\Office\Licenses",
    "$env:LOCALAPPDATA\Microsoft\Office\16.0\Licensing"
)

# --- Delete license folders ---
Write-Host "[🧹] Removing license token folders..." -ForegroundColor Cyan
foreach ($p in $pathsToWipe) {
    if (Test-Path $p) {
        try {
            Remove-Item $p -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host "  Removed: $p" -ForegroundColor Green
        } catch {
            Write-Host "  Failed:  $p" -ForegroundColor Red
        }
    }
}

# --- Registry cleanup ---
Write-Host "`n[🧾] Cleaning registry license data..." -ForegroundColor Cyan
$regKeys = @(
    "HKLM:\SOFTWARE\Microsoft\Office\Licenses",
    "HKLM:\SOFTWARE\Microsoft\Office\Registration",
    "HKLM:\SOFTWARE\Microsoft\Office\16.0\Registration",
    "HKCU:\Software\Microsoft\Office",
    "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
)
foreach ($key in $regKeys) {
    if (Test-Path $key) {
        try {
            Remove-Item $key -Recurse -Force -ErrorAction SilentlyContinue
            Write-Host "  Removed: $key" -ForegroundColor Green
        } catch {
            Write-Host "  Failed:  $key" -ForegroundColor Red
        }
    }
}

# --- Restart ClickToRun service ---
Write-Host "`n[♻️] Restarting ClickToRun service..." -ForegroundColor Cyan
try {
    Start-Service -Name ClickToRunSvc -ErrorAction SilentlyContinue
    Write-Host "  ClickToRun service restarted successfully." -ForegroundColor Green
} catch {
    Write-Host "  (ClickToRun service not found or disabled.)" -ForegroundColor Yellow
}

# --- Completion message ---
Write-Host "`n============================================" -ForegroundColor Cyan
Write-Host " Office License Reset Complete" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "All license files and activation tokens were removed."
Write-Host "Please open Word or Excel and sign in with your valid Microsoft account."
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

pause
exit
