# ============================================================
#  Advance Setup Tools - Office License Reset / Remove Tool
#  Version 1.1  |  Smart Detection + Safe Confirmation
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

# --- Define target paths and registry keys ---
$pathsToWipe = @(
    "C:\ProgramData\Microsoft\Office\Licenses",
    "C:\ProgramData\Microsoft\OfficeSoftwareProtectionPlatform",
    "$env:LOCALAPPDATA\Microsoft\Office\Licensing",
    "$env:LOCALAPPDATA\Microsoft\Office\Licenses",
    "$env:LOCALAPPDATA\Microsoft\Office\16.0\Licensing"
)

$regKeys = @(
    "HKLM:\SOFTWARE\Microsoft\Office\Licenses",
    "HKLM:\SOFTWARE\Microsoft\Office\Registration",
    "HKLM:\SOFTWARE\Microsoft\Office\16.0\Registration",
    "HKCU:\Software\Microsoft\Office",
    "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
)

# --- Step 1: Scan for license files and registry keys ---
Write-Host "[🔍] Scanning system for Office license and activation data..." -ForegroundColor Yellow
$foundItems = [System.Collections.Generic.List[string]]::new()

foreach ($p in $pathsToWipe) {
    if (Test-Path $p) {
        $count = (Get-ChildItem -Path $p -Recurse -ErrorAction SilentlyContinue | Measure-Object).Count
        if ($count -gt 0) {
            $foundItems.Add("Files: $count in $p")
        }
    }
}

foreach ($key in $regKeys) {
    if (Test-Path $key) {
        $foundItems.Add("Registry: $key")
    }
}

if ($foundItems.Count -eq 0) {
    Write-Host "`n[ℹ️] No Office license files or registry entries detected." -ForegroundColor Yellow
    pause
    exit
}

Write-Host "`n[✅] Found the following license data:" -ForegroundColor Cyan
$foundItems | ForEach-Object { Write-Host "  $_" -ForegroundColor Green }

# --- Step 2: Ask for confirmation ---
Write-Host ""
Write-Host "[⚠️] WARNING: This will permanently delete all detected license data." -ForegroundColor Yellow
Write-Host "    After reset, you must re-sign in with your Microsoft account to reactivate Office." -ForegroundColor Yellow
$confirm = Read-Host "`nType YES to continue"
if ($confirm -ne "YES") {
    Write-Host "`n[!] Cancelled by user. No changes made." -ForegroundColor Yellow
    pause
    exit
}

# --- Step 3: Stop Office and ClickToRun services ---
Write-Host "`n[⏹️] Stopping Office processes and services..." -ForegroundColor Cyan
Get-Process WINWORD,EXCEL,OUTLOOK,POWERPNT,ONENOTE,MSACCESS,OfficeClickToRun -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Stop-Service -Name ClickToRunSvc -Force -ErrorAction SilentlyContinue
Stop-Service -Name OfficeSvc -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2

# --- Step 4: Delete license folders ---
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

# --- Step 5: Registry cleanup ---
Write-Host "`n[🧾] Cleaning registry license data..." -ForegroundColor Cyan
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

# --- Step 6: Restart ClickToRun service ---
Write-Host "`n[♻️] Restarting ClickToRun service..." -ForegroundColor Cyan
try {
    Start-Service -Name ClickToRunSvc -ErrorAction SilentlyContinue
    Write-Host "  ClickToRun service restarted successfully." -ForegroundColor Green
} catch {
    Write-Host "  (ClickToRun service not found or disabled.)" -ForegroundColor Yellow
}

# --- Step 7: Summary ---
Write-Host "`n============================================" -ForegroundColor Cyan
Write-Host " Office License Reset Complete" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "All detected license files and registry entries were removed."
Write-Host "Please open Word or Excel and sign in again with your valid Microsoft account."
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""

pause
exit
