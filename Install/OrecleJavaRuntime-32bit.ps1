# ============================================================
# ☕ Oracle Java JRE 32-bit Downloader + Silent Installer
# ============================================================

Add-Type -AssemblyName PresentationFramework

# --- Configuration ---
$Java32Url   = "https://javadl.oracle.com/webapps/download/AutoDL?BundleId=252321_68ce765258164726922591683c51982c"
$DownloadDir = Join-Path $env:TEMP "installer"
$Installer   = Join-Path $DownloadDir "JavaRuntime-Windows-x86.exe"

# --- Ensure download directory exists ---
if (-not (Test-Path $DownloadDir)) {
    New-Item -ItemType Directory -Path $DownloadDir | Out-Null
}

# --- Check URL validity ---
try {
    Write-Host "Checking download link..." -ForegroundColor Cyan
    $response = Invoke-WebRequest -Uri $Java32Url -Method Head -UseBasicParsing -TimeoutSec 10
    if ($response.StatusCode -ne 200) {
        [System.Windows.MessageBox]::Show("The Java 32-bit download link has expired or is invalid.", "Download Failed", "OK", "Error") | Out-Null
        throw "Invalid link"
    }
}
catch {
    [System.Windows.MessageBox]::Show("The Java 32-bit download link has expired or cannot be reached.", "Download Error", "OK", "Error") | Out-Null
    exit
}

# --- Download via BITS ---
try {
    Write-Host "Downloading Java Runtime (32-bit)..." -ForegroundColor Yellow
    Start-BitsTransfer -Source $Java32Url -Destination $Installer -DisplayName "Downloading Java 32-bit Runtime"
    Write-Host "✅ Download completed: $Installer" -ForegroundColor Green
}
catch {
    [System.Windows.MessageBox]::Show("Failed to download Java 32-bit installer via BITS.", "Download Failed", "OK", "Error") | Out-Null
    exit
}

# --- Run silent installer (shows progress UI) ---
try {
    Write-Host "Installing Java Runtime (32-bit)..." -ForegroundColor Cyan

    # The /s argument runs the installer silently (no dialogs),
    # but the main progress window is still displayed.
    $arguments = "/s"
    $process = Start-Process -FilePath $Installer -ArgumentList $arguments -PassThru -Wait

    if ($process.ExitCode -eq 0) {
        Write-Host "✅ Java Runtime installed successfully." -ForegroundColor Green
    } else {
        Write-Warning "⚠ Installer exited with code $($process.ExitCode)."
    }
}
catch {
    [System.Windows.MessageBox]::Show("An error occurred during installation.", "Install Failed", "OK", "Error") | Out-Null
    exit
}

# --- Cleanup downloaded file ---
try {
    Remove-Item -Path $Installer -Force -ErrorAction SilentlyContinue
    Write-Host "🧹 Temporary installer removed." -ForegroundColor DarkGray
}
catch {
    Write-Warning "Could not delete installer file: $Installer"
}
