# ================================
# Install Google Chrome Script
# Winget first → fallback to MSI download
# Auto-delete downloaded installer
# ================================

Write-Host "Installing Google Chrome using winget..." -ForegroundColor Cyan

# Try Winget installation
$wingetResult = winget install --id Google.Chrome --source winget --exact --silent --accept-source-agreements --accept-package-agreements -h 2>&1

if ($LASTEXITCODE -eq 0) {
    Write-Host "Google Chrome installed successfully using winget!" -ForegroundColor Green
    exit
}
else {
    Write-Host "Winget install failed! Fallback to direct download..." -ForegroundColor Yellow
}

# ================================
# Step 2: Download Chrome Installer
# ================================

$chromeUrl = "https://dl.google.com/dl/chrome/install/googlechromestandaloneenterprise64.msi"
$downloadPath = "$env:TEMP\ChromeInstall.msi"

Write-Host "Downloading Chrome installer..." -ForegroundColor Cyan

try {
    Invoke-WebRequest -Uri $chromeUrl -OutFile $downloadPath
    Write-Host "Download completed: $downloadPath" -ForegroundColor Green
}
catch {
    Write-Host "Failed to download Chrome installer!" -ForegroundColor Red
    exit 1
}

# ================================
# Step 3: Install Chrome
# ================================

Write-Host "Installing Chrome using MSI installer..." -ForegroundColor Cyan

$installResult = Start-Process "msiexec.exe" -ArgumentList "/i `"$downloadPath`" /qn /norestart" -Wait -PassThru

# ================================
# Step 4: Cleanup Installer
# ================================

if (Test-Path $downloadPath) {
    Remove-Item $downloadPath -Force
    Write-Host "Cleaned up installer file." -ForegroundColor DarkGray
}

# ================================

if ($installResult.ExitCode -eq 0) {
    Write-Host "Google Chrome installed successfully!" -ForegroundColor Green
}
else {
    Write-Host "Chrome installation failed! Exit code: $($installResult.ExitCode)" -ForegroundColor Red
    exit 1
}
