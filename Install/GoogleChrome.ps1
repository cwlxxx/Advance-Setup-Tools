Write-Host "Installing Google Chrome via winget..."

# Try winget first
winget install --id Google.Chrome --source winget --exact
if ($LASTEXITCODE -eq 0) {
    Write-Host "Chrome installed successfully via winget."
    exit 0
}

Write-Host "Winget failed. Falling back to official Chrome installer..."

# Fallback MSI
$msiUrl  = "https://dl.google.com/dl/chrome/install/googlechromestandaloneenterprise64.msi"
$msiPath = "$env:TEMP\GoogleChromeEnterprise.msi"

try {
    Invoke-WebRequest -Uri $msiUrl -OutFile $msiPath -UseBasicParsing
    Write-Host "Download completed. Launching installer..."
    
    # Normal installer run (UI)
    Start-Process -FilePath $msiPath -Wait

    Write-Host "Chrome installer finished."
}
catch {
    Write-Host "Failed to download or run Chrome installer."
    Write-Host $_
    exit 1
}
finally {
    # Cleanup MSI
    if (Test-Path $msiPath) {
        Remove-Item $msiPath -Force
        Write-Host "Temporary installer removed."
    }
}
