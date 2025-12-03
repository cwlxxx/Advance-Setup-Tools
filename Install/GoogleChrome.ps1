Write-Host "Installing Google Chrome..."

# Try winget first
try {
    winget install --id Google.Chrome -e --silent --accept-package-agreements --accept-source-agreements
    if ($LASTEXITCODE -eq 0) {
        Write-Host "Google Chrome installed successfully via winget."
        return
    } else {
        Write-Host "Winget failed. Trying fallback method..."
    }
}
catch {
    Write-Host "Winget not available. Trying fallback method..."
}

# Fallback download URL
$ChromeURL = "https://dl.google.com/chrome/install/latest/chrome_installer.exe"
$TempFile  = "$env:TEMP\chrome_installer.exe"

try {
    Write-Host "Downloading Chrome installer..."
    Invoke-WebRequest -Uri $ChromeURL -OutFile $TempFile

    Write-Host "Running installer..."
    Start-Process $TempFile -ArgumentList "/silent", "/install" -Wait

    Write-Host "Chrome installed via fallback method."

} catch {
    Write-Host "Failed to install Google Chrome: $($_.Exception.Message)"
}

# Cleanup
if (Test-Path $TempFile) {
    Remove-Item $TempFile -Force
    Write-Host "Cleaned up temporary installer."
}

Write-Host "Chrome install script finished. Continuing..."
