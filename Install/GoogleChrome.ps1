Write-Host "Installing Google Chrome via winget..."

# Run winget and capture output
$wingetOutput = winget install --id Google.Chrome --source winget --exact 2>&1
$exitCode = $LASTEXITCODE

$wingetText = $wingetOutput | Out-String
Write-Host $wingetText

# SUCCESS conditions
if (
    $exitCode -eq 0 -or
    $wingetText -match "Found an existing package already installed" -or
    $wingetText -match "No available upgrade found"
) {
    Write-Host "Chrome already installed or installed successfully."
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
    if (Test-Path $msiPath) {
        Remove-Item $msiPath -Force
        Write-Host "Temporary installer removed."
    }
}
