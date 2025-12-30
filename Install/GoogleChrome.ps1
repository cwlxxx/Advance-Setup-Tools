Write-Host "Installing Google Chrome via winget..."
winget install --id Google.Chrome --source winget --exact

if ($LASTEXITCODE -eq 0) {
    Write-Host "Chrome installed successfully via winget."
    return
}

Write-Host "Winget failed. Falling back to official Chrome installer..."

$msiUrl  = "https://dl.google.com/dl/chrome/install/googlechromestandaloneenterprise64.msi"
$msiPath = "$env:TEMP\GoogleChromeEnterprise.msi"

try {
    Invoke-WebRequest -Uri $msiUrl -OutFile $msiPath
    Write-Host "Download completed. Launching installer..."
    
    Start-Process -FilePath $msiPath -Wait

    Write-Host "Chrome installer finished."
}
catch {
    Write-Host "Failed to download or run Chrome installer."
    Write-Host $_
    return
}
finally {
    if (Test-Path $msiPath) {
        Remove-Item $msiPath -Force
        Write-Host "Temporary installer removed."
    }
}
