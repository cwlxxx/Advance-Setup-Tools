## winget install --id AnyDesk.AnyDesk --source winget --exact

try {
    winget install --id AnyDesk.AnyDesk --source winget --exact
    Write-Host "Winget install completed."
}
catch {
    Write-Host "Winget failed. Using fallback installer..."

    $installerUrl  = "https://download.anydesk.com/AnyDesk.exe"
    $installerPath = "$env:TEMP\AnyDesk.exe"

    Invoke-WebRequest -Uri $installerUrl -OutFile $installerPath

    Start-Process `
        -FilePath $installerPath `
        -ArgumentList "--install `"$env:ProgramFiles\AnyDesk`" --silent" `
        -Wait

    Remove-Item $installerPath -Force -ErrorAction SilentlyContinue

    Write-Host "Fallback install completed."
}

# Script ends naturally here – next .ps1 continues
