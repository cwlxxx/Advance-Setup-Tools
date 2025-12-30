try {
    Write-Host "Trying Winget installation first..." -ForegroundColor DarkGray
    $wingetProcess = Start-Process -FilePath "winget.exe" `
        -ArgumentList 'install --id Softland.doPDF.11 --source winget --exact --silent --accept-package-agreements --accept-source-agreements' `
        -NoNewWindow -Wait -PassThru

    if ($wingetProcess.ExitCode -eq 0) {
        Write-Host "doPDF 11 installed successfully via Winget." -ForegroundColor Green
    }
    else {
        throw "Winget returned exit code $($wingetProcess.ExitCode)"
    }
}
catch {
    Write-Warning "Winget installation failed or hash mismatch detected. Attempting direct installer..."

    try {
        $InstallerURL  = "https://download.dopdf.com/download/setup/dopdf-11.exe"
        $TempInstaller = Join-Path $env:TEMP "dopdf-11.exe"
        Write-Host "Downloading doPDF 11 installer..." -ForegroundColor DarkGray
        Invoke-WebRequest -Uri $InstallerURL -OutFile $TempInstaller -UseBasicParsing
        Write-Host "Running direct installer (silent mode)..." -ForegroundColor DarkGray
        Start-Process -FilePath $TempInstaller -ArgumentList "/silent" -Wait
        Write-Host "doPDF 11 installed successfully via direct installer." -ForegroundColor Green
        Remove-Item $TempInstaller -Force -ErrorAction SilentlyContinue
    }
    catch {
        Write-Error "Both Winget and fallback installer failed for doPDF 11. $($_.Exception.Message)"
    }
}
