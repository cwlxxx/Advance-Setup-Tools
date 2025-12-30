try {
    Write-Host "Installing Sogou Pinyin via Winget..." -ForegroundColor Yellow
    winget install --id Sogou.SogouInput --source winget --exact `
        --accept-package-agreements --accept-source-agreements
    Write-Host "Sogou Pinyin installation step completed." -ForegroundColor Green

    Write-Host "Checking current language list..." -ForegroundColor Cyan
    $langList = Get-WinUserLanguageList

    $usLang = $langList | Where-Object { $_.LanguageTag -eq 'en-US' }

    if ($usLang) {
        Write-Host "English (US) found — setting as default."
        $langList = @($usLang) + ($langList | Where-Object { $_.LanguageTag -ne 'en-US' })
    }
    else {
        Write-Host "English (US) not found — adding manually."
        $usLang = New-WinUserLanguageList en-US
        $langList = @($usLang) + $langList
    }

    Set-WinUserLanguageList $langList -Force | Out-Null
    Write-Host "English (US) set as default successfully." -ForegroundColor Green

    Write-Host "Restarting input service..." -ForegroundColor Cyan
    Start-Process -FilePath "cmd.exe" -ArgumentList '/c taskkill /im ctfmon.exe /f' -WindowStyle Hidden | Out-Null
    Start-Sleep -Milliseconds 800
    Start-Process -FilePath "$env:SystemRoot\System32\ctfmon.exe" -WindowStyle Hidden | Out-Null

    Write-Host "Final Language Layouts:" -ForegroundColor Cyan
    Get-WinUserLanguageList | ForEach-Object { Write-Host " - $($_.LanguageTag)" }

    Write-Host "Setup complete — Sogou Pinyin installed, English (US) is default." -ForegroundColor Cyan
}
catch {
    Write-Host "Error during setup: $($_.Exception.Message)" -ForegroundColor Red
}

