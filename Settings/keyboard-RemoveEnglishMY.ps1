try {
    Write-Host "Checking current keyboard language list..." -ForegroundColor Cyan
    $langList = Get-WinUserLanguageList
    $usLang = $langList | Where-Object { $_.LanguageTag -eq 'en-US' }
    if ($usLang) {
        Write-Host "English (US) found — setting as default." -ForegroundColor Green
        $langList = @($usLang) + ($langList | Where-Object { $_.LanguageTag -ne 'en-US' })
    } else {
        Write-Host "English (US) not found — adding it manually." -ForegroundColor Yellow
        $usLang = New-WinUserLanguageList en-US
        $langList = @($usLang) + $langList
    }

    $beforeCount = $langList.Count
    $langList = $langList | Where-Object { $_.LanguageTag -ne 'en-MY' }
    $afterCount = $langList.Count
    if ($beforeCount -ne $afterCount) {
        Write-Host "Removed English (Malaysia) from keyboard list." -ForegroundColor Yellow
    }

    Set-WinUserLanguageList $langList -Force
    Write-Host "Keyboard language list updated successfully." -ForegroundColor Green

    Write-Host "Refreshing input language state..." -ForegroundColor Cyan

    Start-Process -FilePath "cmd.exe" -ArgumentList '/c taskkill /im ctfmon.exe /f' -WindowStyle Hidden | Out-Null
    Start-Sleep -Seconds 1
    Start-Process -FilePath "$env:SystemRoot\System32\ctfmon.exe" -WindowStyle Hidden

    Write-Host "Input language refreshed — changes should appear instantly." -ForegroundColor Green

    Write-Host "`nFinal keyboard layout list:"
    Get-WinUserLanguageList | ForEach-Object {
        Write-Host " - $($_.LanguageTag)"
    }
}
catch {
    Write-Host "Error updating keyboard language list: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "`nOperation complete — English (US) is default, English (Malaysia) removed, no reboot needed." -ForegroundColor Cyan

