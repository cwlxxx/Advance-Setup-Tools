# =========================================================
# Section : Sogou Pinyin Installer (Minimal Edition)
# Author  : Liang Edition
# Target  : PowerShell 7+
# Desc    : Installs Sogou Pinyin via Winget.
#           Ensures English (US) exists and is set as default.
# =========================================================

try {
    Write-Host "🔍 Checking for existing Sogou installation..." -ForegroundColor Cyan
    $SogouInstalled = winget list | Select-String -Pattern "Sogou Pinyin"

    if (-not $SogouInstalled) {
        Write-Host "➕ Installing Sogou Pinyin via Winget..." -ForegroundColor Yellow
        winget install --id Sogou.SogouInput --source winget --exact --accept-package-agreements --accept-source-agreements
        Write-Host "✅ Sogou Pinyin installation completed." -ForegroundColor Green
    } else {
        Write-Host "✅ Sogou Pinyin already installed." -ForegroundColor Green
    }

    # --- Step 2 : Adjust Windows language list ---
    Write-Host "⚙️ Checking current language list..." -ForegroundColor Cyan
    $langList = Get-WinUserLanguageList

    # --- Step 3 : Ensure English (US) is present and default ---
    $usLang = $langList | Where-Object { $_.LanguageTag -eq 'en-US' }
    if ($usLang) {
        Write-Host "✅ English (US) found — setting as default."
        $langList = @($usLang) + ($langList | Where-Object { $_.LanguageTag -ne 'en-US' })
    } else {
        Write-Host "⚠️ English (US) not found — adding manually."
        $usLang = New-WinUserLanguageList en-US
        $langList = @($usLang) + $langList
    }

    # --- Step 4 : Apply updated language list ---
    Set-WinUserLanguageList $langList -Force | Out-Null
    Write-Host "✅ English (US) set as default successfully." -ForegroundColor Green

    # --- Step 5 : Restart input service (for immediate effect) ---
    Write-Host "🔄 Restarting input service..." -ForegroundColor Cyan
    Start-Process -FilePath "cmd.exe" -ArgumentList '/c taskkill /im ctfmon.exe /f' -WindowStyle Hidden | Out-Null
    Start-Sleep -Milliseconds 800
    Start-Process -FilePath "$env:SystemRoot\System32\ctfmon.exe" -WindowStyle Hidden | Out-Null

    # --- Step 6 : Show summary ---
    Write-Host "`n📋 Final Language Layouts:" -ForegroundColor Cyan
    Get-WinUserLanguageList | ForEach-Object { Write-Host " - $($_.LanguageTag)" }

    Write-Host "`n✅ Setup complete — Sogou Pinyin installed, English (US) is default." -ForegroundColor Cyan
}
catch {
    Write-Host "❌ Error during setup: $($_.Exception.Message)" -ForegroundColor Red
}
# =========================================================
# End of Sogou Pinyin Installer (Minimal Edition)
# =========================================================
