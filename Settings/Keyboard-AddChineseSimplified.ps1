# =========================================================
# Section : Smart Input Language Setup (Silent Edition)
# Author  : Liang Edition (Deep Clone Fixed for PS7)
# Target  : PowerShell 7+
# =========================================================
try {
    Write-Host "🔍 Checking current keyboard language list..." -ForegroundColor Cyan
    $originalLangList = Get-WinUserLanguageList

    # ✅ Deep clone into a new editable List<T>
    $langList = [System.Collections.Generic.List[
        Microsoft.InternationalSettings.Commands.WinUserLanguage]]::new()
    foreach ($lang in $originalLangList) { $langList.Add($lang) }

    # --- Step 1 : Ensure English (US) is present and default ---
    $usLang = $langList | Where-Object { $_.LanguageTag -eq 'en-US' }
    if ($usLang) {
        Write-Host "✅ English (US) found — setting as default."
        foreach ($lang in $usLang) { $null = $langList.Remove($lang) }
        foreach ($lang in [array]$usLang) { $langList.Insert(0, $lang) }
    }
    else {
        Write-Host "⚠️ English (US) not found — adding it manually."
        $usLang = New-WinUserLanguageList en-US
        $langList.Insert(0, $usLang[0])
    }

    # --- Step 2 : Remove English (Malaysia) if present ---
    $beforeCount = $langList.Count
    $filtered = $langList | Where-Object { $_.LanguageTag -ne 'en-MY' }
    $langList = [System.Collections.Generic.List[
        Microsoft.InternationalSettings.Commands.WinUserLanguage]]::new()
    foreach ($lang in $filtered) { $langList.Add($lang) }
    if ($langList.Count -lt $beforeCount) {
        Write-Host "🗑️ Removed English (Malaysia)."
    }

    # --- Step 3 : Detect if Sogou Pinyin is installed ---
    $SogouInstalled = $false
    $imePath = "HKCU:\Software\Microsoft\CTF\TIP"
    if (Test-Path $imePath) {
        $SogouInstalled = (Get-ChildItem $imePath -Recurse -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match "Sogou" }) -ne $null
    }

    # --- Step 4 : Add Microsoft Pinyin if Sogou not installed ---
    if (-not $SogouInstalled) {
        if (-not ($langList | Where-Object { $_.LanguageTag -eq 'zh-CN' })) {
            Write-Host "➕ Adding Microsoft Pinyin (zh-CN)..."
            $zhCN = New-WinUserLanguageList zh-CN
            foreach ($lang in $zhCN) { $langList.Add($lang) }
        } else {
            Write-Host "✅ Chinese (Simplified, zh-CN) already present."
        }
    } else {
        Write-Host "⚠️ Sogou Pinyin detected — skipping Microsoft Pinyin."
    }

    # --- Step 5 : Apply the updated list ---
    Set-WinUserLanguageList $langList -Force | Out-Null
    Write-Host "✅ Language list applied successfully."

    # --- Step 6 : Restart text input service ---
    Start-Process -FilePath "cmd.exe" -ArgumentList '/c taskkill /im ctfmon.exe /f' -WindowStyle Hidden | Out-Null
    Start-Sleep -Milliseconds 800
    Start-Process -FilePath "$env:SystemRoot\System32\ctfmon.exe" -WindowStyle Hidden | Out-Null

    # --- Step 7 : Show summary ---
    Write-Host "`n📋 Final Language Layouts:"
    Get-WinUserLanguageList | ForEach-Object { Write-Host " - $($_.LanguageTag)" }

    if ($SogouInstalled) {
        Write-Host "💡 Note: Sogou detected, Microsoft Pinyin skipped."
    }

    Write-Host "`n✅ Input configuration complete. (No user action needed)"
}
catch {
    Write-Host "❌ Error configuring language setup: $($_.Exception.Message)" -ForegroundColor Red
}
# =========================================================
# End of Smart Input Language Setup (Silent Edition)
# =========================================================
