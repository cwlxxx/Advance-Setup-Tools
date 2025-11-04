# ============================================================
# 🧱 PowerShell 7 Script : Create Microsoft Office Shortcuts
# ============================================================

# ------------------------------------------------------------
# ⚙️ Section : Setup Paths - Start
# ------------------------------------------------------------
$desktopPath = [Environment]::GetFolderPath("Desktop")

# Possible Office installation paths
$officePaths = @(
    "$env:ProgramFiles\Microsoft Office\root\Office16",      # Office 365 / 2016-2021 64-bit
    "$env:ProgramFiles(x86)\Microsoft Office\root\Office16"  # Office 365 / 2016-2021 32-bit
)
# ------------------------------------------------------------
# ⚙️ Section : Setup Paths - End
# ------------------------------------------------------------


# ------------------------------------------------------------
# 🧩 Section : Shortcut Creator Function - Start
# ------------------------------------------------------------
function New-OfficeShortcut {
    param(
        [string]$AppName,
        [string]$ExeName
    )

    # Find the correct path for this Office app
    $exePath = $null
    foreach ($path in $officePaths) {
        $candidate = Join-Path $path $ExeName
        if (Test-Path $candidate) {
            $exePath = $candidate
            break
        }
    }

    if (-not $exePath) {
        Write-Host "❌ $AppName not found on this system." -ForegroundColor Red
        return
    }

    # Create desktop shortcut
    $shortcutPath = Join-Path $desktopPath "$AppName.lnk"
    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut($shortcutPath)
    $shortcut.TargetPath = $exePath
    $shortcut.WorkingDirectory = Split-Path $exePath
    $shortcut.IconLocation = $exePath
    $shortcut.Save()

    Write-Host "✅ Created shortcut for $AppName" -ForegroundColor Green
}
# ------------------------------------------------------------
# 🧩 Section : Shortcut Creator Function - End
# ------------------------------------------------------------


# ------------------------------------------------------------
# 🚀 Section : Create Office Shortcuts - Start
# ------------------------------------------------------------
$apps = @(
    @{ Name = "Word";        Exe = "WINWORD.EXE" },
    @{ Name = "Excel";       Exe = "EXCEL.EXE" },
    @{ Name = "Outlook";     Exe = "OUTLOOK.EXE" },
    @{ Name = "PowerPoint";  Exe = "POWERPNT.EXE" }
)

foreach ($app in $apps) {
    New-OfficeShortcut -AppName $app.Name -ExeName $app.Exe
}
# ------------------------------------------------------------
# 🚀 Section : Create Office Shortcuts - End
# ------------------------------------------------------------
