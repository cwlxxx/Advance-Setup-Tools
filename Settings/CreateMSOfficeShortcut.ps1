$desktopPath = [Environment]::GetFolderPath("Desktop")
$officePaths = @(
    "$env:ProgramFiles\Microsoft Office\root\Office16",
    "$env:ProgramFiles(x86)\Microsoft Office\root\Office16" 
)

function New-OfficeShortcut {
    param(
        [string]$AppName,
        [string]$ExeName
    )

    $exePath = $null
    foreach ($path in $officePaths) {
        $candidate = Join-Path $path $ExeName
        if (Test-Path $candidate) {
            $exePath = $candidate
            break
        }
    }

    if (-not $exePath) {
        Write-Host "$AppName not found on this system." -ForegroundColor Red
        return
    }

    $shortcutPath = Join-Path $desktopPath "$AppName.lnk"
    $shell = New-Object -ComObject WScript.Shell
    $shortcut = $shell.CreateShortcut($shortcutPath)
    $shortcut.TargetPath = $exePath
    $shortcut.WorkingDirectory = Split-Path $exePath
    $shortcut.IconLocation = $exePath
    $shortcut.Save()

    Write-Host "Created shortcut for $AppName" -ForegroundColor Green
}

$apps = @(
    @{ Name = "Word";        Exe = "WINWORD.EXE" },
    @{ Name = "Excel";       Exe = "EXCEL.EXE" },
    @{ Name = "Outlook";     Exe = "OUTLOOK.EXE" },
    @{ Name = "PowerPoint";  Exe = "POWERPNT.EXE" }
)

foreach ($app in $apps) {
    New-OfficeShortcut -AppName $app.Name -ExeName $app.Exe
}

