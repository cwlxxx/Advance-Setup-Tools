# AdvanceSetupTools - Web Installer Bootstrapper
# Author: CWL

Add-Type -Name Window -Namespace Console -MemberDefinition @"
    [DllImport("kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
"@

function Hide-Console {
    $consolePtr = [Console.Window]::GetConsoleWindow()
    if ($consolePtr -ne [IntPtr]::Zero) {
        [Console.Window]::ShowWindow($consolePtr, 0)  # 0 = SW_HIDE
    }
}

function Show-Console {
    $consolePtr = [Console.Window]::GetConsoleWindow()
    if ($consolePtr -ne [IntPtr]::Zero) {
        [Console.Window]::ShowWindow($consolePtr, 5)  # 5 = SW_SHOW
    }
}

# region ───── Helper: Safe folder setup ─────
$TempRoot = Join-Path $env:TEMP "AdvanceSetupTools"
$ZipPath  = Join-Path $TempRoot "AdvanceSetupTools.zip"
$ExePath  = Join-Path $TempRoot "AdvanceSetupTools.exe"

if (-not (Test-Path $TempRoot)) {
    New-Item -ItemType Directory -Path $TempRoot | Out-Null
}
# endregion

# region ───── Step 1: Check .NET 8 Desktop Runtime ─────
Write-Host "🔍 Checking for .NET 8 Desktop Runtime..." -ForegroundColor Cyan

$dotnetInstalled = $false
try {
    $versionOutput = & dotnet --list-runtimes 2>$null
    if ($versionOutput -match "Microsoft\.WindowsDesktop\.App 8\.") {
        $dotnetInstalled = $true
        Write-Host "✅ .NET 8 Desktop Runtime found." -ForegroundColor Green
    }
} catch {
    Write-Host "⚠️ dotnet not found in PATH." -ForegroundColor Yellow
}

if (-not $dotnetInstalled) {
    Write-Host "🚀 Installing .NET 8 Desktop Runtime via winget..." -ForegroundColor Cyan
    try {
        winget install --id Microsoft.DotNet.DesktopRuntime.8 -e --accept-package-agreements --accept-source-agreements -h
    } catch {
        Write-Host "❌ Failed to install .NET 8 Desktop Runtime automatically." -ForegroundColor Red
        Write-Host "Please install manually from: https://dotnet.microsoft.com/en-us/download/dotnet/8.0" -ForegroundColor Yellow
        Pause
        exit
    }
}
# endregion

# region ───── Step 2: Download Release.zip ─────
$Url = "https://github.com/cwlxxx/Advance-Setup-Tools/releases/download/latest/AdvanceSetupTools.zip"
Write-Host "⬇️ Downloading latest AdvanceSetupTools.zip..." -ForegroundColor Cyan

try {
    Invoke-WebRequest -Uri $Url -OutFile $ZipPath -UseBasicParsing
    Write-Host "✅ Download complete: $ZipPath" -ForegroundColor Green
} catch {
    Write-Host "❌ Failed to download AdvanceSetupTools.zip" -ForegroundColor Red
    exit
}
# endregion

# region ───── Step 3: Extract ZIP ─────
Write-Host "📦 Extracting files..." -ForegroundColor Cyan
try {
    Expand-Archive -Path $ZipPath -DestinationPath $TempRoot -Force
    Write-Host "✅ Extracted to: $TempRoot" -ForegroundColor Green
} catch {
    Write-Host "❌ Failed to extract ZIP file." -ForegroundColor Red
    exit
}
# endregion

# region ───── Step 4: Run the app and hide console ─────
if (-not (Test-Path $ExePath)) {
    Write-Host "❌ Could not find AdvanceSetupTools.exe after extraction." -ForegroundColor Red
    exit
}

Write-Host "🚀 Launching AdvanceSetupTools.exe..." -ForegroundColor Cyan

# 🪄 Hide PowerShell window while the app runs
Hide-Console

# Start EXE (wait until closed)
Start-Process -FilePath $ExePath -Wait

# Restore console after EXE closes
Show-Console

Write-Host "🧹 Cleaning up temporary files..." -ForegroundColor DarkGray
# endregion

# region ───── Step 5: Cleanup ─────
try {
    Remove-Item -Recurse -Force $TempRoot
    Write-Host "✅ Clean-up completed." -ForegroundColor Green
} catch {
    Write-Host "⚠️ Could not remove temporary folder: $TempRoot" -ForegroundColor Yellow
}

Write-Host "🎉 All tasks completed. Exiting PowerShell." -ForegroundColor Cyan
Start-Sleep -Seconds 2
exit
# endregion
