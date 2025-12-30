Add-Type -Name Window -Namespace Console -MemberDefinition @"
    [DllImport("kernel32.dll")]
    public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
"@

function Hide-Console {
    $consolePtr = [Console.Window]::GetConsoleWindow()
    if ($consolePtr -ne [IntPtr]::Zero) {
        [Console.Window]::ShowWindow($consolePtr, 0)
    }
}

function Show-Console {
    $consolePtr = [Console.Window]::GetConsoleWindow()
    if ($consolePtr -ne [IntPtr]::Zero) {
        [Console.Window]::ShowWindow($consolePtr, 5)
    }
}

$TempRoot = Join-Path $env:TEMP "AdvanceSetupTools"
$ZipPath  = Join-Path $TempRoot "AdvanceSetupTools.zip"
$ExePath  = Join-Path $TempRoot "AdvanceSetupTools.exe"

if (-not (Test-Path $TempRoot)) {
    New-Item -ItemType Directory -Path $TempRoot | Out-Null
}

$Url = "https://github.com/cwlxxx/Advance-Setup-Tools/releases/download/latest/AdvanceSetupTools.zip"
Write-Host "Downloading latest AdvanceSetupTools.zip..." -ForegroundColor Cyan

try {
    Invoke-WebRequest -Uri $Url -OutFile $ZipPath -UseBasicParsing
    Write-Host "Download complete: $ZipPath" -ForegroundColor Green
} catch {
    Write-Host "Failed to download AdvanceSetupTools.zip" -ForegroundColor Red
    exit
}

Write-Host "Extracting files..." -ForegroundColor Cyan
try {
    Expand-Archive -Path $ZipPath -DestinationPath $TempRoot -Force
    Write-Host "Extracted to: $TempRoot" -ForegroundColor Green
} catch {
    Write-Host "Failed to extract ZIP file." -ForegroundColor Red
    exit
}

if (-not (Test-Path $ExePath)) {
    Write-Host "Could not find AdvanceSetupTools.exe after extraction." -ForegroundColor Red
    exit
}

Write-Host "Launching AdvanceSetupTools.exe..." -ForegroundColor Cyan

Hide-Console

Start-Process -FilePath $ExePath -Wait

Show-Console

Write-Host "Cleaning up temporary files..." -ForegroundColor DarkGray

try {
    Remove-Item -Recurse -Force $TempRoot
    Write-Host "Clean-up completed." -ForegroundColor Green
} catch {
    Write-Host "Could not remove temporary folder: $TempRoot" -ForegroundColor Yellow
}

Write-Host "All tasks completed. Exiting PowerShell." -ForegroundColor Cyan
Start-Sleep -Seconds 2
exit
