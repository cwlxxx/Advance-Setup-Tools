$DownloadUrl   = "https://github.com/cwlxxx/Advance-Setup-Tools/raw/refs/heads/main/Install/avira_installer.zip"
$ZipName       = "avira_installer.zip"
$TargetDir     = Join-Path $env:TEMP "installer"
$ZipPath       = Join-Path $TargetDir $ZipName
$ExtractPath   = Join-Path $TargetDir "Extracted"
$InstallerExe  = Join-Path $ExtractPath "avira_installer.exe"

function Ensure-Directory {
    if (-not (Test-Path $TargetDir)) {
        Write-Host "Creating working directory: $TargetDir" -ForegroundColor Cyan
        New-Item -ItemType Directory -Path $TargetDir -Force | Out-Null
    }
}

function Download-Zip {
    param ($Url, $OutFile)

    try {
        Write-Host "Downloading Avira package..." -ForegroundColor Cyan

        if (Test-Path $OutFile) {
            Remove-Item $OutFile -Force
        }

        Start-BitsTransfer -Source $Url -Destination $OutFile -TransferType Download -ErrorAction Stop

        if (Test-Path $OutFile) {
            Write-Host "✅ Download completed!" -ForegroundColor Green
            return $true
        }
        else {
            throw "ZIP missing after download."
        }
    }
    catch {
        Write-Host "❌ Download failed: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

function Extract-Zip {
    try {
        Write-Host "Extracting ZIP..." -ForegroundColor Cyan

        if (Test-Path $ExtractPath) {
            Remove-Item $ExtractPath -Recurse -Force
        }

        Expand-Archive -Path $ZipPath -DestinationPath $ExtractPath -Force

        if (Test-Path $InstallerExe) {
            Write-Host "✅ Extraction successful." -ForegroundColor Green
            return $true
        }
        else {
            throw "Installer not found after extraction."
        }
    }
    catch {
        Write-Host "❌ Extraction failed: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

function Install-Avira {
    try {
        Write-Host "Starting silent installation..." -ForegroundColor Cyan

        $Arguments = "/silent"

        $process = Start-Process -FilePath $InstallerExe -ArgumentList $Arguments -Wait -PassThru -ErrorAction Stop

        if ($process.ExitCode -eq 0) {
            Write-Host "✅ Avira installed successfully." -ForegroundColor Green
        }
        else {
            Write-Host "⚠️ Installer exit code: $($process.ExitCode)" -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "❌ Installation failed: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Cleanup {
    try {
        Write-Host "Cleaning up..." -ForegroundColor Cyan
        if (Test-Path $TargetDir) {
            Remove-Item -Path $TargetDir -Recurse -Force
        }
    }
    catch {
        Write-Host "⚠️ Cleanup failed." -ForegroundColor Yellow
    }
}

# ===== EXECUTION =====

Ensure-Directory

if (Download-Zip -Url $DownloadUrl -OutFile $ZipPath) {
    if (Extract-Zip) {
        Install-Avira
        Cleanup
        Write-Host "`n✅ All tasks completed." -ForegroundColor Green
    }
}
else {
    Write-Host "`n❌ Download failed. Aborting." -ForegroundColor Red
}
