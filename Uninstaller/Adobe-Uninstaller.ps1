#region === Adobe Uninstaller : Scan & Cache System ===
# Temp cache path
$CachePath = Join-Path $env:TEMP "Adobe_UninstallerCache.json"

Write-Host "`n[ Adobe Uninstaller Tool — Scan Phase ]" -ForegroundColor Cyan
Write-Host "Scanning system for installed Adobe products..." -ForegroundColor Yellow

# Helper: safely read registry
function Get-RegistryValues {
    param([string]$Path)
    try {
        Get-ItemProperty -Path $Path -ErrorAction Stop
    } catch {
        $null
    }
}

# --- Scan Uninstall registry locations (64/32 bit) ---
$UninstallPaths = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
)

$Detected = @()

foreach ($path in $UninstallPaths) {
    foreach ($item in Get-ChildItem $path -ErrorAction SilentlyContinue) {
        $reg = Get-RegistryValues $item.PSPath
        if ($null -ne $reg) {
            $display = $reg.DisplayName
            $publisher = $reg.Publisher
            if ($display -and ($display -match "Adobe") -and ($display -notmatch "Reader")) {
                $Detected += [PSCustomObject]@{
                    Name            = $display
                    Version         = $reg.DisplayVersion
                    InstallPath     = $reg.InstallLocation
                    UninstallString = $reg.UninstallString
                    RegistryPath    = $item.PSPath
                }
            }
        }
    }
}

# --- Add Adobe-specific registry roots ---
$AdobeRegRoots = @(
    "HKLM:\SOFTWARE\Adobe",
    "HKCU:\SOFTWARE\Adobe"
)

foreach ($root in $AdobeRegRoots) {
    if (Test-Path $root) {
        Get-ChildItem $root -ErrorAction SilentlyContinue | ForEach-Object {
            $name = $_.PSChildName
            if ($name -notmatch "Reader") {
                $Detected += [PSCustomObject]@{
                    Name            = "Adobe $name"
                    Version         = ""
                    InstallPath     = ""
                    UninstallString = ""
                    RegistryPath    = $_.PSPath
                }
            }
        }
    }
}

# --- Add folder detections from Program Files ---
$AdobeDirs = @(
    "C:\Program Files\Adobe",
    "C:\Program Files (x86)\Adobe"
)

foreach ($dir in $AdobeDirs) {
    if (Test-Path $dir) {
        Get-ChildItem $dir -Directory -ErrorAction SilentlyContinue | ForEach-Object {
            $folder = $_.Name
            if ($folder -notmatch "Reader") {
                $Detected += [PSCustomObject]@{
                    Name            = "Adobe $folder"
                    Version         = ""
                    InstallPath     = $_.FullName
                    UninstallString = ""
                    RegistryPath    = ""
                }
            }
        }
    }
}

# --- Remove duplicates by Name ---
$Detected = $Detected | Sort-Object Name -Unique

if ($Detected.Count -eq 0) {
    Write-Host "No Adobe products found (excluding Reader DC)." -ForegroundColor DarkGray
} else {
    Write-Host "`nDetected Adobe products:`n" -ForegroundColor Green
    $Detected | ForEach-Object { Write-Host " - $($_.Name)" }
}

# --- Save cache ---
$Detected | ConvertTo-Json -Depth 4 | Set-Content -Path $CachePath -Encoding UTF8
Write-Host "`nCache file created at: $CachePath" -ForegroundColor Cyan
Write-Host "Proceed to next part (UI interface) to continue." -ForegroundColor Yellow
#endregion

#region === Adobe Uninstaller : WinForms UI Interface ===
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Write-Host "`n[ Adobe Uninstaller Tool — UI Phase ]" -ForegroundColor Cyan
Write-Host "Loading detected products from cache..." -ForegroundColor Yellow

# Read cache
if (-not (Test-Path $CachePath)) {
    [System.Windows.Forms.MessageBox]::Show("No cache file found. Please run the scan phase first.","Adobe Uninstaller",0,48)
    exit
}
$AdobeProducts = Get-Content $CachePath | ConvertFrom-Json

# Create Form
$form             = New-Object System.Windows.Forms.Form
$form.Text        = "Adobe Product Uninstaller — Safe Clean Edition"
$form.Size        = New-Object System.Drawing.Size(700,530)
$form.StartPosition = "CenterScreen"
$form.BackColor   = [System.Drawing.Color]::FromArgb(40,40,40)
$form.ForeColor   = [System.Drawing.Color]::White
$form.Font        = New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Regular)

# Title Label
$labelTitle = New-Object System.Windows.Forms.Label
$labelTitle.Text = "Detected Adobe Products"
$labelTitle.AutoSize = $true
$labelTitle.Font = New-Object System.Drawing.Font("Segoe UI Semibold",12)
$labelTitle.Location = New-Object System.Drawing.Point(20,20)
$form.Controls.Add($labelTitle)

# CheckedListBox for product list
$listBox = New-Object System.Windows.Forms.CheckedListBox
$listBox.Size = New-Object System.Drawing.Size(640,300)
$listBox.Location = New-Object System.Drawing.Point(20,60)
$listBox.BackColor = [System.Drawing.Color]::FromArgb(55,55,55)
$listBox.ForeColor = [System.Drawing.Color]::White
$listBox.BorderStyle = "FixedSingle"

foreach ($p in $AdobeProducts) {
    $display = if ($p.Version) { "$($p.Name)  —  v$($p.Version)" } else { $p.Name }
    $listBox.Items.Add($display) | Out-Null
}
$form.Controls.Add($listBox)

# 🆕 Deep Clean Checkbox
$chkDeepClean = New-Object System.Windows.Forms.CheckBox
$chkDeepClean.Text = "🧹 Deep Clean (Remove Licenses & Shared Cache)"
$chkDeepClean.AutoSize = $true
$chkDeepClean.Checked = $false
$chkDeepClean.Location = New-Object System.Drawing.Point(20,370)
$form.Controls.Add($chkDeepClean)

# Buttons
$btnUninstall = New-Object System.Windows.Forms.Button
$btnUninstall.Text = "Uninstall Selected"
$btnUninstall.Size = New-Object System.Drawing.Size(180,40)
$btnUninstall.Location = New-Object System.Drawing.Point(480,430)
$btnUninstall.BackColor = [System.Drawing.Color]::FromArgb(70,130,180)
$btnUninstall.FlatStyle = "Flat"
$btnUninstall.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($btnUninstall)

$btnRefresh = New-Object System.Windows.Forms.Button
$btnRefresh.Text = "Refresh List"
$btnRefresh.Size = New-Object System.Drawing.Size(120,40)
$btnRefresh.Location = New-Object System.Drawing.Point(340,430)
$btnRefresh.BackColor = [System.Drawing.Color]::FromArgb(90,90,90)
$btnRefresh.FlatStyle = "Flat"
$btnRefresh.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($btnRefresh)

$btnExit = New-Object System.Windows.Forms.Button
$btnExit.Text = "Exit"
$btnExit.Size = New-Object System.Drawing.Size(120,40)
$btnExit.Location = New-Object System.Drawing.Point(200,430)
$btnExit.BackColor = [System.Drawing.Color]::FromArgb(120,60,60)
$btnExit.FlatStyle = "Flat"
$btnExit.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($btnExit)

# Status Label
$status = New-Object System.Windows.Forms.Label
$status.Text = "Ready."
$status.AutoSize = $true
$status.Location = New-Object System.Drawing.Point(20,480)
$form.Controls.Add($status)

# Button Events
$btnExit.Add_Click({ $form.Close() })
$btnRefresh.Add_Click({
    if (Test-Path $CachePath) {
        $AdobeProducts = Get-Content $CachePath | ConvertFrom-Json
        $listBox.Items.Clear()
        foreach ($p in $AdobeProducts) {
            $display = if ($p.Version) { "$($p.Name)  —  v$($p.Version)" } else { $p.Name }
            $listBox.Items.Add($display) | Out-Null
        }
        $status.Text = "List refreshed."
    } else {
        $status.Text = "Cache file not found."
    }
})

# When "Uninstall Selected" pressed
$btnUninstall.Add_Click({
    $selected = @()
    foreach ($index in $listBox.CheckedIndices) {
        $selected += $AdobeProducts[$index]
    }
    if ($selected.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select at least one Adobe product to uninstall.","Adobe Uninstaller",0,48)
        return
    }
    $script:SelectedProducts = $selected
    $script:DeepClean = $chkDeepClean.Checked  # 🆕 Pass checkbox value
    $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Close()
})

# Show form
$result = $form.ShowDialog()
#endregion

#region === Adobe Uninstaller : Uninstall + Cleanup ===
if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
    Write-Host "User cancelled uninstallation." -ForegroundColor DarkGray
    if (Test-Path $CachePath) { Remove-Item $CachePath -Force -ErrorAction SilentlyContinue }
    return
}

Write-Host "`n[ Adobe Uninstaller Tool — Uninstall Phase ]" -ForegroundColor Cyan
Write-Host "Starting removal of selected Adobe products..." -ForegroundColor Yellow

# Create log directory on desktop
$LogDir = Join-Path $env:USERPROFILE "Desktop\Adobe_Uninstall_Logs"
if (!(Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir | Out-Null }
$LogFile = Join-Path $LogDir ("UninstallLog_{0:yyyyMMdd_HHmmss}.txt" -f (Get-Date))

function Write-Log {
    param($msg)
    $timestamp = (Get-Date).ToString("HH:mm:ss")
    Add-Content -Path $LogFile -Value "[$timestamp] $msg"
    Write-Host $msg
}

foreach ($app in $SelectedProducts) {
    Write-Log "-------------------------------------------"
    Write-Log "Processing: $($app.Name)"

    # --- Run official uninstaller if found ---
    if ($app.UninstallString) {
        Write-Log "Executing uninstall command: $($app.UninstallString)"
        try {
            $cmd, $args = $null
            if ($app.UninstallString -match '^"(.+?)"\s*(.*)$') {
                $cmd = $matches[1]; $args = $matches[2]
            } else {
                $cmd = $app.UninstallString; $args = ""
            }
            Start-Process -FilePath $cmd -ArgumentList $args -Wait -ErrorAction Stop
            Write-Log "Uninstall completed for $($app.Name)"
        } catch {
            Write-Log "❌ Failed to execute uninstall command: $($_.Exception.Message)"
        }
    } else {
        Write-Log "No uninstall string found — skipping executable uninstall."
    }

    # --- Remove leftover folders ---
    $paths = @()
    if ($app.InstallPath -and (Test-Path $app.InstallPath)) { $paths += $app.InstallPath }
    $paths += @(
        "$env:ProgramData\Adobe\$($app.Name -replace 'Adobe\s*','')",
        "$env:APPDATA\Adobe\$($app.Name -replace 'Adobe\s*','')",
        "$env:LOCALAPPDATA\Adobe\$($app.Name -replace 'Adobe\s*','')"
    )

    foreach ($p in $paths) {
        if (Test-Path $p) {
            try {
                Remove-Item $p -Recurse -Force -ErrorAction Stop
                Write-Log "Removed folder: $p"
            } catch {
                Write-Log "Failed to remove folder: $p — $($_.Exception.Message)"
            }
        }
    }

    # --- Remove related registry keys ---
    $regPaths = @()
    if ($app.RegistryPath) { $regPaths += $app.RegistryPath }
    $regPaths += @(
        "HKLM:\SOFTWARE\Adobe\$($app.Name -replace 'Adobe\s*','')",
        "HKCU:\SOFTWARE\Adobe\$($app.Name -replace 'Adobe\s*','')"
    )

    foreach ($r in $regPaths) {
        if (Test-Path $r) {
            try {
                Remove-Item -Path $r -Recurse -Force -ErrorAction Stop
                Write-Log "Deleted registry key: $r"
            } catch {
                Write-Log "Failed to delete registry key: $r — $($_.Exception.Message)"
            }
        }
    }

    Write-Log "Cleanup finished for $($app.Name)"
}

# 🆕 Deep Clean Section
if ($DeepClean -eq $true) {
    Write-Log "-------------------------------------------"
    Write-Log "Deep Clean Mode Enabled — Removing License & Shared Cache"
    $LicensePaths = @(
        "$env:ProgramData\Adobe\SLStore",
        "$env:ProgramData\Adobe\SLCache",
        "$env:ProgramData\regid.1986-12.com.adobe",
        "$env:ProgramFiles\Common Files\Adobe\Adobe PCD",
        "$env:ProgramFiles(x86)\Common Files\Adobe\Adobe PCD"
    )
    foreach ($lp in $LicensePaths) {
        if (Test-Path $lp) {
            try {
                Remove-Item $lp -Recurse -Force -ErrorAction Stop
                Write-Log "Removed license folder: $lp"
            } catch {
                Write-Log "Failed to remove license folder: $lp — $($_.Exception.Message)"
            }
        }
    }
    Write-Log "License cleanup complete."
}

Write-Log "-------------------------------------------"
Write-Log "All selected uninstall operations complete."

# --- Final cleanup ---
if (Test-Path $CachePath) {
    Remove-Item $CachePath -Force -ErrorAction SilentlyContinue
    Write-Host "`nTemporary cache removed: $CachePath" -ForegroundColor DarkGray
}

Write-Host "`nUninstall log saved at: $LogFile" -ForegroundColor Green
Write-Host "`nProcess complete. You may close this window." -ForegroundColor Cyan
[System.Windows.Forms.MessageBox]::Show("Adobe product removal complete.`nLog saved on Desktop.","Adobe Uninstaller",0,64)
#endregion
