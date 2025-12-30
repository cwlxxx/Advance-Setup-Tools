if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "Please run this script as Administrator." -ForegroundColor Yellow
    pause
    exit
}
Clear-Host
Write-Host "============================================" -ForegroundColor Cyan
Write-Host "     Microsoft Office Full Cleanup Tool     " -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host ""
$TempFolder = "$env:TEMP\officeuninstall"
if (Test-Path $TempFolder) { Remove-Item $TempFolder -Recurse -Force -ErrorAction SilentlyContinue }
New-Item -Path $TempFolder -ItemType Directory -Force | Out-Null
Write-Host "Scanning for Microsoft Office installations..." -ForegroundColor Yellow
$OfficeList = @()
try {
    $ctr = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -ErrorAction SilentlyContinue
    if ($ctr) {
        $ver = $ctr.ProductReleaseIds -join ", "
        $path = $ctr.ClientFolder
        $OfficeList += [PSCustomObject]@{Name="Microsoft Office (Click-to-Run)"; Path=$path; Type="ClickToRun"}
    }
} catch {}
try {
    $uninstKeys = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall","HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    foreach ($key in $uninstKeys) {
        Get-ChildItem $key -ErrorAction SilentlyContinue | ForEach-Object {
            $disp = (Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue).DisplayName
            if ($disp -and $disp -match "Microsoft Office") {
                $lower = $disp.ToLower()
                if ($lower -notmatch "add[- ]?in" -and
                    $lower -notmatch "plugin" -and
                    $lower -notmatch "dopdf" -and
                    $lower -notmatch "pdf" -and
                    $lower -notmatch "compatibility pack" -and
                    $lower -notmatch "visual studio" -and
                    $lower -notmatch "onenote importer") {
                    $OfficeList += [PSCustomObject]@{
                        Name = $disp
                        Path = (Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue).InstallLocation
                        Type = "MSI"
                    }
                }
            }
        }
    }
} catch {}
$possiblePaths = @(
    "C:\Program Files\Microsoft Office",
    "C:\Program Files (x86)\Microsoft Office"
)
foreach ($p in $possiblePaths) {
    if (Test-Path $p) {
        $OfficeList += [PSCustomObject]@{Name="Microsoft Office (Folder Detected)"; Path=$p; Type="FolderOnly"}
    }
}
if ($OfficeList.Count -eq 0) {
    Write-Host "`nNo Microsoft Office installations detected." -ForegroundColor Yellow
    pause
    exit
}

Write-Host "`nFound $($OfficeList.Count) Office installation(s)."
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$form = New-Object System.Windows.Forms.Form
$form.Text = "Microsoft Office Cleanup"
$form.Size = New-Object System.Drawing.Size(500,420)
$form.StartPosition = "CenterScreen"
$form.TopMost = $true
$selectAllButton = New-Object System.Windows.Forms.Button
$selectAllButton.Text = "Select All"
$selectAllButton.Location = New-Object System.Drawing.Point(20,10)
$selectAllButton.Size = New-Object System.Drawing.Size(100,25)
$form.Controls.Add($selectAllButton)
$selectNoneButton = New-Object System.Windows.Forms.Button
$selectNoneButton.Text = "Select None"
$selectNoneButton.Location = New-Object System.Drawing.Point(130,10)
$selectNoneButton.Size = New-Object System.Drawing.Size(100,25)
$form.Controls.Add($selectNoneButton)
$checkedList = New-Object System.Windows.Forms.CheckedListBox
$checkedList.Location = New-Object System.Drawing.Point(20,40)
$checkedList.Size = New-Object System.Drawing.Size(440,260)
$checkedList.CheckOnClick = $true
foreach ($item in $OfficeList) { [void]$checkedList.Items.Add($item.Name) }
$form.Controls.Add($checkedList)
$okButton = New-Object System.Windows.Forms.Button
$okButton.Text = "Uninstall"
$okButton.Location = New-Object System.Drawing.Point(150,320)
$okButton.Size = New-Object System.Drawing.Size(100,30)
$form.Controls.Add($okButton)
$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Text = "Cancel"
$cancelButton.Location = New-Object System.Drawing.Point(270,320)
$cancelButton.Size = New-Object System.Drawing.Size(100,30)
$form.Controls.Add($cancelButton)
$selectedItems = New-Object System.Collections.ArrayList
$selectAllButton.Add_Click({
    for ($i = 0; $i -lt $checkedList.Items.Count; $i++) {
        $checkedList.SetItemChecked($i, $true)
    }
})
$selectNoneButton.Add_Click({
    for ($i = 0; $i -lt $checkedList.Items.Count; $i++) {
        $checkedList.SetItemChecked($i, $false)
    }
})
$okButton.Add_Click({
    foreach ($index in $checkedList.CheckedIndices) {
        [void]$selectedItems.Add($OfficeList[$index])
    }
    $form.Close()
})
$cancelButton.Add_Click({
    $form.Tag = "Cancelled"
    $form.Close()
})
[void]$form.ShowDialog()
if ($form.Tag -eq "Cancelled") {
    Write-Host "`nCancelled by user. Exiting..." -ForegroundColor Yellow
    pause
    return
}
if ($selectedItems.Count -eq 0) {
    Write-Host "`nNo selections made. Exiting..." -ForegroundColor Yellow
    pause
    exit
}
Write-Host "`nStarting cleanup for selected Office installations..."
Start-Sleep 1
foreach ($item in $selectedItems) {
    Write-Host "`n--------------------------------------------"
    Write-Host "Removing: $($item.Name)" -ForegroundColor Cyan
    Write-Host "Type: $($item.Type)" -ForegroundColor DarkGray
    try {
        Write-Host "→ Stopping Office processes..."
        Get-Process WINWORD,EXCEL,OUTLOOK,POWERPNT,ONENOTE,MSACCESS,OfficeClickToRun -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
        if ($item.Type -eq "ClickToRun") {
            Write-Host "→ Running Click-to-Run uninstaller..."
            $setupExe = Join-Path $item.Path "OfficeClickToRun.exe"
            if (Test-Path $setupExe) {
                Start-Process -FilePath $setupExe -ArgumentList "/update user displaylevel=false forceappshutdown=true" -Wait
            } else {
                Write-Host "  (ClickToRun uninstaller not found, skipping...)" -ForegroundColor Yellow
            }
        } elseif ($item.Type -eq "MSI") {
            Write-Host "Attempting MSI uninstall..."
            $uninstKey = Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall" -Recurse -ErrorAction SilentlyContinue |
                Where-Object {($_.GetValue("DisplayName") -match $item.Name) -and ($_.GetValue("UninstallString"))} |
                Select-Object -First 1
            if ($uninstKey) {
                $cmd = $uninstKey.GetValue("UninstallString")
                Start-Process -FilePath "cmd.exe" -ArgumentList "/c $cmd /quiet /norestart" -Wait
            } else {
                Write-Host "  (No MSI uninstall key found, skipping...)" -ForegroundColor Yellow
            }
        }
        Write-Host "Removing leftover directories..."
        Remove-Item "C:\Program Files\Microsoft Office" -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item "C:\Program Files (x86)\Microsoft Office" -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item "$env:ProgramData\Microsoft\Office" -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item "$env:AppData\Microsoft\Office" -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item "$env:LocalAppData\Microsoft\Office" -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "Cleaning registry keys..."
        Remove-Item -Path "HKLM:\SOFTWARE\Microsoft\Office" -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item -Path "HKCU:\Software\Microsoft\Office" -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item -Path "HKLM:\SOFTWARE\Microsoft\ClickToRun" -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "Cleaning leftover uninstall entries..." -ForegroundColor DarkGray
        $uninstallPaths = @(
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
            "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
        )
        foreach ($path in $uninstallPaths) {
            Get-ChildItem $path -ErrorAction SilentlyContinue | ForEach-Object {
                try {
                    $name = (Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue).DisplayName
                    if ($name -and $name -match "Microsoft Office") {
                        $installLocation = (Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue).InstallLocation
                        if (-not (Test-Path $installLocation)) {
                            Remove-Item $_.PSPath -Recurse -Force -ErrorAction SilentlyContinue
                        }
                    }
                } catch {}
            }
        }
        $wordPath = "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"
        $wordPathX86 = "C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE"
        if (!(Test-Path $wordPath) -and !(Test-Path $wordPathX86)) {
            Write-Host "$($item.Name) fully removed (no executable remains)." -ForegroundColor Green
        } else {
            Write-Host "Some Office files still exist on disk, manual cleanup may be needed." -ForegroundColor Yellow
        }
        Write-Host "$($item.Name) removed successfully." -ForegroundColor Green
    } catch {
        Write-Host "Error removing $($item.Name): $_" -ForegroundColor Red
    }
}
Write-Host "`n============================================" -ForegroundColor Cyan
Write-Host " Office Uninstall Summary" -ForegroundColor Cyan
Write-Host "============================================" -ForegroundColor Cyan
Write-Host " Total Detected : $($OfficeList.Count)"
Write-Host " Removed        : $($selectedItems.Count)"
Write-Host " Temp Folder    : Removed"
Write-Host "============================================" -ForegroundColor Cyan
Remove-Item $TempFolder -Recurse -Force -ErrorAction SilentlyContinue
Write-Host ""
Read-Host "Press [Enter] to exit..."
exit
