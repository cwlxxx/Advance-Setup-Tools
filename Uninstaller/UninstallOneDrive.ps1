if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}
Add-Type -AssemblyName PresentationFramework
$result = [System.Windows.MessageBox]::Show(
    "WARNING: This will perform a complete uninstallation of Microsoft OneDrive.`n`nThis affects all user profiles on this PC.`n`nAre you sure you want to proceed?",
    "Confirm Uninstall",
    [System.Windows.MessageBoxButton]::YesNo,
    [System.Windows.MessageBoxImage]::Warning
)
if ($result -ne [System.Windows.MessageBoxResult]::Yes) {
    [System.Windows.MessageBox]::Show("Uninstallation cancelled by user.", "Cancelled", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
    exit
}
Write-Host "Stopping OneDrive..." -ForegroundColor Yellow
Stop-Process -Name "OneDrive" -Force -ErrorAction SilentlyContinue | Out-Null
Write-Host "Running official OneDrive uninstaller..." -ForegroundColor Yellow
try {
    $setupPath = "$env:SystemRoot\SysWOW64\OneDriveSetup.exe"
    if (-not (Test-Path $setupPath)) {
        $setupPath = "$env:SystemRoot\System32\OneDriveSetup.exe"
    }
    if (Test-Path $setupPath) {
        Start-Process -FilePath $setupPath -ArgumentList "/uninstall" -Wait -NoNewWindow
        Write-Host "Uninstaller executed successfully." -ForegroundColor Green
    }
    else {
        Write-Warning "OneDriveSetup.exe not found — continuing cleanup."
    }
}
catch {
    Write-Error "Error running uninstaller: $_"
}
Write-Host "Cleaning leftover folders..." -ForegroundColor Yellow
$users = Get-ChildItem -Path C:\Users -Exclude Public, Default.migrated | Where-Object { $_.PSIsContainer }
foreach ($user in $users) {
    $profile = $user.FullName
    $appData = "$profile\AppData\Local\Microsoft\OneDrive"
    $folder = "$profile\OneDrive"
    if (Test-Path $folder) { Remove-Item $folder -Recurse -Force -ErrorAction SilentlyContinue }
    if (Test-Path $appData) { Remove-Item $appData -Recurse -Force -ErrorAction SilentlyContinue }
}
if (Test-Path "$env:ProgramData\Microsoft OneDrive") {
    Remove-Item "$env:ProgramData\Microsoft OneDrive" -Recurse -Force -ErrorAction SilentlyContinue
}
Write-Host "Removing from File Explorer sidebar..." -ForegroundColor Yellow
$regPath = "HKCU:\Software\Classes\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}"
if (Test-Path $regPath) {
    Set-ItemProperty -Path $regPath -Name "System.IsPinnedToNameSpaceTree" -Value 0 -ErrorAction SilentlyContinue
}
Write-Host "Cleaning registry entries..." -ForegroundColor Yellow
Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\OneDrive" -Name "DisableFileSyncNGSC" -Value 1 -ErrorAction SilentlyContinue | Out-Null
Remove-Item -Path "HKCU:\SOFTWARE\Microsoft\OneDrive" -Recurse -Force -ErrorAction SilentlyContinue | Out-Null
Remove-Item -Path "HKLM:\SOFTWARE\Microsoft\OneDrive" -Recurse -Force -ErrorAction SilentlyContinue | Out-Null
Remove-Item -Path "HKLM:\SOFTWARE\WOW6432Node\Microsoft\OneDrive" -Recurse -Force -ErrorAction SilentlyContinue | Out-Null
Stop-Process -Id $PID
