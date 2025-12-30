
$regPath = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System'
try {
    Write-Host "Disabling User Account Control (UAC)..." -ForegroundColor Cyan
    Set-ItemProperty -Path $regPath -Name 'EnableLUA' -Value 0 -Force
    Set-ItemProperty -Path $regPath -Name 'ConsentPromptBehaviorAdmin' -Value 0 -Force
    Write-Host "UAC has been disabled (Never Notify mode)." -ForegroundColor Green
    Write-Host "A system reboot is required for changes to fully take effect." -ForegroundColor Yellow
}
catch {
    Write-Host "Failed to modify UAC settings: $($_.Exception.Message)" -ForegroundColor Red
}
