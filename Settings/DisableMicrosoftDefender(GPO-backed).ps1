# ==============================
# Disable Microsoft Defender (GPO-backed)
# ==============================

$defenderPolicy = "HKLM:\SOFTWARE\Policies\Microsoft\Windows Defender"
$rtpPolicy      = "$defenderPolicy\Real-Time Protection"

# Create policy paths
New-Item -Path $defenderPolicy -Force | Out-Null
New-Item -Path $rtpPolicy -Force | Out-Null

# Main GPO switch (same as gpedit.msc)
New-ItemProperty `
  -Path $defenderPolicy `
  -Name "DisableAntiSpyware" `
  -Value 1 `
  -PropertyType DWORD `
  -Force

# Disable real-time components
New-ItemProperty -Path $rtpPolicy -Name "DisableRealtimeMonitoring" -Value 1 -PropertyType DWORD -Force
New-ItemProperty -Path $rtpPolicy -Name "DisableBehaviorMonitoring" -Value 1 -PropertyType DWORD -Force
New-ItemProperty -Path $rtpPolicy -Name "DisableOnAccessProtection" -Value 1 -PropertyType DWORD -Force
New-ItemProperty -Path $rtpPolicy -Name "DisableScanOnRealtimeEnable" -Value 1 -PropertyType DWORD -Force

# Disable removable drive scanning (USB)
New-ItemProperty `
  -Path $defenderPolicy `
  -Name "DisableRemovableDriveScanning" `
  -Value 1 `
  -PropertyType DWORD `
  -Force

Write-Host "Defender GPO policy applied. Reboot required."
