# Requires PowerShell 7+ and Administrator privileges
# Set the display (screen) to never turn off
powercfg /change monitor-timeout-ac 0

# Set the computer to never sleep
powercfg /change standby-timeout-ac 0

Write-Output "Power settings updated successfully for desktop. Screen and sleep set to 'Never'."
