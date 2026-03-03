$RegPath = "HKLM:\SYSTEM\CurrentControlSet\Control\CI\Policy"
$ValueName = "VerifiedAndReputablePolicyState"
$ValueData = 0

# Ensure the key exists
if (-not (Test-Path $RegPath)) {
    New-Item -Path $RegPath -Force | Out-Null
}

# Create or update the DWORD value
New-ItemProperty -Path $RegPath `
                 -Name $ValueName `
                 -Value $ValueData `
                 -PropertyType DWord `
                 -Force | Out-Null

Write-Host "$ValueName has been set to $ValueData successfully."
