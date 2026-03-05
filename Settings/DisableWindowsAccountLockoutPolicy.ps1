# Disable Windows Account Lockout Policy
# Works on Windows 10 and Windows 11

Write-Host "Configuring Account Lockout Policy..."

try {

    # Disable lockout threshold
    net accounts /lockoutthreshold:0 | Out-Null

    # Optional: reset other values to safe defaults
    net accounts /lockoutduration:0 | Out-Null
    net accounts /lockoutwindow:0 | Out-Null

    Write-Host "Account lockout policy disabled successfully."

}
catch {

    Write-Host "Failed to configure account lockout policy."

}
