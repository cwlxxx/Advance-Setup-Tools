Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic
$currentName = $env:COMPUTERNAME

function Get-NewComputerName {
    param ($currentName)

    while ($true) {
        $message = "Current PC Name: $currentName`n`nEnter a New PC Name:"
        $newName = [Microsoft.VisualBasic.Interaction]::InputBox(
            $message,
            "Rename Computer",
            ""
        )

        if ([string]::IsNullOrWhiteSpace($newName)) {
            [System.Windows.Forms.MessageBox]::Show(
                "No name entered. Operation cancelled.",
                "Cancelled",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
            return $null
        }

        if ($newName -ieq $currentName) {
            [System.Windows.Forms.MessageBox]::Show(
                "The new name cannot be the same as the current computer name ('$currentName').`nPlease enter a different name.",
                "Invalid Name",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            ) | Out-Null
            continue
        }

        $validNamePattern = '^(?!-)(?!.*--)[A-Za-z0-9-]{1,15}(?<!-)$'
        if ($newName -notmatch $validNamePattern) {
            [System.Windows.Forms.MessageBox]::Show(
                "Invalid computer name.`n`nRules:
 - 1–15 characters only
 - Letters, numbers, and hyphens (-) allowed
 - Cannot start or end with a hyphen
 - No spaces or special characters (\ / : * ? "" < > |)",
                "Invalid Name Format",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null
            continue
        }
        return $newName
    }
}

$newName = Get-NewComputerName -currentName $currentName
if (-not $newName) {
    Write-Host "User cancelled operation."
    return
}

[System.Windows.Forms.MessageBox]::Show(
    "The computer name will be changed from '$currentName' to '$newName'.`n`nA restart is required to apply this change.`n(PC will NOT restart automatically.)",
    "Confirm Rename",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Information
) | Out-Null

try {
    Rename-Computer -NewName $newName -Force
    [System.Windows.Forms.MessageBox]::Show(
        "Computer name changed successfully to '$newName'.`nPlease restart manually to apply the change.",
        "Success",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null
}
catch {
    [System.Windows.Forms.MessageBox]::Show(
        "Error: $($_.Exception.Message)",
        "Rename Failed",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
}

Write-Host "Returning to main script... (PC not restarted)"
