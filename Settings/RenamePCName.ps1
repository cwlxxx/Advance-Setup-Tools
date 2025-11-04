# ========================================
# Rename-PC with Popup Prompts, Validation & Re-entry Loop
# Shows old name and prompts for new name
# Works in PowerShell 7+, Windows 10/11
# ========================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic

$currentName = $env:COMPUTERNAME

function Get-NewComputerName {
    param ($currentName)

    while ($true) {
        # --- Improved message text ---
        $message = "Current PC Name: $currentName`n`nEnter a New PC Name:"

        # Ask for new PC name (InputBox)
        $newName = [Microsoft.VisualBasic.Interaction]::InputBox(
            $message,
            "Rename Computer",
            ""
        )

        # --- Validation 1: Cancel or empty ---
        if ([string]::IsNullOrWhiteSpace($newName)) {
            [System.Windows.Forms.MessageBox]::Show(
                "No name entered. Operation cancelled.",
                "Cancelled",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
            return $null
        }

        # --- Validation 2: Same as current ---
        if ($newName -ieq $currentName) {
            [System.Windows.Forms.MessageBox]::Show(
                "The new name cannot be the same as the current computer name ('$currentName').`nPlease enter a different name.",
                "Invalid Name",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            ) | Out-Null
            continue
        }

        # --- Validation 3: Invalid characters or length ---
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

        # If we get here, the name is valid — return it
        return $newName
    }
}

# --- Get validated name ---
$newName = Get-NewComputerName -currentName $currentName
if (-not $newName) {
    Write-Host "User cancelled operation."
    return
}

# --- Confirm rename ---
[System.Windows.Forms.MessageBox]::Show(
    "The computer name will be changed from '$currentName' to '$newName'.`n`nA restart is required to apply this change.`n(PC will NOT restart automatically.)",
    "Confirm Rename",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Information
) | Out-Null

# --- Apply rename (no restart) ---
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
