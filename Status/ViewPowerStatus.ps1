function Get-PowerStatus {
    try {
        # --- Get the active power scheme GUID ---
        $schemeLine = powercfg /getactivescheme 2>$null
        if ($schemeLine -match '([0-9a-fA-F\-]{36})') {
            $schemeGuid = $matches[1]
        } else {
            throw "Unable to detect active power scheme."
        }

        # --- Registry base path for active scheme ---
        $basePath = "HKLM:\SYSTEM\CurrentControlSet\Control\Power\User\PowerSchemes\$schemeGuid"

        # --- Registry keys for Display Off and Sleep timeouts ---
        $videoKey = Join-Path $basePath "7516b95f-f776-4464-8c53-06167f40cc99\3c0bc021-c8a8-4e07-a973-6b14cbcb2b7e"   # SUB_VIDEO\VIDEOIDLE
        $sleepKey = Join-Path $basePath "238C9FA8-0AAD-41ED-83F4-97BE242C8F20\29F6C1DB-86DA-48C5-9FDB-F2B67B1F44DA"   # SUB_SLEEP\STANDBYIDLE

        # --- Read registry values safely ---
        $videoValue = Get-ItemProperty -Path $videoKey -ErrorAction SilentlyContinue
        $sleepValue = Get-ItemProperty -Path $sleepKey -ErrorAction SilentlyContinue

        if (-not $videoValue -or -not $sleepValue) {
            throw "Cannot read registry values (try running PowerShell as Administrator)."
        }

        # --- Helper: format timeout values into human readable form ---
        $formatTimeout = {
            param($secs)
            if ($null -eq $secs) { return "Unknown" }
            if ($secs -eq 0) { return "Never" }
            $mins = [math]::Floor($secs / 60)
            if ($mins -eq 0) { return "$secs sec" }
            elseif ($mins -eq 1) { return "1 min" }
            else { return "$mins min" }
        }

        # --- Return as simple object (only what we need) ---
        return [pscustomobject]@{
            Monitor = & $formatTimeout $videoValue.ACSettingIndex
            Sleep   = & $formatTimeout $sleepValue.ACSettingIndex
        }
    }
    catch {
        return [pscustomobject]@{
            Monitor = "Error"
            Sleep   = "Error"
        }
    }
}
$status = Get-PowerStatus
#Write-Host "Monitor : $($status.Monitor)"
#Write-Host "Sleep   : $($status.Sleep)"
