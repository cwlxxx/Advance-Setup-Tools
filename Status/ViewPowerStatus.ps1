function Get-PowerStatus {
    try {
        $schemeLine = powercfg /getactivescheme 2>$null
        if ($schemeLine -match '([0-9a-fA-F\-]{36})') {
            $schemeGuid = $matches[1]
        } else {
            throw "Unable to detect active power scheme."
        }
        $basePath = "HKLM:\SYSTEM\CurrentControlSet\Control\Power\User\PowerSchemes\$schemeGuid"

        $videoKey = Join-Path $basePath "7516b95f-f776-4464-8c53-06167f40cc99\3c0bc021-c8a8-4e07-a973-6b14cbcb2b7e"
        $sleepKey = Join-Path $basePath "238C9FA8-0AAD-41ED-83F4-97BE242C8F20\29F6C1DB-86DA-48C5-9FDB-F2B67B1F44DA"

        $videoValue = Get-ItemProperty -Path $videoKey -ErrorAction SilentlyContinue
        $sleepValue = Get-ItemProperty -Path $sleepKey -ErrorAction SilentlyContinue

        if (-not $videoValue -or -not $sleepValue) {
            throw "Cannot read registry values (try running PowerShell as Administrator)."
        }

        $formatTimeout = {
            param($secs)
            if ($null -eq $secs) { return "Unknown" }
            if ($secs -eq 0) { return "Never" }
            $mins = [math]::Floor($secs / 60)
            if ($mins -eq 0) { return "$secs sec" }
            elseif ($mins -eq 1) { return "1 min" }
            else { return "$mins min" }
        }

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
