
function Get-FastStartupStatus {
    try {
        $key = "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power"
        $value = (Get-ItemProperty -Path $key -Name HiberbootEnabled -ErrorAction SilentlyContinue).HiberbootEnabled

        switch ($value) {
            1 { return "Enabled" }
            0 { return "Disabled" }
            default { return "Unknown" }
        }
    }
    catch {
        return "Error"
    }
}
$faststartupstatus = Get-FastStartupStatus
