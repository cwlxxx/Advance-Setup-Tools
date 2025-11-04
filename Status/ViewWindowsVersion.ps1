function Get-WindowsEdition {
    $os = Get-CimInstance Win32_OperatingSystem
    $caption = $os.Caption
    $build = [System.Environment]::OSVersion.Version.Build

    # --- Edition Detection ---
    if ($caption -match "Home") { $edition = "Home" }
    elseif ($caption -match "Pro") { $edition = "Pro" }
    elseif ($caption -match "Enterprise") { $edition = "Ent" }
    elseif ($caption -match "Education") { $edition = "Edu" }
    else { $edition = "Edition" }

    # --- Version Tag Based on Build ---
    switch ($build) {
        {$_ -ge 26200} { $version = "25H2"; break }
        {$_ -ge 26100} { $version = "24H2"; break }
        {$_ -ge 22621} { $version = "22H2"; break }
        {$_ -ge 22000} { $version = "21H2"; break }
        {$_ -ge 19045} { $version = "22H2"; break }
        {$_ -ge 19044} { $version = "21H2"; break }
        {$_ -ge 19043} { $version = "21H1"; break }
        {$_ -ge 19042} { $version = "20H2"; break }
        default { $version = "N/A" }
    }

    # --- Windows Generation ---
    if ($build -ge 22000) {
        $osVer = "Win11 $edition $version"
    } else {
        $osVer = "Win10 $edition $version"
    }

    return $osVer
}
