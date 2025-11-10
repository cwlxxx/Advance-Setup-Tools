<# 
.SYNOPSIS
  Completely uninstall Avira products (Avira Security/Antivirus/Launcher/VPN, etc.)
  Removes: products via uninstall strings, services/drivers, scheduled tasks,
  running processes, folders, and common registry keys.

  Run as Administrator.
#>

# --- Guard: Admin -------------------------------------------------------------
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()
).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "Please run this script as Administrator." -ForegroundColor Yellow
    exit 1
}

$ErrorActionPreference = 'Continue'
$VerbosePreference = 'Continue'
$StartTime = Get-Date
$LogPath = Join-Path $env:TEMP ("Avira_Uninstall_{0:yyyyMMdd_HHmmss}.log" -f $StartTime)

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $line = "[{0:yyyy-MM-dd HH:mm:ss}] [{1}] {2}" -f (Get-Date), $Level.ToUpper(), $Message
    Add-Content -Path $LogPath -Value $line
    if ($Level -eq "ERROR") { Write-Host $line -ForegroundColor Red }
    elseif ($Level -eq "WARN") { Write-Host $line -ForegroundColor Yellow }
    else { Write-Host $line }
}

Write-Log "=== Avira full removal started ==="
Write-Log "Log: $LogPath"

# --- Helper: Run process silently --------------------------------------------
function Invoke-Silent {
    param(
        [Parameter(Mandatory)][string]$FilePath,
        [string]$Arguments = "",
        [int]$TimeoutSec = 900
    )
    try {
        Write-Log "Running: `"$FilePath`" $Arguments"
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $FilePath
        $psi.Arguments = $Arguments
        $psi.UseShellExecute = $false
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError  = $true
        $p = [System.Diagnostics.Process]::Start($psi)
        if ($p.WaitForExit($TimeoutSec * 1000) -eq $false) {
            $p.Kill()
            Write-Log "Process timed out and was killed." "WARN"
        }
        else {
            Write-Log "Exit code: $($p.ExitCode)"
        }
    } catch {
        Write-Log "Invoke-Silent error: $($_.Exception.Message)" "ERROR"
    }
}

# --- 1) Kill Avira processes (best effort) -----------------------------------
$procNames = @(
    'Avira.ServiceHost','Avira.Spotlight.Service','Avira.Updater.Service',
    'Avira.Systray','Avira.Security','avguard','avshadow','avgnt','avcenter',
    'avscan','avconfig','avfwsvc','avmailc','avwebgrd','avwebloader','avnotifier',
    'avira','avlaunch','avupdater','avtray'
)
Get-Process | Where-Object { $procNames -contains $_.Name } | ForEach-Object {
    try {
        Write-Log "Stopping process $($_.Name) (Id=$($_.Id))"
        Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue
    } catch { Write-Log "Failed to stop $($_.Name): $($_.Exception.Message)" "WARN" }
}

# --- 2) Stop & remove services/drivers named like Avira ----------------------
function Remove-ServiceSafe {
    param([string]$Name)
    try {
        $svc = Get-Service -Name $Name -ErrorAction SilentlyContinue
        if ($svc) {
            Write-Log "Stopping service $Name"
            Stop-Service -Name $Name -Force -ErrorAction SilentlyContinue
            Write-Log "Disabling service $Name"
            sc.exe config $Name start= disabled | Out-Null
            Write-Log "Deleting service $Name"
            sc.exe delete $Name | Out-Null
        }
    } catch {
        Write-Log "Remove-ServiceSafe $Name failed: $($_.Exception.Message)" "WARN"
    }
}

# Any service with "Avira" in its name
Get-Service | Where-Object { $_.Name -match 'Avira' -or $_.DisplayName -match 'Avira' } |
    ForEach-Object { Remove-ServiceSafe -Name $_.Name }

# --- 3) Uninstall via registry uninstall strings -----------------------------
$UninstallRoots = @(
  'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
  'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
)

$targets = @()
foreach ($root in $UninstallRoots) {
    if (Test-Path $root) {
        Get-ChildItem $root | ForEach-Object {
            $k = Get-ItemProperty -Path $_.PSPath -ErrorAction SilentlyContinue
            if ($null -ne $k) {
                $isAvira =
                    ($k.DisplayName -match 'Avira') -or
                    ($k.Publisher -match 'Avira') -or
                    ($k.DisplayIcon -match 'Avira') -or
                    ($k.InstallLocation -match 'Avira')
                if ($isAvira -and $k.UninstallString) {
                    $targets += [PSCustomObject]@{
                        DisplayName     = $k.DisplayName
                        UninstallString = $k.UninstallString
                        InstallLocation = $k.InstallLocation
                        KeyPath         = $_.PSPath
                    }
                }
            }
        }
    }
}

if ($targets.Count -eq 0) {
    Write-Log "No Avira uninstall entries found in registry (may already be gone)." "WARN"
} else {
    Write-Log "Found $($targets.Count) Avira uninstall entries."
    foreach ($t in $targets) {
        Write-Log "Preparing to uninstall: $($t.DisplayName)"

        $cmd  = $t.UninstallString
        $file = $null
        $args = $null

        # Normalize MSIexec / product code cases
        if ($cmd -match 'msiexec\.exe' -or $cmd -match 'MsiExec\.exe' -or $cmd -match 'msiexec') {
            # Extract product code if present
            if ($cmd -match '{[0-9A-Fa-f-]{36}}') {
                $guid = $Matches[0]
                $file = "$env:WINDIR\System32\msiexec.exe"
                $args = "/x $guid /qn /norestart"
            } else {
                $file = "$env:WINDIR\System32\msiexec.exe"
                $args = "/x /qn /norestart"
            }
        } else {
            # Split path and args, then add silent guesses
            if ($cmd.StartsWith('"')) {
                $file = $cmd.Split('"')[1]
                $args = $cmd.Substring($cmd.IndexOf('"',1)+1).Trim()
            } else {
                $parts = $cmd.Split(' ',2)
                $file = $parts[0]
                $args = if ($parts.Count -gt 1) { $parts[1] } else { "" }
            }
            # Add common silent switches if not present
            if ($args -notmatch '/quiet' -and $args -notmatch '/S' -and $args -notmatch '/silent') {
                $args = "$args /quiet /silent /S"
            }
            if ($args -notmatch '/uninstall' -and $file -match '\.exe$') {
                $args = "/uninstall $args"
            }
        }

        if (Test-Path $file) {
            Invoke-Silent -FilePath $file -Arguments $args
        } else {
            Write-Log "Uninstall target not found: $file (from '$cmd')" "WARN"
        }
    }
}

# Optional: try winget if available (best effort)
try {
    if (Get-Command winget -ErrorAction SilentlyContinue) {
        $wingetPkgs = winget list | Select-String -Pattern '^Avira' -SimpleMatch
        if ($wingetPkgs) {
            Write-Log "winget found Avira packages; attempting removal via winget."
            winget list | ForEach-Object {
                if ($_ -match 'Avira') {
                    $name = ($_ -split '\s{2,}')[0]
                    if ($name) {
                        Write-Log "winget uninstall `"$name`""
                        winget uninstall --exact --silent --accept-source-agreements --accept-package-agreements "$name" | Out-Null
                    }
                }
            }
        }
    }
} catch { Write-Log "winget phase error: $($_.Exception.Message)" "WARN" }

# --- 4) Remove scheduled tasks ------------------------------------------------
$tasks = schtasks /Query /FO CSV /V | ConvertFrom-Csv | Where-Object {
    $_.TaskName -match 'Avira' -or $_.TaskToRun -match 'Avira'
}
foreach ($t in $tasks) {
    try {
        Write-Log "Deleting scheduled task $($t.TaskName)"
        schtasks /Delete /TN $t.TaskName /F | Out-Null
    } catch {
        Write-Log "Failed to delete task $($t.TaskName): $($_.Exception.Message)" "WARN"
    }
}

# --- 5) Delete leftover folders ----------------------------------------------
$paths = @(
    "$env:ProgramFiles\Avira",
    "$env:ProgramFiles(x86)\Avira",
    "$env:ProgramData\Avira",
    "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\Avira",
    "$env:LOCALAPPDATA\Avira",
    "$env:APPDATA\Avira",
    "$env:ProgramData\Package Cache"  # clean only Avira subfolders below
)

foreach ($p in $paths) {
    if (Test-Path $p) {
        if ($p -like "*Package Cache") {
            Get-ChildItem $p -Directory -ErrorAction SilentlyContinue |
                Where-Object { $_.FullName -match 'Avira' } |
                ForEach-Object {
                    Write-Log "Removing cache folder $($_.FullName)"
                    Remove-Item $_.FullName -Recurse -Force -ErrorAction SilentlyContinue
                }
        } else {
            Write-Log "Removing folder $p"
            Remove-Item $p -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
}

# --- 6) Registry cleanup ------------------------------------------------------
$regKeys = @(
    'HKLM:\SOFTWARE\Avira',
    'HKLM:\SOFTWARE\WOW6432Node\Avira',
    'HKCU:\Software\Avira'
)
foreach ($rk in $regKeys) {
    if (Test-Path $rk) {
        Write-Log "Deleting registry key $rk"
        Remove-Item $rk -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# Remove Avira uninstall entries that might remain
foreach ($root in $UninstallRoots) {
    if (Test-Path $root) {
        Get-ChildItem $root | ForEach-Object {
            $k = Get-ItemProperty -Path $_.PSPath -ErrorAction SilentlyContinue
            if ($k -and ($k.DisplayName -match 'Avira' -or $k.Publisher -match 'Avira')) {
                Write-Log "Deleting orphaned uninstall entry: $($k.DisplayName)"
                Remove-Item $_.PSPath -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
    }
}

# Remove residual Avira-named services (registry)
Get-ChildItem 'HKLM:\SYSTEM\CurrentControlSet\Services' |
    Where-Object { $_.PSChildName -match 'Avira' } |
    ForEach-Object {
        Write-Log "Deleting service registry key: $($_.PSChildName)"
        Remove-Item $_.PSPath -Recurse -Force -ErrorAction SilentlyContinue
    }

# --- 7) Network drivers / filters (best effort, Avira Web/Firewall) ----------
# Often tied to services above. Nothing specific to remove here if services are gone.

# --- Wrap up ------------------------------------------------------------------
Write-Log "=== Avira removal finished. Elapsed: $([int](New-TimeSpan -Start $StartTime -End (Get-Date)).TotalSeconds) sec) ==="
Write-Host "`nCleanup log saved to: $LogPath" -ForegroundColor Cyan
Write-Host "It's recommended to **reboot now** to release locked files and finalize removal." -ForegroundColor Green
