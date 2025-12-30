if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()
).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "Please run this script as Administrator." -ForegroundColor Yellow
    exit 1
}
$ErrorActionPreference = 'Continue'
$VerbosePreference = 'Continue'
$StartTime = Get-Date
$LogPath = Join-Path $env:TEMP ("McAfee_Uninstall_{0:yyyyMMdd_HHmmss}.log" -f $StartTime)
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $line = "[{0:yyyy-MM-dd HH:mm:ss}] [{1}] {2}" -f (Get-Date), $Level.ToUpper(), $Message
    Add-Content -Path $LogPath -Value $line
    if ($Level -eq "ERROR") { Write-Host $line -ForegroundColor Red }
    elseif ($Level -eq "WARN") { Write-Host $line -ForegroundColor Yellow }
    else { Write-Host $line }
}
Write-Log "=== McAfee full removal started ==="
Write-Log "Log: $LogPath"
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
$procNames = @(
    'mcshield','mfemms','mfefire','mfemactl','mctray','mcupdate',
    'mfewc','mfeesp','mfeavsvc','mfewch','mcagent','mcods','mcsvhost',
    'McUICnt','McAfeeDLPAgent','masvc','macmnsvc','macompatsvc'
)
Get-Process | Where-Object { $procNames -contains $_.Name } | ForEach-Object {
    try {
        Write-Log "Stopping process $($_.Name) (Id=$($_.Id))"
        Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue
    } catch { Write-Log "Failed to stop $($_.Name): $($_.Exception.Message)" "WARN" }
}
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
Get-Service | Where-Object { $_.Name -match 'McAfee' -or $_.DisplayName -match 'McAfee' -or $_.Name -match 'MFE' } |
    ForEach-Object { Remove-ServiceSafe -Name $_.Name }
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
                $isMcAfee =
                    ($k.DisplayName -match 'McAfee') -or
                    ($k.Publisher -match 'McAfee') -or
                    ($k.DisplayIcon -match 'McAfee') -or
                    ($k.InstallLocation -match 'McAfee')
                if ($isMcAfee -and $k.UninstallString) {
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
    Write-Log "No McAfee uninstall entries found in registry (may already be gone)." "WARN"
} else {
    Write-Log "Found $($targets.Count) McAfee uninstall entries."
    foreach ($t in $targets) {
        Write-Log "Preparing to uninstall: $($t.DisplayName)"
        $cmd  = $t.UninstallString
        $file = $null
        $args = $null
        if ($cmd -match 'msiexec\.exe') {
            if ($cmd -match '{[0-9A-Fa-f-]{36}}') {
                $guid = $Matches[0]
                $file = "$env:WINDIR\System32\msiexec.exe"
                $args = "/x $guid /qn /norestart"
            } else {
                $file = "$env:WINDIR\System32\msiexec.exe"
                $args = "/x /qn /norestart"
            }
        } else {
            if ($cmd.StartsWith('"')) {
                $file = $cmd.Split('"')[1]
                $args = $cmd.Substring($cmd.IndexOf('"',1)+1).Trim()
            } else {
                $parts = $cmd.Split(' ',2)
                $file = $parts[0]
                $args = if ($parts.Count -gt 1) { $parts[1] } else { "" }
            }
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
$tasks = schtasks /Query /FO CSV /V | ConvertFrom-Csv | Where-Object {
    $_.TaskName -match 'McAfee' -or $_.TaskToRun -match 'McAfee'
}
foreach ($t in $tasks) {
    try {
        Write-Log "Deleting scheduled task $($t.TaskName)"
        schtasks /Delete /TN $t.TaskName /F | Out-Null
    } catch {
        Write-Log "Failed to delete task $($t.TaskName): $($_.Exception.Message)" "WARN"
    }
}
$paths = @(
    "$env:ProgramFiles\McAfee",
    "$env:ProgramFiles(x86)\McAfee",
    "$env:ProgramData\McAfee",
    "$env:LOCALAPPDATA\McAfee",
    "$env:APPDATA\McAfee",
    "$env:ProgramData\McAfee Installer",
    "$env:ProgramData\McAfee Security Scan",
    "$env:ProgramData\McAfee LiveSafe",
    "$env:ProgramData\McAfee Agent",
    "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\McAfee"
)
foreach ($p in $paths) {
    if (Test-Path $p) {
        Write-Log "Removing folder $p"
        Remove-Item $p -Recurse -Force -ErrorAction SilentlyContinue
    }
}
$regKeys = @(
    'HKLM:\SOFTWARE\McAfee',
    'HKLM:\SOFTWARE\WOW6432Node\McAfee',
    'HKCU:\Software\McAfee'
)
foreach ($rk in $regKeys) {
    if (Test-Path $rk) {
        Write-Log "Deleting registry key $rk"
        Remove-Item $rk -Recurse -Force -ErrorAction SilentlyContinue
    }
}
foreach ($root in $UninstallRoots) {
    if (Test-Path $root) {
        Get-ChildItem $root | ForEach-Object {
            $k = Get-ItemProperty -Path $_.PSPath -ErrorAction SilentlyContinue
            if ($k -and ($k.DisplayName -match 'McAfee' -or $k.Publisher -match 'McAfee')) {
                Write-Log "Deleting orphaned uninstall entry: $($k.DisplayName)"
                Remove-Item $_.PSPath -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
    }
}
Get-ChildItem 'HKLM:\SYSTEM\CurrentControlSet\Services' |
    Where-Object { $_.PSChildName -match 'McAfee' -or $_.PSChildName -match 'MFE' } |
    ForEach-Object {
        Write-Log "Deleting service registry key: $($_.PSChildName)"
        Remove-Item $_.PSPath -Recurse -Force -ErrorAction SilentlyContinue
    }
Write-Log "Checking for McAfee network filters..."
try {
    Get-NetAdapterBinding -ComponentID * | Where-Object { $_.ComponentID -match 'MFE' } | ForEach-Object {
        Disable-NetAdapterBinding -Name $_.Name -ComponentID $_.ComponentID -PassThru | Out-Null
        Write-Log "Disabled McAfee network filter: $($_.Name) / $($_.ComponentID)"
    }
} catch {
    Write-Log "Network filter cleanup skipped or failed: $($_.Exception.Message)" "WARN"
}
Write-Log "=== McAfee removal finished. Elapsed: $([int](New-TimeSpan -Start $StartTime -End (Get-Date)).TotalSeconds) sec) ==="
Write-Host "`nCleanup log saved to: $LogPath" -ForegroundColor Cyan
Write-Host "It's recommended to **reboot now** to finalize removal." -ForegroundColor Green
