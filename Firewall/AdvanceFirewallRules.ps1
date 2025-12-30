$ScriptName    = "Advance Firewall Rules Script"
$ScriptVersion = "2.2"
$ScriptAuthor  = "Liang"
$Title         = "$ScriptName - v$ScriptVersion by $ScriptAuthor"

Write-Host "`n==============================================" -ForegroundColor DarkCyan
Write-Host "$ScriptName" -ForegroundColor Cyan
Write-Host "Version : $ScriptVersion" -ForegroundColor Gray
Write-Host "Author  : $ScriptAuthor" -ForegroundColor Gray
Write-Host "==============================================" -ForegroundColor DarkCyan

$Host.UI.RawUI.WindowTitle = $Title

$TargetFolders = @(
    "C:\Program Files\Adobe\",
    "C:\Program Files (x86)\Adobe\",
    "C:\Program Files (x86)\Common Files\Adobe\",
    "C:\Program Files\Common Files\Adobe\",
    "C:\ProgramData\Adobe\",
    "%USERPROFILE%\AppData\Local\Adobe\",
    "%USERPROFILE%\AppData\Roaming\Adobe\",
    
    "C:\Program Files(x86)\Foxit Software\",
    "C:\Program Files\Foxit Software\",
    
    "C:\Program Files\Nitro\",
    "C:\Program Files (x86)\Nitro\",
    
    "C:\Program Files\CADSoftTools\",
    "C:\ProgramData\CADSoftTools\",
    "%USERPROFILE%\AppData\Roaming\CADSoftTools\",
    
    "C:\Program Files\Autodesk\",
    "C:\Program Files (x86)\Autodesk\",
    "C:\ProgramData\Autodesk\",
    "C:\Program Files\Common Files\Autodesk Shared\",
    "C:\Program Files\Common Files\Macrovision Shared\",
    "C:\Program Files (x86)\Common Files\Autodesk Shared\",
    "%USERPROFILE%\AppData\Local\Autodesk\",
    "%USERPROFILE%\AppData\Local\Programs\Autodesk\",
    "%USERPROFILE%\AppData\Roaming\Autodesk\",
    "%USERPROFILE%\Autodesk\",

    "C:\Program Files\Chaos Group\",
    "C:\Program Files\Common Files\ChaosGroup\",
    "%USERPROFILE%\AppData\Local\ChaosGroup\",
    
    "C:\Program Files\Corel\",
    "C:\Program Files\Common Files\Protexis\",
    "C:\Program Files\Common Files\Corel\",
    "C:\ProgramData\Corel\",
    "%USERPROFILE%\AppData\Roaming\Corel\",
    
    "C:\ProgramData\SketchUp\",
    "C:\Program Files\SketchUp\",
    "%USERPROFILE%\AppData\Local\SketchUp\",
    "%USERPROFILE%\AppData\Roaming\SketchUp\",
    
    "C:\Program Files\SOLIDWORKS Corp\",
    "C:\Program Files\Common Files\SOLIDWORKS Shared\"
    "C:\Program Files (x86)\Common Files\SOLIDWORKS Installation Manager\",
    "C:\SOLIDWORKS Data\",
    "C:\ProgramData\SOLIDWORKS\",
    "C:\ProgramData\SOLIDWORKS Electrical\",
    "%USERPROFILE%\AppData\Roaming\SOLIDWORKS\"

    
)
$RulePrefix    = "Blocked By Advance Setup Tools"
$FileType      = ".exe"
$ExitDelaySec  = 5

function Test-FirewallRuleMatch {
    param (
        [Parameter(Mandatory)] [Microsoft.Management.Infrastructure.CimInstance] $Rule,
        [Parameter(Mandatory)] [string] $ExpectedName,
        [Parameter(Mandatory)] [string] $ExpectedProgram,
        [Parameter(Mandatory)] [string] $ExpectedDirection
    )

    $filter = Get-NetFirewallApplicationFilter -ErrorAction SilentlyContinue |
              Where-Object { $_.InstanceID -eq $Rule.InstanceID }
    $programPath = $filter.Program

    $match = (
        ($Rule.DisplayName -eq $ExpectedName) -and
        ($Rule.Direction   -eq $ExpectedDirection) -and
        ($Rule.Action      -eq 'Block') -and
        ($Rule.Enabled     -eq 'True') -and
        ($Rule.Profile     -eq 'Any') -and
        ($programPath      -eq $ExpectedProgram)
    )

    if (-not $match) {
        $reasons = @()
        if ($Rule.DisplayName -ne $ExpectedName)      { $reasons += "Rule name mismatch" }
        if ($Rule.Direction   -ne $ExpectedDirection) { $reasons += "Direction mismatch" }
        if ($Rule.Action      -ne 'Block')            { $reasons += "Action not Block" }
        if ($Rule.Enabled     -ne 'True')             { $reasons += "Rule disabled" }
        if ($Rule.Profile     -ne 'Any')              { $reasons += "Profile mismatch" }
        if ($programPath      -ne $ExpectedProgram)   { $reasons += "Program path mismatch" }

        return ,$false, ($reasons -join ", ")
    }

    return ,$true, "OK"
}

Write-Host "`nIndexing all existing firewall rules..." -ForegroundColor Cyan
$allRules = @()
$ruleCount = 0

Get-NetFirewallRule -ErrorAction SilentlyContinue | ForEach-Object {
    $ruleCount++
    if ($ruleCount % 50 -eq 0) {
        Write-Host "   Indexed $ruleCount rules..." -ForegroundColor DarkGray
    }
    $appFilter = Get-NetFirewallApplicationFilter -AssociatedNetFirewallRule $_ -ErrorAction SilentlyContinue
    if ($appFilter.Program) {
        $allRules += [PSCustomObject]@{
            Name    = $_.Name
            Display = $_.DisplayName
            Program = $appFilter.Program
        }
    }
}

Write-Host "Done indexing firewall rules. Total indexed: $ruleCount" -ForegroundColor Green

[int]$addedCount = 0
[int]$removedDupCount = 0

foreach ($folder in $TargetFolders) {
    Write-Host "`nScanning folder: $folder" -ForegroundColor Cyan
    $exeFiles = Get-ChildItem -Path $folder -Recurse -Filter "*$FileType" -ErrorAction SilentlyContinue

    if (-not $exeFiles) {
        Write-Host "No files found in $folder" -ForegroundColor Red
        continue
    }

    foreach ($exe in $exeFiles) {
        $exePath = $exe.FullName
        $ruleNameIn  = "$RulePrefix IN $exePath"
        $ruleNameOut = "$RulePrefix OUT $exePath"

        Write-Host "`nChecking: $exePath" -ForegroundColor Yellow

        $dupRules = $allRules | Where-Object { $_.Program -eq $exePath }
        if ($dupRules) {
            foreach ($r in $dupRules) {
                Remove-NetFirewallRule -Name $r.Name -ErrorAction SilentlyContinue
                Write-Host "   Removed duplicate rule → $($r.Display)" -ForegroundColor DarkGray
                $removedDupCount++
            }
        }

        foreach ($direction in @("Inbound", "Outbound")) {
            $ruleName = if ($direction -eq "Inbound") { $ruleNameIn } else { $ruleNameOut }

            New-NetFirewallRule -DisplayName $ruleName `
                                -Direction $direction `
                                -Action Block `
                                -Program $exePath `
                                -Profile Any `
                                -Enabled True `
                                -ErrorAction SilentlyContinue | Out-Null

            Write-Host "   [$direction] $exePath → Added new firewall rule." -ForegroundColor Green
            $addedCount++
        }
    }
}

Write-Host "`n==============================================" -ForegroundColor DarkCyan
Write-Host "Completed processing of all target folders" -ForegroundColor Cyan
Write-Host "Duplicates Removed : $removedDupCount" -ForegroundColor DarkGray
Write-Host "New Rules Added    : $addedCount" -ForegroundColor Green
Write-Host "==============================================" -ForegroundColor DarkCyan
Write-Host "`nWaiting $ExitDelaySec seconds before exit..." -ForegroundColor Gray
Start-Sleep -Seconds $ExitDelaySec
