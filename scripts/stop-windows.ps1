param(
    [int[]]$Ports = @(51877, 51878),
    [switch]$IncludeLegacyPorts
)

$ErrorActionPreference = "Continue"

if ($IncludeLegacyPorts) {
    $Ports = ($Ports + @(5000, 50051)) | Sort-Object -Unique
}

$selfPid = $PID
$targetPids = New-Object System.Collections.Generic.HashSet[int]

function Add-TargetPid {
    param([int]$ProcessId)
    if ($ProcessId -gt 0 -and $ProcessId -ne $selfPid) {
        [void]$targetPids.Add($ProcessId)
    }
}

function Stop-TargetPid {
    param(
        [int]$ProcessId,
        [string]$Reason
    )

    try {
        $proc = Get-Process -Id $ProcessId -ErrorAction Stop
        Write-Host "Stopping PID $ProcessId ($($proc.ProcessName)) - $Reason"
        Stop-Process -Id $ProcessId -Force -ErrorAction Stop
        return $true
    }
    catch {
        return $false
    }
}

# 1) Known DesktopAgent process names
$knownNames = @(
    "DesktopAgent.Tray",
    "DesktopAgent.Adapter.Windows",
    "DesktopAgent.Cli",
    "DesktopAgent.Web"
)

foreach ($name in $knownNames) {
    Get-Process -Name $name -ErrorAction SilentlyContinue | ForEach-Object {
        Add-TargetPid -ProcessId $_.Id
    }
}

# 2) dotnet/powershell hosts started with DesktopAgent projects/scripts
$patterns = @(
    "DesktopAgent.Adapter.Windows",
    "DesktopAgent.Tray",
    "DesktopAgent.Cli",
    "DesktopAgent.Web",
    "start-windows.ps1",
    "run-local.ps1"
)

Get-CimInstance Win32_Process -ErrorAction SilentlyContinue |
    Where-Object {
        $cmd = $_.CommandLine
        if ([string]::IsNullOrWhiteSpace($cmd)) { return $false }
        foreach ($pattern in $patterns) {
            if ($cmd -match [regex]::Escape($pattern)) { return $true }
        }
        return $false
    } |
    ForEach-Object {
        Add-TargetPid -ProcessId $_.ProcessId
    }

# 3) Anything listening on DesktopAgent ports
foreach ($port in $Ports) {
    try {
        Get-NetTCPConnection -State Listen -LocalPort $port -ErrorAction Stop |
            Select-Object -ExpandProperty OwningProcess -Unique |
            ForEach-Object {
                Add-TargetPid -ProcessId $_
            }
    }
    catch {
        # Port not in use or command not available.
    }
}

if ($targetPids.Count -eq 0) {
    Write-Host "No DesktopAgent-related processes found."
    exit 0
}

$stopped = 0
foreach ($pidToStop in $targetPids) {
    if (Stop-TargetPid -ProcessId $pidToStop -Reason "DesktopAgent stop script") {
        $stopped++
    }
}

Write-Host "Stopped $stopped process(es)."
