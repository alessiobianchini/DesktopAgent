param(
    [string]$Configuration = "Release",
    [string]$Runtime = "win-x64",
    [string]$OutputRoot = "dist/win-x64",
    [string]$Version = "0.1.0",
    [switch]$NoSingleFile,
    [switch]$FrameworkDependent,
    [switch]$SkipTests
)

$ErrorActionPreference = "Stop"

function Invoke-Checked {
    param(
        [string]$FileName,
        [string[]]$Arguments,
        [string]$WorkingDirectory
    )

    Write-Host ">> $FileName $($Arguments -join ' ')"
    $process = Start-Process -FilePath $FileName -ArgumentList $Arguments -WorkingDirectory $WorkingDirectory -NoNewWindow -PassThru -Wait
    if ($process.ExitCode -ne 0) {
        throw "Command failed with exit code $($process.ExitCode): $FileName $($Arguments -join ' ')"
    }
}

function Write-TextFile {
    param(
        [string]$Path,
        [string]$Text
    )

    $dir = Split-Path -Parent $Path
    if (-not [string]::IsNullOrWhiteSpace($dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }

    Set-Content -Path $Path -Value $Text -Encoding UTF8
}

function Resolve-AssemblyVersion {
    param([string]$VersionLabel)

    $normalized = $VersionLabel.Trim().TrimStart('v', 'V')
    $match = [regex]::Match($normalized, '^(?<major>\d+)\.(?<minor>\d+)\.(?<patch>\d+)(?:\.(?<rev>\d+))?')
    if (-not $match.Success) {
        return "1.0.0.0"
    }

    $major = $match.Groups["major"].Value
    $minor = $match.Groups["minor"].Value
    $patch = $match.Groups["patch"].Value
    $rev = if ($match.Groups["rev"].Success) { $match.Groups["rev"].Value } else { "0" }
    return "$major.$minor.$patch.$rev"
}

$repoRoot = Split-Path -Parent $PSScriptRoot
$output = Join-Path $repoRoot $OutputRoot
$adapterOut = Join-Path $output "adapter"
$trayOut = Join-Path $output "tray"

if (-not $SkipTests) {
    Invoke-Checked -FileName "dotnet" -Arguments @("test", "core/DesktopAgent.Tests/DesktopAgent.Tests.csproj", "-c", $Configuration) -WorkingDirectory $repoRoot
}

if (Test-Path $output) {
    Remove-Item -Recurse -Force $output
}

New-Item -ItemType Directory -Path $adapterOut -Force | Out-Null
New-Item -ItemType Directory -Path $trayOut -Force | Out-Null

$singleFile = if ($NoSingleFile) { "false" } else { "true" }
$selfContained = if ($FrameworkDependent) { "false" } else { "true" }
$normalizedVersion = $Version.Trim().TrimStart('v', 'V')
if ([string]::IsNullOrWhiteSpace($normalizedVersion)) {
    $normalizedVersion = "0.1.0"
}
$assemblyVersion = Resolve-AssemblyVersion -VersionLabel $normalizedVersion

$commonPublishArgs = @(
    "-c", $Configuration,
    "-r", $Runtime,
    "--self-contained", $selfContained,
    "-p:PublishSingleFile=$singleFile",
    "-p:Version=$normalizedVersion",
    "-p:InformationalVersion=$normalizedVersion",
    "-p:AssemblyVersion=$assemblyVersion",
    "-p:FileVersion=$assemblyVersion"
)

$adapterPublishArgs = @("publish", "adapters/windows/DesktopAgent.Adapter.Windows/DesktopAgent.Adapter.Windows.csproj") + $commonPublishArgs + @("-o", $adapterOut)
$trayPublishArgs = @("publish", "core/DesktopAgent.Tray/DesktopAgent.Tray.csproj") + $commonPublishArgs + @("-o", $trayOut)

Invoke-Checked -FileName "dotnet" -Arguments $adapterPublishArgs -WorkingDirectory $repoRoot
Invoke-Checked -FileName "dotnet" -Arguments $trayPublishArgs -WorkingDirectory $repoRoot

Copy-Item -Path (Join-Path $repoRoot "core/DesktopAgent.Tray/appsettings.json") -Destination (Join-Path $trayOut "appsettings.json") -Force
Copy-Item -Path (Join-Path $repoRoot "core/DesktopAgent.Cli/appsettings.json") -Destination (Join-Path $trayOut "agentsettings.json") -Force

$startScript = @'
param(
    [int]$AdapterPort = 51877,
    [switch]$NoTray,
    [switch]$ShowConsoles
)

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$adapterExe = Join-Path $root "adapter\DesktopAgent.Adapter.Windows.exe"
$trayExe = Join-Path $root "tray\DesktopAgent.Tray.exe"
$trayAgentConfig = Join-Path $root "tray\agentsettings.json"
$pidFile = Join-Path $root ".desktopagent.pids.json"

function Test-PortOpen {
    param([int]$Port)
    try {
        $client = New-Object System.Net.Sockets.TcpClient
        $iar = $client.BeginConnect("127.0.0.1", $Port, $null, $null)
        $ok = $iar.AsyncWaitHandle.WaitOne(250)
        if ($ok) {
            $null = $client.EndConnect($iar)
            $client.Close()
            return $true
        }
        $client.Close()
        return $false
    }
    catch {
        return $false
    }
}

function Start-AgentProcess {
    param(
        [string]$Name,
        [string]$FilePath,
        [string]$WorkingDirectory,
        [switch]$HideWindow
    )

    if (-not (Test-Path $FilePath)) {
        throw "$Name executable not found: $FilePath"
    }

    Write-Host "Starting $Name..."
    if ($HideWindow -and -not $ShowConsoles) {
        return Start-Process -FilePath $FilePath -WorkingDirectory $WorkingDirectory -WindowStyle Hidden -PassThru
    }

    return Start-Process -FilePath $FilePath -WorkingDirectory $WorkingDirectory -PassThru
}

$pids = @{}

if (-not (Test-PortOpen -Port $AdapterPort)) {
    $env:DESKTOP_AGENT_PORT = "$AdapterPort"
    $adapter = Start-AgentProcess -Name "Adapter" -FilePath $adapterExe -WorkingDirectory (Split-Path -Parent $adapterExe) -HideWindow
    $pids.Adapter = $adapter.Id
} else {
    Write-Host "Adapter already listening on port $AdapterPort"
}

if (-not $NoTray) {
    $runningTray = Get-Process -Name "DesktopAgent.Tray" -ErrorAction SilentlyContinue
    if (-not $runningTray) {
        $env:DESKTOP_AGENT_TRAY_ADAPTERENDPOINT = "http://localhost:$AdapterPort"
        $env:DESKTOP_AGENT_TRAY_AGENTCONFIGPATH = $trayAgentConfig
        $env:DESKTOP_AGENT_TRAY_AUTOSTARTWEB = "false"
        $tray = Start-AgentProcess -Name "Tray" -FilePath $trayExe -WorkingDirectory (Split-Path -Parent $trayExe)
        $pids.Tray = $tray.Id
    } else {
        Write-Host "Tray already running"
    }
}

$json = $pids | ConvertTo-Json
Set-Content -Path $pidFile -Value $json -Encoding UTF8

Write-Host "DesktopAgent started (tray-only mode)."
'@

$stopScript = @'
$ErrorActionPreference = "SilentlyContinue"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$pidFile = Join-Path $root ".desktopagent.pids.json"

if (Test-Path $pidFile) {
    try {
        $raw = Get-Content $pidFile -Raw
        $data = $raw | ConvertFrom-Json
        foreach ($value in $data.PSObject.Properties.Value) {
            if ($value) {
                Stop-Process -Id ([int]$value) -Force -ErrorAction SilentlyContinue
            }
        }
    }
    catch {
    }

    Remove-Item $pidFile -Force -ErrorAction SilentlyContinue
}

Get-Process -Name "DesktopAgent.Adapter.Windows","DesktopAgent.Tray" -ErrorAction SilentlyContinue |
    Stop-Process -Force -ErrorAction SilentlyContinue

Write-Host "DesktopAgent stopped."
'@

$startCmd = "@echo off`r`npowershell -NoProfile -ExecutionPolicy Bypass -File `"%~dp0start-desktopagent.ps1`" %*`r`n"
$stopCmd = "@echo off`r`npowershell -NoProfile -ExecutionPolicy Bypass -File `"%~dp0stop-desktopagent.ps1`" %*`r`n"

Write-TextFile -Path (Join-Path $output "start-desktopagent.ps1") -Text $startScript
Write-TextFile -Path (Join-Path $output "stop-desktopagent.ps1") -Text $stopScript
Write-TextFile -Path (Join-Path $output "start-desktopagent.cmd") -Text $startCmd
Write-TextFile -Path (Join-Path $output "stop-desktopagent.cmd") -Text $stopCmd

Write-Host "Publish completed: $output"
