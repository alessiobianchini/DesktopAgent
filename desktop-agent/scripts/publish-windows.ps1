param(
    [string]$Configuration = "Release",
    [string]$Runtime = "win-x64",
    [string]$OutputRoot = "dist/win-x64",
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

$repoRoot = Split-Path -Parent $PSScriptRoot
$output = Join-Path $repoRoot $OutputRoot
$adapterOut = Join-Path $output "adapter"
$webOut = Join-Path $output "web"
$trayOut = Join-Path $output "tray"

if (-not $SkipTests) {
    Invoke-Checked -FileName "dotnet" -Arguments @("test", "core/DesktopAgent.Tests/DesktopAgent.Tests.csproj", "-c", $Configuration) -WorkingDirectory $repoRoot
}

if (Test-Path $output) {
    Remove-Item -Recurse -Force $output
}

New-Item -ItemType Directory -Path $adapterOut -Force | Out-Null
New-Item -ItemType Directory -Path $webOut -Force | Out-Null
New-Item -ItemType Directory -Path $trayOut -Force | Out-Null

$singleFile = if ($NoSingleFile) { "false" } else { "true" }
$selfContained = if ($FrameworkDependent) { "false" } else { "true" }

$commonPublishArgs = @(
    "-c", $Configuration,
    "-r", $Runtime,
    "--self-contained", $selfContained,
    "-p:PublishSingleFile=$singleFile"
)

$adapterPublishArgs = @("publish", "adapters/windows/DesktopAgent.Adapter.Windows/DesktopAgent.Adapter.Windows.csproj") + $commonPublishArgs + @("-o", $adapterOut)
$webPublishArgs = @("publish", "core/DesktopAgent.Web/DesktopAgent.Web.csproj") + $commonPublishArgs + @("-o", $webOut)
$trayPublishArgs = @("publish", "core/DesktopAgent.Tray/DesktopAgent.Tray.csproj") + $commonPublishArgs + @("-o", $trayOut)

Invoke-Checked -FileName "dotnet" -Arguments $adapterPublishArgs -WorkingDirectory $repoRoot
Invoke-Checked -FileName "dotnet" -Arguments $webPublishArgs -WorkingDirectory $repoRoot
Invoke-Checked -FileName "dotnet" -Arguments $trayPublishArgs -WorkingDirectory $repoRoot

Copy-Item -Path (Join-Path $repoRoot "core/DesktopAgent.Web/appsettings.json") -Destination (Join-Path $webOut "appsettings.json") -Force
Copy-Item -Path (Join-Path $repoRoot "core/DesktopAgent.Tray/appsettings.json") -Destination (Join-Path $trayOut "appsettings.json") -Force

$startScript = @'
param(
    [int]$AdapterPort = 50051,
    [int]$WebPort = 5000,
    [switch]$NoTray,
    [switch]$NoBrowser
)

$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$adapterExe = Join-Path $root "adapter\DesktopAgent.Adapter.Windows.exe"
$webExe = Join-Path $root "web\DesktopAgent.Web.exe"
$trayExe = Join-Path $root "tray\DesktopAgent.Tray.exe"
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
        [string]$WorkingDirectory
    )

    if (-not (Test-Path $FilePath)) {
        throw "$Name executable not found: $FilePath"
    }

    Write-Host "Starting $Name..."
    return Start-Process -FilePath $FilePath -WorkingDirectory $WorkingDirectory -PassThru
}

$pids = @{}

if (-not (Test-PortOpen -Port $AdapterPort)) {
    $env:DESKTOP_AGENT_PORT = "$AdapterPort"
    $adapter = Start-AgentProcess -Name "Adapter" -FilePath $adapterExe -WorkingDirectory (Split-Path -Parent $adapterExe)
    $pids.Adapter = $adapter.Id
} else {
    Write-Host "Adapter already listening on port $AdapterPort"
}

if (-not (Test-PortOpen -Port $WebPort)) {
    $env:ASPNETCORE_URLS = "http://localhost:$WebPort"
    $env:DESKTOP_AGENT_ADAPTERENDPOINT = "http://localhost:$AdapterPort"
    $web = Start-AgentProcess -Name "Web" -FilePath $webExe -WorkingDirectory (Split-Path -Parent $webExe)
    $pids.Web = $web.Id
} else {
    Write-Host "Web already listening on port $WebPort"
}

if (-not $NoTray) {
    $runningTray = Get-Process -Name "DesktopAgent.Tray" -ErrorAction SilentlyContinue
    if (-not $runningTray) {
        $env:DESKTOP_AGENT_TRAY_ADAPTERENDPOINT = "http://localhost:$AdapterPort"
        $env:DESKTOP_AGENT_TRAY_WEBUIURL = "http://localhost:$WebPort"
        $tray = Start-AgentProcess -Name "Tray" -FilePath $trayExe -WorkingDirectory (Split-Path -Parent $trayExe)
        $pids.Tray = $tray.Id
    } else {
        Write-Host "Tray already running"
    }
}

$json = $pids | ConvertTo-Json
Set-Content -Path $pidFile -Value $json -Encoding UTF8

if (-not $NoBrowser) {
    Start-Process "http://localhost:$WebPort" | Out-Null
}

Write-Host "DesktopAgent started."
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

Get-Process -Name "DesktopAgent.Adapter.Windows","DesktopAgent.Web","DesktopAgent.Tray" -ErrorAction SilentlyContinue |
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
