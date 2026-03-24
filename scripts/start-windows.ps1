param(
    [int]$AdapterPort = 51877,
    [int]$WebPort = 51878, # legacy: ignored in tray-only mode
    [switch]$StartTray,
    [switch]$NoTray,
    [switch]$ShowConsoles,
    [switch]$Build,
    [switch]$NoBrowser, # legacy: ignored in tray-only mode
    [switch]$DryRun
)

$ErrorActionPreference = "Stop"

function Ensure-Command {
    param([string]$Name)
    if (-not (Get-Command $Name -ErrorAction SilentlyContinue)) {
        throw "Required command '$Name' was not found in PATH."
    }
}

function Test-PortInUse {
    param([int]$Port)
    try {
        $listener = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Loopback, $Port)
        $listener.Start()
        $listener.Stop()
        return $false
    }
    catch {
        return $true
    }
}

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

function Start-ProjectProcess {
    param(
        [string]$Title,
        [string]$ProjectPath,
        [hashtable]$EnvironmentVariables,
        [string]$Root,
        [switch]$VisibleConsole,
        [switch]$DryRunMode
    )

    $rootEscaped = $Root.Replace("'", "''")
    $projectEscaped = $ProjectPath.Replace("'", "''")

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add("`$host.UI.RawUI.WindowTitle = '$($Title.Replace("'", "''"))'")
    $lines.Add("Set-Location -Path '$rootEscaped'")

    foreach ($item in $EnvironmentVariables.GetEnumerator()) {
        $keyEscaped = $item.Key.Replace("'", "''")
        $valueEscaped = ($item.Value.ToString()).Replace("'", "''")
        $lines.Add("`$env:$keyEscaped = '$valueEscaped'")
    }

    $lines.Add("dotnet run --project '$projectEscaped'")
    $command = $lines -join "; "

    if ($DryRunMode) {
        Write-Host "[dry-run] powershell -NoExit -ExecutionPolicy Bypass -Command $command"
        return
    }

    $arguments = @(
        "-ExecutionPolicy", "Bypass",
        "-Command", $command
    )
    if ($VisibleConsole) {
        $arguments = @("-NoExit") + $arguments
    }

    if ($VisibleConsole) {
        Start-Process -FilePath "powershell" -WorkingDirectory $Root -ArgumentList $arguments | Out-Null
    } else {
        Start-Process -FilePath "powershell" -WorkingDirectory $Root -ArgumentList $arguments -WindowStyle Hidden | Out-Null
    }
}

Ensure-Command -Name "dotnet"

$root = Split-Path -Parent $PSScriptRoot
$adapterProject = Join-Path $root "adapters\windows\DesktopAgent.Adapter.Windows\DesktopAgent.Adapter.Windows.csproj"
$trayProject = Join-Path $root "core\DesktopAgent.Tray\DesktopAgent.Tray.csproj"
$repoAgentConfig = Join-Path $root "core\DesktopAgent.Cli\appsettings.json"
$localAgentDir = Join-Path $env:LOCALAPPDATA "DesktopAgent"
$agentConfig = Join-Path $localAgentDir "agentsettings.json"
$useTray = $StartTray -or (-not $NoTray)

if (-not (Test-Path $adapterProject)) { throw "Adapter project not found: $adapterProject" }
if ($useTray -and -not (Test-Path $trayProject)) { throw "Tray project not found: $trayProject" }

if ($useTray) {
    New-Item -Path $localAgentDir -ItemType Directory -Force | Out-Null
    if (-not (Test-Path $agentConfig)) {
        if (Test-Path $repoAgentConfig) {
            Copy-Item -Path $repoAgentConfig -Destination $agentConfig -Force
            Write-Host "Seeded local agent config: $agentConfig"
        } else {
            throw "Agent config not found. Expected one of: $agentConfig, $repoAgentConfig"
        }
    }
}

if (Test-PortInUse -Port $AdapterPort) {
    Write-Warning "Port $AdapterPort is already in use. Adapter may fail to bind."
}

if ($Build) {
    Invoke-Checked -FileName "dotnet" -Arguments @("build", $adapterProject, "-c", "Debug") -WorkingDirectory $root
    if ($useTray) {
        Invoke-Checked -FileName "dotnet" -Arguments @("build", $trayProject, "-c", "Debug") -WorkingDirectory $root
    }
}

Start-ProjectProcess -Title "DesktopAgent Adapter (Windows)" -ProjectPath $adapterProject -EnvironmentVariables @{
    DESKTOP_AGENT_PORT = "$AdapterPort"
} -Root $root -VisibleConsole:$ShowConsoles -DryRunMode:$DryRun

if ($useTray) {
    Start-ProjectProcess -Title "DesktopAgent Tray" -ProjectPath $trayProject -EnvironmentVariables @{
        DESKTOP_AGENT_TRAY_ADAPTERENDPOINT = "http://localhost:$AdapterPort"
        DESKTOP_AGENT_TRAY_AGENTCONFIGPATH = $agentConfig
        DESKTOP_AGENT_TRAY_AUTOSTARTWEB = "false"
    } -Root $root -VisibleConsole:$ShowConsoles -DryRunMode:$DryRun
}

Write-Host "DesktopAgent start sequence completed."
