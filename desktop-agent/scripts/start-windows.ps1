param(
    [int]$AdapterPort = 50051,
    [int]$WebPort = 5000,
    [switch]$StartTray,
    [switch]$NoTray,
    [switch]$Build,
    [switch]$NoBrowser,
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

function Start-ProjectWindow {
    param(
        [string]$Title,
        [string]$ProjectPath,
        [hashtable]$EnvironmentVariables,
        [string]$Root,
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

    Start-Process -FilePath "powershell" -WorkingDirectory $Root -ArgumentList @(
        "-NoExit",
        "-ExecutionPolicy", "Bypass",
        "-Command", $command
    ) | Out-Null
}

Ensure-Command -Name "dotnet"

$root = Split-Path -Parent $PSScriptRoot
$adapterProject = Join-Path $root "adapters\windows\DesktopAgent.Adapter.Windows\DesktopAgent.Adapter.Windows.csproj"
$webProject = Join-Path $root "core\DesktopAgent.Web\DesktopAgent.Web.csproj"
$trayProject = Join-Path $root "core\DesktopAgent.Tray\DesktopAgent.Tray.csproj"
$useTray = $StartTray -or (-not $NoTray)

if (-not (Test-Path $adapterProject)) { throw "Adapter project not found: $adapterProject" }
if (-not (Test-Path $webProject)) { throw "Web project not found: $webProject" }
if ($useTray -and -not (Test-Path $trayProject)) { throw "Tray project not found: $trayProject" }

if (Test-PortInUse -Port $AdapterPort) {
    Write-Warning "Port $AdapterPort is already in use. Adapter may fail to bind."
}
if (Test-PortInUse -Port $WebPort) {
    Write-Warning "Port $WebPort is already in use. Web server may fail to bind."
}

if ($Build) {
    Invoke-Checked -FileName "dotnet" -Arguments @("build", $adapterProject, "-c", "Debug") -WorkingDirectory $root
    Invoke-Checked -FileName "dotnet" -Arguments @("build", $webProject, "-c", "Debug") -WorkingDirectory $root
    if ($useTray) {
        Invoke-Checked -FileName "dotnet" -Arguments @("build", $trayProject, "-c", "Debug") -WorkingDirectory $root
    }
}

Start-ProjectWindow -Title "DesktopAgent Adapter (Windows)" -ProjectPath $adapterProject -EnvironmentVariables @{
    DESKTOP_AGENT_PORT = "$AdapterPort"
} -Root $root -DryRunMode:$DryRun

Start-ProjectWindow -Title "DesktopAgent Web" -ProjectPath $webProject -EnvironmentVariables @{
    ASPNETCORE_URLS = "http://localhost:$WebPort"
    DESKTOP_AGENT_ADAPTERENDPOINT = "http://localhost:$AdapterPort"
} -Root $root -DryRunMode:$DryRun

if ($useTray) {
    Start-ProjectWindow -Title "DesktopAgent Tray" -ProjectPath $trayProject -EnvironmentVariables @{} -Root $root -DryRunMode:$DryRun
}

if (-not $NoBrowser) {
    $url = "http://localhost:$WebPort"
    Write-Host "Opening Web UI: $url"
    if (-not $DryRun) {
        Start-Sleep -Seconds 2
        Start-Process $url | Out-Null
    }
}

Write-Host "DesktopAgent start sequence completed."
