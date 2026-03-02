param(
    [string]$CliArgs = "status"
)

$root = Split-Path -Parent $PSScriptRoot
$env:DESKTOP_AGENT_PORT = $env:DESKTOP_AGENT_PORT ?? "50051"

Write-Host "Starting Windows adapter on port $env:DESKTOP_AGENT_PORT"
Start-Process dotnet -ArgumentList "run --project $root\adapters\windows\DesktopAgent.Adapter.Windows\DesktopAgent.Adapter.Windows.csproj" -WorkingDirectory $root

Start-Sleep -Seconds 2
Write-Host "Running CLI: $CliArgs"
dotnet run --project $root\core\DesktopAgent.Cli\DesktopAgent.Cli.csproj -- $CliArgs
