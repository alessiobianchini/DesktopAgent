param(
    [string]$Configuration = "Release",
    [string]$Runtime = "win-x64",
    [string]$DistRoot = "dist/win-x64",
    [string]$InstallerOutput = "dist/installer",
    [string]$Version = "0.1.0",
    [switch]$SkipPublish
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

$repoRoot = Split-Path -Parent $PSScriptRoot
$distPath = Join-Path $repoRoot $DistRoot
$installerOutPath = Join-Path $repoRoot $InstallerOutput
$issPath = Join-Path $repoRoot "installer/DesktopAgent.iss"
$publishScript = Join-Path $repoRoot "scripts/publish-windows.ps1"

if (-not (Test-Path $issPath)) {
    throw "Inno Setup script not found: $issPath"
}

if (-not $SkipPublish) {
    Invoke-Checked -FileName "powershell" -Arguments @(
        "-NoProfile",
        "-ExecutionPolicy", "Bypass",
        "-File", $publishScript,
        "-Configuration", $Configuration,
        "-Runtime", $Runtime,
        "-OutputRoot", $DistRoot,
        "-Version", $Version
    ) -WorkingDirectory $repoRoot
}

if (-not (Test-Path (Join-Path $distPath "start-desktopagent.ps1"))) {
    throw "Publish output missing start script in: $distPath"
}

$iscc = Get-Command "iscc" -ErrorAction SilentlyContinue
if (-not $iscc) {
    throw "Inno Setup compiler (iscc) not found in PATH. Install with: winget install JRSoftware.InnoSetup"
}

New-Item -ItemType Directory -Path $installerOutPath -Force | Out-Null

Invoke-Checked -FileName $iscc.Source -Arguments @(
    "/DMyAppVersion=$Version",
    "/DDistDir=$distPath",
    "/DInstallerOut=$installerOutPath",
    $issPath
) -WorkingDirectory $repoRoot

Write-Host "Installer created in: $installerOutPath"
