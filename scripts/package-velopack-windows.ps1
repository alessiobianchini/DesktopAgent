param(
    [string]$Configuration = "Release",
    [string]$Runtime = "win-x64",
    [string]$Version = "0.1.0",
    [string]$PublishOutputRoot = "dist/win-x64",
    [string]$OutputRoot = "dist/velopack-win-x64",
    [string]$PackId = "DesktopAgent",
    [string]$PackTitle = "DesktopAgent",
    [string]$VpkVersion = "0.0.1298",
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

function Normalize-SemVer {
    param([string]$Value)

    $raw = if ($null -eq $Value) { "" } else { $Value }
    $normalized = $raw.Trim().TrimStart('v', 'V')
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return "0.1.0"
    }

    $match = [regex]::Match($normalized, '^(?<a>\d+)\.(?<b>\d+)\.(?<c>\d+)(?:\.(?<d>\d+))?(?<suffix>.*)$')
    if (-not $match.Success) {
        throw "Invalid version '$Value'. Expected semver like 0.5.4 or 0.5.4-beta.1"
    }

    $suffix = $match.Groups["suffix"].Value
    return "$($match.Groups["a"].Value).$($match.Groups["b"].Value).$($match.Groups["c"].Value)$suffix"
}

function Resolve-Vpk {
    param(
        [string]$RepoRoot,
        [string]$ToolVersion,
        [string]$OutRoot
    )

    $global = Get-Command "vpk" -ErrorAction SilentlyContinue
    if ($global) {
        return $global.Source
    }

    $toolPath = Join-Path $OutRoot ".tools\vpk"
    New-Item -ItemType Directory -Path $toolPath -Force | Out-Null

    $vpkExe = Join-Path $toolPath "vpk.exe"
    if (Test-Path $vpkExe) {
        Invoke-Checked -FileName "dotnet" -Arguments @("tool", "update", "vpk", "--tool-path", $toolPath, "--version", $ToolVersion) -WorkingDirectory $RepoRoot
    }
    else {
        Invoke-Checked -FileName "dotnet" -Arguments @("tool", "install", "vpk", "--tool-path", $toolPath, "--version", $ToolVersion) -WorkingDirectory $RepoRoot
    }

    if (-not (Test-Path $vpkExe)) {
        throw "vpk executable not found after install: $vpkExe"
    }

    return $vpkExe
}

$repoRoot = Split-Path -Parent $PSScriptRoot
$publishScript = Join-Path $repoRoot "scripts\publish-windows.ps1"
$publishRoot = Join-Path $repoRoot $PublishOutputRoot
$trayRoot = Join-Path $publishRoot "tray"
$adapterRoot = Join-Path $publishRoot "adapter"
$outputPath = Join-Path $repoRoot $OutputRoot
$stagePath = Join-Path $outputPath "staging"
$releasePath = Join-Path $outputPath "releases"
$semVersion = Normalize-SemVer -Value $Version

if (Test-Path $outputPath) {
    Remove-Item -Recurse -Force $outputPath
}

New-Item -ItemType Directory -Path $stagePath -Force | Out-Null
New-Item -ItemType Directory -Path $releasePath -Force | Out-Null

Invoke-Checked -FileName "powershell" -Arguments @(
    "-NoProfile",
    "-ExecutionPolicy", "Bypass",
    "-File", $publishScript,
    "-Configuration", $Configuration,
    "-Runtime", $Runtime,
    "-OutputRoot", $PublishOutputRoot,
    "-Version", $semVersion
) + $(if ($SkipTests) { @("-SkipTests") } else { @() }) -WorkingDirectory $repoRoot

if (-not (Test-Path $trayRoot)) {
    throw "Tray publish output not found: $trayRoot"
}

if (-not (Test-Path $adapterRoot)) {
    throw "Adapter publish output not found: $adapterRoot"
}

Copy-Item -Path (Join-Path $trayRoot "*") -Destination $stagePath -Recurse -Force
Copy-Item -Path $adapterRoot -Destination (Join-Path $stagePath "adapter") -Recurse -Force

$mainExe = Join-Path $stagePath "DesktopAgent.Tray.exe"
if (-not (Test-Path $mainExe)) {
    throw "Main tray executable not found in staging folder: $mainExe"
}

$vpk = Resolve-Vpk -RepoRoot $repoRoot -ToolVersion $VpkVersion -OutRoot $outputPath
Invoke-Checked -FileName $vpk -Arguments @(
    "pack",
    "--packId", $PackId,
    "--packVersion", $semVersion,
    "--packDir", $stagePath,
    "--mainExe", "DesktopAgent.Tray.exe",
    "--packTitle", $PackTitle,
    "--channel", "win",
    "--runtime", $Runtime,
    "--outputDir", $releasePath
) -WorkingDirectory $repoRoot

Write-Host "Velopack package created: $releasePath"
