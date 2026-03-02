param(
    [Parameter(Mandatory = $true)]
    [string]$CertPath,
    [Parameter(Mandatory = $true)]
    [string]$CertPassword,
    [string]$InputRoot = "dist/win-x64",
    [string]$InstallerGlob = "dist/installer/*.exe",
    [string]$TimestampUrl = "http://timestamp.digicert.com",
    [string]$DigestAlgorithm = "sha256",
    [switch]$IncludeInstaller
)

$ErrorActionPreference = "Stop"

function Find-SignTool {
    $cmd = Get-Command "signtool.exe" -ErrorAction SilentlyContinue
    if ($cmd) {
        return $cmd.Source
    }

    $candidates = New-Object System.Collections.Generic.List[string]
    $candidates.Add("$env:ProgramFiles(x86)\Windows Kits\10\bin\x64\signtool.exe")
    $candidates.Add("$env:ProgramFiles(x86)\Windows Kits\10\App Certification Kit\signtool.exe")
    $candidates.Add("$env:ProgramFiles\Microsoft SDKs\ClickOnce\SignTool\signtool.exe")

    $kitBinRoot = Join-Path "${env:ProgramFiles(x86)}" "Windows Kits\10\bin"
    if (Test-Path $kitBinRoot) {
        Get-ChildItem -Path $kitBinRoot -Directory -ErrorAction SilentlyContinue |
            Sort-Object Name -Descending |
            ForEach-Object {
                $x64 = Join-Path $_.FullName "x64\signtool.exe"
                $x86 = Join-Path $_.FullName "x86\signtool.exe"
                $arm64 = Join-Path $_.FullName "arm64\signtool.exe"
                $candidates.Add($x64)
                $candidates.Add($x86)
                $candidates.Add($arm64)
            }
    }

    foreach ($candidate in $candidates) {
        if (Test-Path $candidate) {
            return $candidate
        }
    }

    throw "signtool.exe not found. Install Windows SDK or ensure signtool is in PATH."
}

function Invoke-Checked {
    param(
        [string]$FileName,
        [string[]]$Arguments
    )

    Write-Host ">> $FileName $($Arguments -join ' ')"
    $process = Start-Process -FilePath $FileName -ArgumentList $Arguments -NoNewWindow -PassThru -Wait
    if ($process.ExitCode -ne 0) {
        throw "Command failed with exit code $($process.ExitCode): $FileName $($Arguments -join ' ')"
    }
}

if (-not (Test-Path $CertPath)) {
    throw "Certificate not found: $CertPath"
}

$signTool = Find-SignTool
$targets = New-Object System.Collections.Generic.List[string]

if (Test-Path $InputRoot) {
    Get-ChildItem -Path $InputRoot -Filter "*.exe" -Recurse |
        ForEach-Object { $targets.Add($_.FullName) }
}

if ($IncludeInstaller) {
    Get-ChildItem -Path $InstallerGlob -ErrorAction SilentlyContinue |
        ForEach-Object { $targets.Add($_.FullName) }
}

$targets = $targets | Sort-Object -Unique
if (-not $targets -or $targets.Count -eq 0) {
    Write-Host "No EXE files found to sign."
    exit 0
}

foreach ($target in $targets) {
    $signature = Get-AuthenticodeSignature -FilePath $target -ErrorAction SilentlyContinue
    if ($signature -and $signature.Status -eq "Valid") {
        Write-Host "Skipping already signed file: $target"
        continue
    }

    Invoke-Checked -FileName $signTool -Arguments @(
        "sign",
        "/f", $CertPath,
        "/p", $CertPassword,
        "/fd", $DigestAlgorithm,
        "/tr", $TimestampUrl,
        "/td", $DigestAlgorithm,
        "/v",
        $target
    )
}

Write-Host "Code signing completed."
