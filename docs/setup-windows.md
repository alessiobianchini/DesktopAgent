# Windows Setup

## Prereqs
- .NET 9 SDK
- Windows 10/11 with UI Automation enabled (default)
- Optional: Tesseract OCR if `OcrEnabled=true`
- Optional: FFmpeg for screen recording (`record screen`, `start recording`, `stop recording`)

## Build
```
dotnet build DesktopAgentSolution.sln
```

## Run Windows Adapter
```
dotnet run --project adapters/windows/DesktopAgent.Adapter.Windows/DesktopAgent.Adapter.Windows.csproj
```

## Run CLI
```
dotnet run --project core/DesktopAgent.Cli/DesktopAgent.Cli.csproj -- status
```

## Start Everything (Windows)
Launch Adapter + Tray icon with one command (tray-only mode):
```
powershell -ExecutionPolicy Bypass -File scripts/start-windows.ps1
```

Useful options:
```
# Build first, then start Adapter + Tray
powershell -ExecutionPolicy Bypass -File scripts/start-windows.ps1 -Build

# Disable Tray app
powershell -ExecutionPolicy Bypass -File scripts/start-windows.ps1 -NoTray

# Custom adapter port
powershell -ExecutionPolicy Bypass -File scripts/start-windows.ps1 -AdapterPort 51877

# Show adapter/tray consoles (disabled by default)
powershell -ExecutionPolicy Bypass -File scripts/start-windows.ps1 -ShowConsoles
```

## Notes
- The adapter starts disarmed. Run `arm` before any actions.
- If UI automation fails due to permissions, run as administrator or check Windows security settings.

## File Search
DesktopAgent supports local file search commands (recursive):

```text
file search report
file search *.pdf in "C:\Users\<you>\Documents"
search file invoice in "D:\Work"
cerca file bolletta in .
```

Behavior:
- Search is bounded to `FilesystemAllowedRoots` (outside paths are blocked by policy).
- `*` and `?` wildcards are supported.
- Results are capped (first 200 matches).

Configure allowed roots in:
- `core/DesktopAgent.Cli/appsettings.json` (template)
- `%LocalAppData%\DesktopAgent\agentsettings.json` (runtime tray config)

## Build Portable Package
Create self-contained binaries under `dist/win-x64`:
```
powershell -ExecutionPolicy Bypass -File scripts/publish-windows.ps1
```

Output:
- `dist/win-x64/adapter`
- `dist/win-x64/tray`
- `dist/win-x64/start-desktopagent.cmd`
- `dist/win-x64/stop-desktopagent.cmd`

## Build Installer (Inno Setup)
1. Install Inno Setup compiler:
```
winget install JRSoftware.InnoSetup
```
2. Build installer:
```
powershell -ExecutionPolicy Bypass -File scripts/build-installer.ps1 -Version 0.1.0
```

Installer behavior:
- Offers optional plugin-style install tasks:
  - **FFmpeg plugin** for screen recording.
  - **OCR plugin (Tesseract)** for vision fallback.
- Installer now shows a dedicated **Optional Plugins** wizard page with descriptions and status (already installed / unavailable).
- FFmpeg task uses `winget install -e --id Gyan.FFmpeg`.
- OCR task tries `UB-Mannheim.TesseractOCR` and falls back to `tesseract-ocr.tesseract`.
- Post-install, if tools are still missing, installer shows actionable manual install hints.

Silent install plugin flags:
```
DesktopAgent-Setup-<version>.exe /VERYSILENT /installffmpeg=1 /installocr=1
```

Output:
- `dist/installer/DesktopAgent-Setup-0.1.0.exe`

## GitHub Release Automation
Repository includes workflow:
- `.github/workflows/release-tag.yml`

How to trigger:
1. Create and push a tag (example):
```
git tag v0.2.0
git push origin v0.2.0
```
2. GitHub Action builds portable zip + Velopack assets and publishes a GitHub Release.

Release assets:
- `DesktopAgent-win-x64-<version>.zip`
- `DesktopAgent-win-Setup.exe` (Velopack installer)
- `releases.win.json`
- `assets.win.json`
- `RELEASES` (legacy compatibility)
- `DesktopAgent-<version>-full.nupkg`

Manual re-sign of an existing release:
- Workflow: `.github/workflows/manual-sign.yml`
- Inputs:
  - `tag` (required)
  - `replace_assets` (true = overwrite existing release files)

## Code Signing (GitHub Release)
The release workflow can sign binaries if these repository secrets are set:
- `WINDOWS_CODESIGN_PFX_BASE64`: base64 of your `.pfx` certificate
- `WINDOWS_CODESIGN_PASSWORD`: password of the `.pfx`

Local signing command:
```powershell
powershell -ExecutionPolicy Bypass -File scripts/sign-windows.ps1 `
  -CertPath "C:\certs\codesign.pfx" `
  -CertPassword "YOUR_PASSWORD" `
  -InputRoot "dist/win-x64" `
  -IncludeInstaller
```

## Auto Updates (Tray / Velopack)
The tray app now supports automatic update checks via Velopack.

Tray config (`core/DesktopAgent.Tray/appsettings.json`):
```json
{
  "AutoUpdateEnabled": true,
  "AutoUpdateSource": "https://your-update-feed-url",
  "AutoUpdateCheckIntervalMinutes": 60,
  "AutoUpdateAutoApply": false
}
```

Notes:
- When running with `dotnet run`, status is usually `dev mode` (not installed package).
- To apply updates, the app should run from an installed/published package with a valid Velopack feed.
- User agent settings are persisted in `%LocalAppData%\DesktopAgent\agentsettings.json` (survive app updates).
- Tray menu includes:
  - `Check updates now`
  - `Apply downloaded update`
- Web UI `Configuration > Utilities` includes:
  - tool detection for `ffmpeg` and `tesseract`
  - install buttons (Windows uses `winget`)
  - `Install OCR + Enable` / `Enable OCR only`
- Tray app includes a first-run **Optional Plugins** wizard (can be reopened from tray menu: `Install Optional Plugins...`).

Build a Windows Velopack package locally:
```powershell
powershell -ExecutionPolicy Bypass -File scripts/package-velopack-windows.ps1 -Version 0.5.4
```
Output feed folder: `dist/velopack-win-x64/releases`

GitHub Actions manual packaging workflow:
- `.github/workflows/package-windows-velopack.yml`
- Run it with `version` (example `0.5.4`), then publish uploaded artifact files (`releases.win.json`, `assets.win.json`, `.nupkg`, `Setup.exe`) to your release/feed URL.
