# Windows Setup

## Prereqs
- .NET 9 SDK
- Windows 10/11 with UI Automation enabled (default)
- Optional: Tesseract OCR if `OcrEnabled=true`

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
Launch Adapter + Web UI + Tray icon with one command:
```
powershell -ExecutionPolicy Bypass -File scripts/start-windows.ps1
```

Useful options:
```
# Build first, then start Adapter + Web
powershell -ExecutionPolicy Bypass -File scripts/start-windows.ps1 -Build

# Disable Tray app
powershell -ExecutionPolicy Bypass -File scripts/start-windows.ps1 -NoTray

# Custom ports
powershell -ExecutionPolicy Bypass -File scripts/start-windows.ps1 -AdapterPort 50051 -WebPort 5000
```

## Notes
- The adapter starts disarmed. Run `arm` before any actions.
- If UI automation fails due to permissions, run as administrator or check Windows security settings.

## Build Portable Package
Create self-contained binaries under `dist/win-x64`:
```
powershell -ExecutionPolicy Bypass -File scripts/publish-windows.ps1
```

Output:
- `dist/win-x64/adapter`
- `dist/win-x64/web`
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
2. GitHub Action builds portable zip + installer and publishes a GitHub Release.

Release assets:
- `DesktopAgent-win-x64-<version>.zip`
- `DesktopAgent-Setup-<version>.exe`

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
- Tray menu includes:
  - `Check updates now`
  - `Apply downloaded update`
