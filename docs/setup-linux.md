# Linux Setup

## Prereqs
- Python 3.10+
- gRPC runtime: `pip install -r adapters/linux/DesktopAgent.Adapter.Linux/requirements.txt`
- Accessibility stack: AT-SPI2 (varies by distro)
- Optional: `python3-pyatspi2` (Ubuntu/Debian) or `pyatspi` package for AT-SPI2 bindings
- Optional: `xdotool` for input injection on X11
- Optional: `xclip` or `wl-clipboard` for clipboard support
- Screenshot tools for X11: ImageMagick `import` command (optional)

## Run Linux Adapter
```
python adapters/linux/DesktopAgent.Adapter.Linux/server.py
```

## Run Tray (recommended)
In a second terminal:
```
DESKTOP_AGENT_TRAY_ADAPTERENDPOINT=http://localhost:51877 \
DESKTOP_AGENT_TRAY_AGENTCONFIGPATH=core/DesktopAgent.Cli/appsettings.json \
DESKTOP_AGENT_TRAY_AUTOSTARTWEB=false \
dotnet run --project core/DesktopAgent.Tray/DesktopAgent.Tray.csproj
```

## Start Everything (Linux)
```
bash scripts/start-linux.sh
```

Stop:
```
bash scripts/stop-linux.sh
```

## Run CLI
```
dotnet run --project core/DesktopAgent.Cli/DesktopAgent.Cli.csproj -- status
```

## Publish Portable (Linux)
```
bash scripts/publish-linux.sh
```

Output:
- `dist/linux-x64/adapter`
- `dist/linux-x64/tray`
- `dist/linux-x64/start-desktopagent.sh`
- `dist/linux-x64/stop-desktopagent.sh`

## Wayland Notes
- Wayland restricts global input injection. Expect limited automation.
- Screenshot capture may require `xdg-desktop-portal` permissions.
 - Clipboard operations prefer `wl-clipboard` if available.

## X11 Notes
- If `import` from ImageMagick is available, the adapter can capture screenshots.
- Input injection uses `xdotool` when installed.
