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

## Run CLI
```
dotnet run --project desktop-agent/core/DesktopAgent.Cli/DesktopAgent.Cli.csproj -- status
```

## Wayland Notes
- Wayland restricts global input injection. Expect limited automation.
- Screenshot capture may require `xdg-desktop-portal` permissions.
 - Clipboard operations prefer `wl-clipboard` if available.

## X11 Notes
- If `import` from ImageMagick is available, the adapter can capture screenshots.
- Input injection uses `xdotool` when installed.
