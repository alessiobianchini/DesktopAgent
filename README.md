# DesktopAgent

DesktopAgent is a local, safety-first desktop automation agent with a cross-platform core and OS-specific adapters.

- Core: `.NET 9` (`core/DesktopAgent.Core`)
- UI: Tray app (`core/DesktopAgent.Tray`)
- Adapters: Windows (`.NET`), macOS (`Swift`), Linux (`Python`)
- Transport: gRPC (`proto/desktop_adapter.proto`)

## What It Does

- Automates desktop actions through accessibility/UI tree when available.
- Falls back to vision/OCR flows when UI tree is unavailable.
- Supports natural-language commands (rule-based + optional local LLM rewrite).
- Keeps full local audit logs and safety guardrails.

## Safety Guardrails

- `DISARMED` by default
- app/window allowlist enforcement
- dangerous action confirmation (`submit`, `send`, `pay`, `delete`, ...)
- kill switch
- rate limiting
- quiz/exam safe mode (explain-only policy hooks)
- dry-run planning mode

## Repo Layout

- `core/` core libraries, CLI, tray, tests
- `adapters/` windows / macos / linux adapter servers
- `proto/` gRPC contracts
- `scripts/` local start/publish/packaging scripts
- `docs/` architecture, security, setup guides
- `.github/workflows/` CI/package/release workflows

## Quick Start (Windows)

From repo root:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/start-windows.ps1
```

This starts the Windows adapter and tray app.

## Packaging

- Windows (portable + optional Inno installer):  
  `scripts/publish-windows.ps1` / `scripts/build-installer.ps1`
- Linux:  
  `scripts/publish-linux.sh`
- macOS:  
  `scripts/publish-macos.sh`
- Windows Velopack package (auto-update capable):  
  `scripts/package-velopack-windows.ps1 -Version 0.5.4`

## GitHub Actions

- `Package Windows` (push + manual)
- `Package Linux (Manual)`
- `Package macOS (Manual)`
- `Package Windows Velopack (Manual)`
- `Release (GitHub)` (tag-based release flow)

Artifacts are available in each workflow run under **Actions → Artifacts**.

## Auto-Update Notes

- True auto-update requires a Velopack install/feed (`RELEASES` + `.nupkg` + `Setup.exe`).
- Inno/zip installs do not provide full Velopack update lifecycle.
- Tray config includes:
  - `AutoUpdateEnabled`
  - `AutoUpdateSource`
  - `AutoUpdateCheckIntervalMinutes`
  - `AutoUpdateAutoApply`

## Configuration

- Agent config template: `core/DesktopAgent.Cli/appsettings.json`
- Tray config: `core/DesktopAgent.Tray/appsettings.json`
- Runtime writable config is stored under user profile (to avoid `Program Files` permission issues).

## Documentation

- `docs/architecture.md`
- `docs/security.md`
- `docs/setup-windows.md`
- `docs/setup-macos.md`
- `docs/setup-linux.md`

---

DesktopAgent is intended for accessibility, UI testing, and productivity automation.  
It is not designed for cheating or auto-submitting answers in quiz/exam contexts.
