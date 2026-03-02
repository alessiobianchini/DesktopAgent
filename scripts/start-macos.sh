#!/usr/bin/env bash
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
ADAPTER_PORT="${DESKTOP_AGENT_PORT:-51877}"
USE_TRAY=1
BUILD_FIRST=0
NOHUP_MODE=1

while [[ $# -gt 0 ]]; do
  case "$1" in
    --no-tray) USE_TRAY=0; shift ;;
    --build) BUILD_FIRST=1; shift ;;
    --foreground) NOHUP_MODE=0; shift ;;
    *) echo "Unknown option: $1"; exit 1 ;;
  esac
done

if [[ "$(uname)" != "Darwin" ]]; then
  echo "This script is for macOS only."
  exit 1
fi

if [[ $BUILD_FIRST -eq 1 ]]; then
  dotnet build "$ROOT/core/DesktopAgent.Tray/DesktopAgent.Tray.csproj" -c Debug
fi

PID_FILE="$ROOT/.desktopagent-macos.pids"
rm -f "$PID_FILE"

start_bg() {
  local label="$1"
  shift
  if [[ $NOHUP_MODE -eq 1 ]]; then
    nohup "$@" >"/tmp/${label}.log" 2>&1 &
  else
    "$@" &
  fi
  local pid=$!
  echo "$label:$pid" >> "$PID_FILE"
  echo "Started $label (pid=$pid)"
}

start_bg "adapter-macos" env \
  DESKTOP_AGENT_PORT="$ADAPTER_PORT" \
  bash -lc "cd '$ROOT/adapters/macos/DesktopAgent.Adapter.Mac' && swift run"

if [[ $USE_TRAY -eq 1 ]]; then
  start_bg "tray" env \
    DESKTOP_AGENT_TRAY_ADAPTERENDPOINT="http://localhost:${ADAPTER_PORT}" \
    DESKTOP_AGENT_TRAY_AGENTCONFIGPATH="$ROOT/core/DesktopAgent.Cli/appsettings.json" \
    DESKTOP_AGENT_TRAY_AUTOSTARTWEB="false" \
    dotnet run --project "$ROOT/core/DesktopAgent.Tray/DesktopAgent.Tray.csproj"
fi

echo "DesktopAgent started on macOS."
echo "Use: bash scripts/stop-macos.sh"
