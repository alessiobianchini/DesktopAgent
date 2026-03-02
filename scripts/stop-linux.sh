#!/usr/bin/env bash
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
PID_FILE="$ROOT/.desktopagent-linux.pids"

if [[ -f "$PID_FILE" ]]; then
  while IFS=: read -r label pid; do
    if [[ -n "${pid:-}" ]] && kill -0 "$pid" 2>/dev/null; then
      kill "$pid" 2>/dev/null || true
      echo "Stopped $label (pid=$pid)"
    fi
  done < "$PID_FILE"
  rm -f "$PID_FILE"
fi

pkill -f "DesktopAgent.Tray" 2>/dev/null || true
pkill -f "DesktopAgent.Adapter.Linux/server.py" 2>/dev/null || true

echo "DesktopAgent stopped on Linux."
