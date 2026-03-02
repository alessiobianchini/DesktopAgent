#!/usr/bin/env bash
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
PORT="${DESKTOP_AGENT_PORT:-51877}"

if [[ "$(uname)" == "Darwin" ]]; then
  echo "Starting macOS adapter on port $PORT"
  (cd "$ROOT/adapters/macos/DesktopAgent.Adapter.Mac" && DESKTOP_AGENT_PORT="$PORT" swift run) &
elif [[ "$(uname)" == "Linux" ]]; then
  echo "Starting Linux adapter on port $PORT"
  (cd "$ROOT" && DESKTOP_AGENT_PORT="$PORT" python adapters/linux/DesktopAgent.Adapter.Linux/server.py) &
else
  echo "Unsupported OS for run-local.sh"
  exit 1
fi

sleep 2

dotnet run --project "$ROOT/core/DesktopAgent.Cli/DesktopAgent.Cli.csproj" -- "$@"
