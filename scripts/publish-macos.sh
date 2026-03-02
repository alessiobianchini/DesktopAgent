#!/usr/bin/env bash
set -euo pipefail

if [[ "$(uname)" != "Darwin" ]]; then
  echo "publish-macos.sh must run on macOS."
  exit 1
fi

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
CONFIGURATION="${CONFIGURATION:-Release}"
RUNTIME="${RUNTIME:-osx-arm64}"
OUTPUT_ROOT="${OUTPUT_ROOT:-dist/macos}"

OUT="$ROOT/$OUTPUT_ROOT"
ADAPTER_OUT="$OUT/adapter"
ADAPTER_SRC_OUT="$OUT/adapter-src"
TRAY_OUT="$OUT/tray"

rm -rf "$OUT"
mkdir -p "$ADAPTER_OUT" "$ADAPTER_SRC_OUT" "$TRAY_OUT"

dotnet test "$ROOT/core/DesktopAgent.Tests/DesktopAgent.Tests.csproj" -c "$CONFIGURATION"
dotnet publish "$ROOT/core/DesktopAgent.Tray/DesktopAgent.Tray.csproj" -c "$CONFIGURATION" -r "$RUNTIME" --self-contained true -p:PublishSingleFile=true -o "$TRAY_OUT"

ADAPTER_BINARY_OK=0
pushd "$ROOT/adapters/macos/DesktopAgent.Adapter.Mac" >/dev/null
if swift build -c release; then
  if [[ -f "$ROOT/adapters/macos/DesktopAgent.Adapter.Mac/.build/release/DesktopAgentAdapterMac" ]]; then
    cp "$ROOT/adapters/macos/DesktopAgent.Adapter.Mac/.build/release/DesktopAgentAdapterMac" "$ADAPTER_OUT/DesktopAgent.Adapter.Mac"
    chmod +x "$ADAPTER_OUT/DesktopAgent.Adapter.Mac"
    ADAPTER_BINARY_OK=1
  fi
fi
popd >/dev/null

cp -R "$ROOT/adapters/macos/DesktopAgent.Adapter.Mac/"* "$ADAPTER_SRC_OUT/"
cp "$ROOT/core/DesktopAgent.Cli/appsettings.json" "$TRAY_OUT/agentsettings.json"
cp "$ROOT/core/DesktopAgent.Tray/appsettings.json" "$TRAY_OUT/appsettings.json"

cat > "$OUT/start-desktopagent.sh" <<'EOF'
#!/usr/bin/env bash
set -euo pipefail
ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
ADAPTER_PORT="${DESKTOP_AGENT_PORT:-51877}"

if [[ -x "$ROOT/adapter/DesktopAgent.Adapter.Mac" ]]; then
  nohup env DESKTOP_AGENT_PORT="$ADAPTER_PORT" "$ROOT/adapter/DesktopAgent.Adapter.Mac" > /tmp/desktopagent-adapter-macos.log 2>&1 &
else
  nohup env DESKTOP_AGENT_PORT="$ADAPTER_PORT" bash -lc "cd \"$ROOT/adapter-src\" && swift run" > /tmp/desktopagent-adapter-macos.log 2>&1 &
fi
echo "adapter:$!" > "$ROOT/.desktopagent.pids"

nohup env \
  DESKTOP_AGENT_TRAY_ADAPTERENDPOINT="http://localhost:${ADAPTER_PORT}" \
  DESKTOP_AGENT_TRAY_AGENTCONFIGPATH="$ROOT/tray/agentsettings.json" \
  DESKTOP_AGENT_TRAY_AUTOSTARTWEB="false" \
  "$ROOT/tray/DesktopAgent.Tray" > /tmp/desktopagent-tray-macos.log 2>&1 &
echo "tray:$!" >> "$ROOT/.desktopagent.pids"

echo "DesktopAgent started."
EOF

cat > "$OUT/stop-desktopagent.sh" <<'EOF'
#!/usr/bin/env bash
set -euo pipefail
ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PID_FILE="$ROOT/.desktopagent.pids"
if [[ -f "$PID_FILE" ]]; then
  while IFS=: read -r _ pid; do
    kill "$pid" 2>/dev/null || true
  done < "$PID_FILE"
  rm -f "$PID_FILE"
fi
pkill -f DesktopAgent.Tray 2>/dev/null || true
pkill -f DesktopAgent.Adapter.Mac 2>/dev/null || true
pkill -f "swift run" 2>/dev/null || true
echo "DesktopAgent stopped."
EOF

chmod +x "$OUT/start-desktopagent.sh" "$OUT/stop-desktopagent.sh"
if [[ "$ADAPTER_BINARY_OK" -eq 1 ]]; then
  echo "Publish completed (with native mac adapter binary): $OUT"
else
  echo "Publish completed (fallback adapter source package, run-time swift required): $OUT"
fi
