#!/usr/bin/env bash
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
CONFIGURATION="${CONFIGURATION:-Release}"
RUNTIME="${RUNTIME:-linux-x64}"
OUTPUT_ROOT="${OUTPUT_ROOT:-dist/linux-x64}"

OUT="$ROOT/$OUTPUT_ROOT"
ADAPTER_OUT="$OUT/adapter"
TRAY_OUT="$OUT/tray"

rm -rf "$OUT"
mkdir -p "$ADAPTER_OUT" "$TRAY_OUT"

dotnet test "$ROOT/core/DesktopAgent.Tests/DesktopAgent.Tests.csproj" -c "$CONFIGURATION"
dotnet publish "$ROOT/core/DesktopAgent.Tray/DesktopAgent.Tray.csproj" -c "$CONFIGURATION" -r "$RUNTIME" --self-contained true -p:PublishSingleFile=true -o "$TRAY_OUT"

cp "$ROOT/adapters/linux/DesktopAgent.Adapter.Linux/server.py" "$ADAPTER_OUT/"
cp "$ROOT/adapters/linux/DesktopAgent.Adapter.Linux/desktop_adapter_pb2.py" "$ADAPTER_OUT/"
cp "$ROOT/adapters/linux/DesktopAgent.Adapter.Linux/desktop_adapter_pb2_grpc.py" "$ADAPTER_OUT/"
cp "$ROOT/adapters/linux/DesktopAgent.Adapter.Linux/requirements.txt" "$ADAPTER_OUT/"
cp "$ROOT/core/DesktopAgent.Cli/appsettings.json" "$TRAY_OUT/agentsettings.json"
cp "$ROOT/core/DesktopAgent.Tray/appsettings.json" "$TRAY_OUT/appsettings.json"

cat > "$OUT/start-desktopagent.sh" <<'EOF'
#!/usr/bin/env bash
set -euo pipefail
ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
ADAPTER_PORT="${DESKTOP_AGENT_PORT:-51877}"
PYTHON_BIN="${PYTHON_BIN:-python3}"

nohup "$PYTHON_BIN" "$ROOT/adapter/server.py" > /tmp/desktopagent-adapter-linux.log 2>&1 &
echo "adapter:$!" > "$ROOT/.desktopagent.pids"

nohup env \
  DESKTOP_AGENT_TRAY_ADAPTERENDPOINT="http://localhost:${ADAPTER_PORT}" \
  DESKTOP_AGENT_TRAY_AGENTCONFIGPATH="$ROOT/tray/agentsettings.json" \
  DESKTOP_AGENT_TRAY_AUTOSTARTWEB="false" \
  "$ROOT/tray/DesktopAgent.Tray" > /tmp/desktopagent-tray-linux.log 2>&1 &
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
pkill -f server.py 2>/dev/null || true
echo "DesktopAgent stopped."
EOF

chmod +x "$OUT/start-desktopagent.sh" "$OUT/stop-desktopagent.sh"
echo "Publish completed: $OUT"
