# macOS Setup

## Prereqs
- Xcode or Swift toolchain (Swift 5.9+)
- Homebrew (recommended)
- Accessibility permission for the adapter binary
- Screen recording permission for screenshots

## Build (Adapter)
The macOS adapter is a Swift package. Generate gRPC Swift code from `proto/desktop_adapter.proto` and then build.

```
brew install protobuf
brew install swift-protobuf
brew install grpc-swift

cd adapters/macos/DesktopAgent.Adapter.Mac
protoc -I ../../../proto \
  --swift_out=Sources/DesktopAgentAdapterMac \
  --grpc-swift_out=Sources/DesktopAgentAdapterMac \
  ../../../proto/desktop_adapter.proto

swift build
```

## Run Adapter
```
export DESKTOP_AGENT_PORT=51877
swift run
```

## Run Tray (recommended)
In a second terminal:
```
DESKTOP_AGENT_TRAY_ADAPTERENDPOINT=http://localhost:51877 \
DESKTOP_AGENT_TRAY_AGENTCONFIGPATH=core/DesktopAgent.Cli/appsettings.json \
DESKTOP_AGENT_TRAY_AUTOSTARTWEB=false \
dotnet run --project core/DesktopAgent.Tray/DesktopAgent.Tray.csproj
```

## Start Everything (macOS)
```
bash scripts/start-macos.sh
```

Stop:
```
bash scripts/stop-macos.sh
```

## Publish Portable (macOS)
Run on macOS host:
```
bash scripts/publish-macos.sh
```

Output:
- `dist/macos/adapter`
- `dist/macos/tray`
- `dist/macos/start-desktopagent.sh`
- `dist/macos/stop-desktopagent.sh`

## Permissions
- System Settings -> Privacy & Security -> Accessibility -> allow the adapter
- System Settings -> Privacy & Security -> Screen Recording -> allow the adapter

## Notes
- The adapter starts disarmed; arm it from tray/CLI before actions.
- Wayland-style restrictions do not apply on macOS, but permissions are enforced by the OS.
