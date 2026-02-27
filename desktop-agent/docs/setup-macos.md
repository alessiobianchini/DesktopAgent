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

cd desktop-agent/adapters/macos/DesktopAgent.Adapter.Mac
protoc -I ../../../proto \
  --swift_out=Sources/DesktopAgentAdapterMac \
  --grpc-swift_out=Sources/DesktopAgentAdapterMac \
  ../../../proto/desktop_adapter.proto

swift build
```

## Run Adapter
```
export DESKTOP_AGENT_PORT=50051
swift run
```

## Permissions
- System Settings -> Privacy & Security -> Accessibility -> allow the adapter
- System Settings -> Privacy & Security -> Screen Recording -> allow the adapter

## Notes
- The adapter must be explicitly armed by the CLI before executing actions.
- Wayland-style restrictions do not apply on macOS, but permissions are enforced by the OS.
