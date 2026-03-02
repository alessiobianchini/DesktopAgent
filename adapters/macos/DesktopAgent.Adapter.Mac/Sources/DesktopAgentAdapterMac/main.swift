import Foundation
import GRPC
import NIO
import AppKit

let portEnv = ProcessInfo.processInfo.environment["DESKTOP_AGENT_PORT"] ?? "51877"
let port = Int(portEnv) ?? 51877

if !AccessibilityHelper.isTrusted(prompt: true) {
    print("Accessibility permission not granted. The adapter will start but UI automation may fail.")
}

let group = MultiThreadedEventLoopGroup(numberOfThreads: System.coreCount)
let state = AdapterState()
let service = DesktopAdapterService(state: state)

let server = try Server.insecure(group: group)
    .withServiceProviders([service])
    .bind(host: "0.0.0.0", port: port)
    .wait()

print("DesktopAgent.Adapter.Mac listening on \(server.channel.localAddress!)")

try server.onClose.wait()
try group.syncShutdownGracefully()
