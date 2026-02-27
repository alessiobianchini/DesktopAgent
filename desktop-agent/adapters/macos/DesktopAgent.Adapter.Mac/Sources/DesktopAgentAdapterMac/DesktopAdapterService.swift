import Foundation
import GRPC
import NIO
import SwiftProtobuf
import AppKit
import ApplicationServices

final class DesktopAdapterService: Desktopagent_DesktopAdapterProvider {
    private let state: AdapterState
    private let maxDepth = 5

    init(state: AdapterState) {
        self.state = state
    }

    var interceptors: Desktopagent_DesktopAdapterServerInterceptorFactoryProtocol? { nil }

    func getActiveWindow(request: Desktopagent_Empty, context: StatusOnlyCallContext) -> EventLoopFuture<Desktopagent_WindowRef> {
        let promise = context.eventLoop.makePromise(of: Desktopagent_WindowRef.self)
        if let (window, info) = AccessibilityHelper.activeWindow() {
            let id = state.rememberWindow(window)
            promise.succeed(windowRef(id: id, info: info))
        } else {
            promise.succeed(Desktopagent_WindowRef())
        }
        return promise.futureResult
    }

    func listWindows(request: Desktopagent_Empty, context: StatusOnlyCallContext) -> EventLoopFuture<Desktopagent_WindowList> {
        let promise = context.eventLoop.makePromise(of: Desktopagent_WindowList.self)
        var list = Desktopagent_WindowList()
        for (window, info) in AccessibilityHelper.listWindows() {
            let id = state.rememberWindow(window)
            list.windows.append(windowRef(id: id, info: info))
        }
        promise.succeed(list)
        return promise.futureResult
    }

    func getUiTree(request: Desktopagent_WindowRequest, context: StatusOnlyCallContext) -> EventLoopFuture<Desktopagent_UiTree> {
        let promise = context.eventLoop.makePromise(of: Desktopagent_UiTree.self)
        guard let window = resolveWindow(id: request.windowID),
              let info = windowInfo(for: window) else {
            promise.succeed(Desktopagent_UiTree())
            return promise.futureResult
        }
        var tree = Desktopagent_UiTree()
        tree.window = windowRef(id: request.windowID, info: info)
        tree.root = buildNode(element: window, depth: 0)
        promise.succeed(tree)
        return promise.futureResult
    }

    func findElements(request: Desktopagent_FindRequest, context: StatusOnlyCallContext) -> EventLoopFuture<Desktopagent_ElementList> {
        let promise = context.eventLoop.makePromise(of: Desktopagent_ElementList.self)
        let selector = request.selector
        let scope = resolveWindow(id: selector.windowID) ?? AccessibilityHelper.activeWindow()?.0
        guard let root = scope else {
            promise.succeed(Desktopagent_ElementList())
            return promise.futureResult
        }

        var matches: [Desktopagent_ElementRef] = []
        traverse(element: root, depth: 0, ancestors: []) { element, ancestors in
            if matchesSelector(selector, element: element, ancestors: ancestors) {
                let id = state.rememberElement(element)
                matches.append(elementRef(id: id, element: element, ancestors: ancestors))
            }
        }

        if selector.index > 0 && selector.index < matches.count {
            matches = [matches[Int(selector.index)]]
        }

        var list = Desktopagent_ElementList()
        list.elements = matches
        promise.succeed(list)
        return promise.futureResult
    }

    func invokeElement(request: Desktopagent_ElementRequest, context: StatusOnlyCallContext) -> EventLoopFuture<Desktopagent_ActionResult> {
        return performIfArmed(context: context) {
            guard let element = self.state.getElement(request.elementID) else {
                return actionResult(false, "Element not found")
            }
            let result = AXUIElementPerformAction(element, kAXPressAction as CFString)
            return actionResult(result == .success, result == .success ? "Invoked" : "Invoke failed")
        }
    }

    func setElementValue(request: Desktopagent_SetValueRequest, context: StatusOnlyCallContext) -> EventLoopFuture<Desktopagent_ActionResult> {
        return performIfArmed(context: context) {
            guard let element = self.state.getElement(request.elementID) else {
                return actionResult(false, "Element not found")
            }
            let result = AXUIElementSetAttributeValue(element, kAXValueAttribute as CFString, request.value as CFTypeRef)
            return actionResult(result == .success, result == .success ? "Value set" : "Set value failed")
        }
    }

    func clickPoint(request: Desktopagent_ClickRequest, context: StatusOnlyCallContext) -> EventLoopFuture<Desktopagent_ActionResult> {
        return performIfArmed(context: context) {
            InputHelper.click(point: CGPoint(x: Int(request.x), y: Int(request.y)))
            return actionResult(true, "Clicked")
        }
    }

    func typeText(request: Desktopagent_TypeTextRequest, context: StatusOnlyCallContext) -> EventLoopFuture<Desktopagent_ActionResult> {
        return performIfArmed(context: context) {
            InputHelper.typeText(request.text)
            return actionResult(true, "Typed")
        }
    }

    func keyCombo(request: Desktopagent_KeyComboRequest, context: StatusOnlyCallContext) -> EventLoopFuture<Desktopagent_ActionResult> {
        return performIfArmed(context: context) {
            InputHelper.keyCombo(keys: request.keys)
            return actionResult(true, "Key combo sent")
        }
    }

    func openApp(request: Desktopagent_OpenAppRequest, context: StatusOnlyCallContext) -> EventLoopFuture<Desktopagent_ActionResult> {
        return performIfArmed(context: context) {
            let input = request.appIDOrPath.trimmingCharacters(in: .whitespacesAndNewlines)
            if input.isEmpty {
                return actionResult(false, "App id missing")
            }
            if FileManager.default.fileExists(atPath: input) {
                let url = URL(fileURLWithPath: input)
                let ok = NSWorkspace.shared.open(url)
                return actionResult(ok, ok ? "App launched" : "Open app failed")
            }

            let apps = self.installedApps()
            if let match = self.findBestMatch(query: input, apps: apps) {
                let url = URL(fileURLWithPath: match.path)
                let ok = NSWorkspace.shared.open(url)
                return actionResult(ok, ok ? "App launched: \(match.name)" : "Open app failed")
            }

            let ok = NSWorkspace.shared.launchApplication(input)
            if ok {
                return actionResult(true, "App launched")
            }

            let suggestions = self.suggestMatches(query: input, apps: apps, maxResults: 5)
            if !suggestions.isEmpty {
                return actionResult(false, "No confident match for '\(input)'. Top matches: \(suggestions.joined(separator: ", "))")
            }
            return actionResult(false, "Open app failed")
        }
    }

    func captureScreen(request: Desktopagent_ScreenshotRequest, context: StatusOnlyCallContext) -> EventLoopFuture<Desktopagent_ScreenshotResponse> {
        let promise = context.eventLoop.makePromise(of: Desktopagent_ScreenshotResponse.self)
        var region: CGRect? = nil
        if request.hasRegion {
            region = CGRect(x: Int(request.region.x), y: Int(request.region.y), width: Int(request.region.width), height: Int(request.region.height))
        }
        if let (data, size) = ScreenshotHelper.capture(region: region) {
            var response = Desktopagent_ScreenshotResponse()
            response.png = data
            response.width = Int32(size.width)
            response.height = Int32(size.height)
            promise.succeed(response)
        } else {
            promise.succeed(Desktopagent_ScreenshotResponse())
        }
        return promise.futureResult
    }

    func getClipboard(request: Desktopagent_Empty, context: StatusOnlyCallContext) -> EventLoopFuture<Desktopagent_ClipboardResponse> {
        let promise = context.eventLoop.makePromise(of: Desktopagent_ClipboardResponse.self)
        var response = Desktopagent_ClipboardResponse()
        response.text = ClipboardHelper.getText()
        promise.succeed(response)
        return promise.futureResult
    }

    func setClipboard(request: Desktopagent_SetClipboardRequest, context: StatusOnlyCallContext) -> EventLoopFuture<Desktopagent_ActionResult> {
        return performIfArmed(context: context) {
            ClipboardHelper.setText(request.text)
            return actionResult(true, "Clipboard set")
        }
    }

    func arm(request: Desktopagent_ArmRequest, context: StatusOnlyCallContext) -> EventLoopFuture<Desktopagent_Status> {
        let promise = context.eventLoop.makePromise(of: Desktopagent_Status.self)
        state.arm(requireUserPresence: request.requireUserPresence)
        var status = Desktopagent_Status()
        status.armed = true
        status.requireUserPresence = state.requireUserPresence
        status.message = "Armed"
        promise.succeed(status)
        return promise.futureResult
    }

    func disarm(request: Desktopagent_Empty, context: StatusOnlyCallContext) -> EventLoopFuture<Desktopagent_Status> {
        let promise = context.eventLoop.makePromise(of: Desktopagent_Status.self)
        state.disarm()
        var status = Desktopagent_Status()
        status.armed = false
        status.requireUserPresence = false
        status.message = "Disarmed"
        promise.succeed(status)
        return promise.futureResult
    }

    func getStatus(request: Desktopagent_Empty, context: StatusOnlyCallContext) -> EventLoopFuture<Desktopagent_Status> {
        let promise = context.eventLoop.makePromise(of: Desktopagent_Status.self)
        var status = Desktopagent_Status()
        status.armed = state.armed
        status.requireUserPresence = state.requireUserPresence
        status.message = state.armed ? "Armed" : "Disarmed"
        promise.succeed(status)
        return promise.futureResult
    }

    private func performIfArmed<T>(context: StatusOnlyCallContext, action: () -> T) -> EventLoopFuture<T> {
        let promise = context.eventLoop.makePromise(of: T.self)
        if state.armed {
            promise.succeed(action())
        } else {
            if T.self == Desktopagent_ActionResult.self {
                let result = actionResult(false, "Adapter is disarmed") as! T
                promise.succeed(result)
            } else {
                promise.fail(GRPCStatus(code: .failedPrecondition, message: "Adapter is disarmed"))
            }
        }
        return promise.futureResult
    }

    private func resolveWindow(id: String) -> AXUIElement? {
        guard !id.isEmpty else { return nil }
        return state.getWindow(id)
    }

    private func windowInfo(for window: AXUIElement) -> WindowInfo? {
        var pid: pid_t = 0
        AXUIElementGetPid(window, &pid)
        if let app = NSRunningApplication(processIdentifier: pid) {
            let appId = app.bundleIdentifier ?? app.localizedName ?? ""
            return AccessibilityHelper.windowInfo(window, appId: appId)
        }
        return AccessibilityHelper.windowInfo(window, appId: "")
    }

    private func buildNode(element: AXUIElement, depth: Int) -> Desktopagent_UiNode {
        var node = Desktopagent_UiNode()
        node.id = state.rememberElement(element)
        node.role = AccessibilityHelper.getString(element, kAXRoleAttribute) ?? ""
        node.name = AccessibilityHelper.getString(element, kAXTitleAttribute) ?? AccessibilityHelper.getString(element, kAXValueAttribute) ?? ""
        node.automationID = AccessibilityHelper.getString(element, kAXIdentifierAttribute) ?? ""
        node.className = ""
        if let bounds = AccessibilityHelper.getBounds(element) {
            node.bounds = rect(bounds)
        }

        if depth < maxDepth, let children = AccessibilityHelper.copyAttributeValue(element, kAXChildrenAttribute) as? [AXUIElement] {
            for child in children {
                node.children.append(buildNode(element: child, depth: depth + 1))
            }
        }

        return node
    }

    private func traverse(element: AXUIElement, depth: Int, ancestors: [String], visit: (AXUIElement, [String]) -> Void) {
        visit(element, ancestors)
        guard depth < maxDepth else { return }
        if let children = AccessibilityHelper.copyAttributeValue(element, kAXChildrenAttribute) as? [AXUIElement] {
            for child in children {
                let name = AccessibilityHelper.getString(child, kAXTitleAttribute) ?? ""
                traverse(element: child, depth: depth + 1, ancestors: ancestors + [name], visit: visit)
            }
        }
    }

    private func matchesSelector(_ selector: Desktopagent_Selector, element: AXUIElement, ancestors: [String]) -> Bool {
        if !selector.role.isEmpty {
            let role = AccessibilityHelper.getString(element, kAXRoleAttribute) ?? ""
            if !role.lowercased().contains(selector.role.lowercased()) {
                return false
            }
        }
        if !selector.nameContains.isEmpty {
            let name = AccessibilityHelper.getString(element, kAXTitleAttribute) ?? AccessibilityHelper.getString(element, kAXValueAttribute) ?? ""
            if !name.lowercased().contains(selector.nameContains.lowercased()) {
                return false
            }
        }
        if !selector.automationID.isEmpty {
            let ident = AccessibilityHelper.getString(element, kAXIdentifierAttribute) ?? ""
            if ident.lowercased() != selector.automationID.lowercased() {
                return false
            }
        }
        if !selector.ancestorNameContains.isEmpty {
            let match = ancestors.contains { $0.lowercased().contains(selector.ancestorNameContains.lowercased()) }
            if !match {
                return false
            }
        }
        if selector.hasBoundsHint {
            if let bounds = AccessibilityHelper.getBounds(element) {
                let hint = CGRect(x: Int(selector.boundsHint.x), y: Int(selector.boundsHint.y), width: Int(selector.boundsHint.width), height: Int(selector.boundsHint.height))
                if !bounds.intersects(hint) {
                    return false
                }
            }
        }
        return true
    }

    private func elementRef(id: String, element: AXUIElement, ancestors: [String]) -> Desktopagent_ElementRef {
        var ref = Desktopagent_ElementRef()
        ref.id = id
        ref.role = AccessibilityHelper.getString(element, kAXRoleAttribute) ?? ""
        ref.name = AccessibilityHelper.getString(element, kAXTitleAttribute) ?? AccessibilityHelper.getString(element, kAXValueAttribute) ?? ""
        ref.automationID = AccessibilityHelper.getString(element, kAXIdentifierAttribute) ?? ""
        ref.className = ""
        if let bounds = AccessibilityHelper.getBounds(element) {
            ref.bounds = rect(bounds)
        }
        ref.pathHints = ancestors.joined(separator: "/")
        return ref
    }

    private func windowRef(id: String, info: WindowInfo) -> Desktopagent_WindowRef {
        var window = Desktopagent_WindowRef()
        window.id = id
        window.title = info.title
        window.appID = info.appId
        window.bounds = rect(info.bounds)
        return window
    }

    private func rect(_ rect: CGRect) -> Desktopagent_Rect {
        var result = Desktopagent_Rect()
        result.x = Int32(rect.origin.x)
        result.y = Int32(rect.origin.y)
        result.width = Int32(rect.size.width)
        result.height = Int32(rect.size.height)
        return result
    }

    private struct AppEntry {
        let name: String
        let path: String
    }

    private static var appCache: [AppEntry] = []
    private static var appCacheTime = Date.distantPast

    private func installedApps() -> [AppEntry] {
        let now = Date()
        if now.timeIntervalSince(Self.appCacheTime) < 300, !Self.appCache.isEmpty {
            return Self.appCache
        }

        var results: [AppEntry] = []
        let home = FileManager.default.homeDirectoryForCurrentUser.path
        let roots = ["/Applications", "/System/Applications", "\(home)/Applications"]
        for root in roots where FileManager.default.fileExists(atPath: root) {
            results.append(contentsOf: enumerateApps(root: root, maxDepth: 2))
        }

        let deduped = deduplicateApps(results)
        Self.appCache = deduped
        Self.appCacheTime = now
        return deduped
    }

    private func enumerateApps(root: String, maxDepth: Int) -> [AppEntry] {
        var results: [AppEntry] = []
        var queue: [(path: String, depth: Int)] = [(root, 0)]
        while !queue.isEmpty {
            let current = queue.removeFirst()
            if current.depth > maxDepth {
                continue
            }
            guard let subdirs = try? FileManager.default.contentsOfDirectory(atPath: current.path) else {
                continue
            }
            for entry in subdirs {
                let fullPath = (current.path as NSString).appendingPathComponent(entry)
                var isDir: ObjCBool = false
                if FileManager.default.fileExists(atPath: fullPath, isDirectory: &isDir), isDir.boolValue {
                    if fullPath.lowercased().hasSuffix(".app") {
                        let name = (fullPath as NSString).lastPathComponent.replacingOccurrences(of: ".app", with: "")
                        if !name.isEmpty {
                            results.append(AppEntry(name: name, path: fullPath))
                        }
                    } else if current.depth + 1 <= maxDepth {
                        queue.append((fullPath, current.depth + 1))
                    }
                }
            }
        }
        return results
    }

    private func deduplicateApps(_ apps: [AppEntry]) -> [AppEntry] {
        var seen = Set<String>()
        var results: [AppEntry] = []
        for app in apps {
            let key = "\(app.name)|\(app.path)".lowercased()
            if !seen.contains(key) {
                seen.insert(key)
                results.append(app)
            }
        }
        return results
    }

    private func findBestMatch(query: String, apps: [AppEntry]) -> AppEntry? {
        let normalized = normalizeText(query)
        if normalized.isEmpty {
            return nil
        }
        var best: (entry: AppEntry?, score: Double) = (nil, 0.0)
        for app in apps {
            let score = scoreMatch(query: normalized, candidate: app.name)
            if score > best.score {
                best = (app, score)
            }
        }
        return best.score >= 0.72 ? best.entry : nil
    }

    private func suggestMatches(query: String, apps: [AppEntry], maxResults: Int) -> [String] {
        let normalized = normalizeText(query)
        if normalized.isEmpty {
            return []
        }
        var scored: [(name: String, score: Double)] = []
        for app in apps {
            let score = scoreMatch(query: normalized, candidate: app.name)
            if score >= 0.35 {
                scored.append((app.name, score))
            }
        }
        scored.sort { lhs, rhs in
            if lhs.score == rhs.score {
                return lhs.name.lowercased() < rhs.name.lowercased()
            }
            return lhs.score > rhs.score
        }
        var results: [String] = []
        for entry in scored {
            if !results.contains(entry.name) {
                results.append(entry.name)
            }
            if results.count >= maxResults {
                break
            }
        }
        return results
    }

    private func scoreMatch(query: String, candidate: String) -> Double {
        let normalizedCandidate = normalizeText(candidate)
        if normalizedCandidate.isEmpty {
            return 0.0
        }
        if normalizedCandidate == query {
            return 1.0
        }
        if normalizedCandidate.hasPrefix(query) {
            return 0.9
        }
        if normalizedCandidate.contains(query) {
            return 0.8
        }

        let queryTokens = tokenize(query)
        let candidateTokens = tokenize(normalizedCandidate)
        if queryTokens.isEmpty || candidateTokens.isEmpty {
            return 0.0
        }
        let intersection = queryTokens.intersection(candidateTokens).count
        let overlap = Double(intersection) / Double(max(queryTokens.count, candidateTokens.count))
        let allTokens = intersection == queryTokens.count
        var tokenScore = allTokens ? 0.85 : overlap

        let acronym = initialism(normalizedCandidate)
        if query.count <= 4 && acronym.hasPrefix(query) {
            tokenScore = max(tokenScore, 0.75)
        }
        return tokenScore
    }

    private func normalizeText(_ text: String) -> String {
        var value = text.trimmingCharacters(in: .whitespacesAndNewlines).lowercased()
        value = value.replacingOccurrences(of: "per favore", with: "")
        value = value.replacingOccurrences(of: "perfavore", with: "")
        value = value.replacingOccurrences(of: "please", with: "")
        value = value.replacingOccurrences(of: "application", with: "")
        value = value.replacingOccurrences(of: "applicazione", with: "")
        value = value.replacingOccurrences(of: "programma", with: "")
        value = value.replacingOccurrences(of: "program", with: "")
        value = value.replacingOccurrences(of: "app", with: "")
        value = value.replacingOccurrences(of: "[^a-z0-9 ]", with: " ", options: .regularExpression)
        value = value.replacingOccurrences(of: "\\s+", with: " ", options: .regularExpression)
        return value.trimmingCharacters(in: .whitespacesAndNewlines)
    }

    private func tokenize(_ text: String) -> Set<String> {
        let parts = text.split(separator: " ").map { String($0) }
        return Set(parts)
    }

    private func initialism(_ text: String) -> String {
        let parts = text.split(separator: " ").filter { !$0.isEmpty }
        return parts.map { String($0.prefix(1)) }.joined()
    }

    private func actionResult(_ success: Bool, _ message: String) -> Desktopagent_ActionResult {
        var result = Desktopagent_ActionResult()
        result.success = success
        result.message = message
        return result
    }
}
