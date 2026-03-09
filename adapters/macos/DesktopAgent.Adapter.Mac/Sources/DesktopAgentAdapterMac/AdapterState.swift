import Foundation
import AppKit
import ApplicationServices

struct WindowInfo {
    let title: String
    let appId: String
    let bounds: CGRect
}

final class AdapterState {
    private var elements: [String: AXUIElement] = [:]
    private var windows: [String: AXUIElement] = [:]
    private let lock = NSLock()

    private(set) var armed: Bool = false
    private(set) var requireUserPresence: Bool = false

    func arm(requireUserPresence: Bool) {
        lock.lock()
        defer { lock.unlock() }
        armed = true
        self.requireUserPresence = requireUserPresence
    }

    func disarm() {
        lock.lock()
        defer { lock.unlock() }
        armed = false
        requireUserPresence = false
    }

    func rememberWindow(_ element: AXUIElement) -> String {
        let id = UUID().uuidString.replacingOccurrences(of: "-", with: "")
        lock.lock()
        windows[id] = element
        elements[id] = element
        lock.unlock()
        return id
    }

    func rememberElement(_ element: AXUIElement) -> String {
        let id = UUID().uuidString.replacingOccurrences(of: "-", with: "")
        lock.lock()
        elements[id] = element
        lock.unlock()
        return id
    }

    func getWindow(_ id: String) -> AXUIElement? {
        lock.lock()
        defer { lock.unlock() }
        return windows[id]
    }

    func getElement(_ id: String) -> AXUIElement? {
        lock.lock()
        defer { lock.unlock() }
        return elements[id]
    }
}

enum AccessibilityHelper {
    static func isTrusted(prompt: Bool = false) -> Bool {
        let options: NSDictionary = [kAXTrustedCheckOptionPrompt.takeUnretainedValue() as String: prompt]
        return AXIsProcessTrustedWithOptions(options)
    }

    static func activeWindow() -> (AXUIElement, WindowInfo)? {
        guard let app = NSWorkspace.shared.frontmostApplication else {
            return nil
        }
        let appElement = AXUIElementCreateApplication(app.processIdentifier)
        guard let window = copyAttributeValue(appElement, kAXFocusedWindowAttribute) as? AXUIElement else {
            return nil
        }
        let info = windowInfo(window, appId: app.bundleIdentifier ?? app.localizedName ?? "")
        return (window, info)
    }

    static func listWindows() -> [(AXUIElement, WindowInfo)] {
        var results: [(AXUIElement, WindowInfo)] = []
        for app in NSWorkspace.shared.runningApplications where app.isFinishedLaunching && !app.isHidden {
            let appElement = AXUIElementCreateApplication(app.processIdentifier)
            guard let windows = copyAttributeValue(appElement, kAXWindowsAttribute) as? [AXUIElement] else {
                continue
            }
            for window in windows {
                let info = windowInfo(window, appId: app.bundleIdentifier ?? app.localizedName ?? "")
                results.append((window, info))
            }
        }
        return results
    }

    static func windowInfo(_ window: AXUIElement, appId: String) -> WindowInfo {
        let title = getString(window, kAXTitleAttribute) ?? ""
        let bounds = getBounds(window) ?? .zero
        return WindowInfo(title: title, appId: appId, bounds: bounds)
    }

    static func copyAttributeValue(_ element: AXUIElement, _ attribute: CFString) -> AnyObject? {
        var value: AnyObject?
        let result = AXUIElementCopyAttributeValue(element, attribute, &value)
        if result == .success {
            return value
        }
        return nil
    }

    static func getString(_ element: AXUIElement, _ attribute: CFString) -> String? {
        return copyAttributeValue(element, attribute) as? String
    }

    static func getBounds(_ element: AXUIElement) -> CGRect? {
        guard let positionValue = copyAttributeValue(element, kAXPositionAttribute) as? AXValue,
              let sizeValue = copyAttributeValue(element, kAXSizeAttribute) as? AXValue else {
            return nil
        }
        var point = CGPoint.zero
        var size = CGSize.zero
        AXValueGetValue(positionValue, .cgPoint, &point)
        AXValueGetValue(sizeValue, .cgSize, &size)
        return CGRect(origin: point, size: size)
    }
}

enum ClipboardHelper {
    static func getText() -> String {
        NSPasteboard.general.string(forType: .string) ?? ""
    }

    static func setText(_ text: String) {
        let pb = NSPasteboard.general
        pb.clearContents()
        pb.setString(text, forType: .string)
    }
}

enum ScreenshotHelper {
    static func screenBounds(index: Int) -> CGRect? {
        let screens = NSScreen.screens
            .sorted { lhs, rhs in
                if lhs.frame.minX == rhs.frame.minX {
                    return lhs.frame.minY < rhs.frame.minY
                }
                return lhs.frame.minX < rhs.frame.minX
            }

        guard index >= 0, index < screens.count else {
            return nil
        }

        return screens[index].frame
    }

    static func capture(region: CGRect?) -> (Data, CGSize)? {
        let rect = region ?? CGRect(x: 0, y: 0, width: NSScreen.main?.frame.width ?? 1024, height: NSScreen.main?.frame.height ?? 768)
        guard let image = CGWindowListCreateImage(rect, .optionOnScreenOnly, kCGNullWindowID, .bestResolution) else {
            return nil
        }
        let rep = NSBitmapImageRep(cgImage: image)
        guard let data = rep.representation(using: .png, properties: [:]) else {
            return nil
        }
        return (data, CGSize(width: image.width, height: image.height))
    }
}

enum InputHelper {
    static func click(point: CGPoint) {
        guard let source = CGEventSource(stateID: .hidSystemState) else { return }
        let mouseDown = CGEvent(mouseEventSource: source, mouseType: .leftMouseDown, mouseCursorPosition: point, mouseButton: .left)
        let mouseUp = CGEvent(mouseEventSource: source, mouseType: .leftMouseUp, mouseCursorPosition: point, mouseButton: .left)
        mouseDown?.post(tap: .cghidEventTap)
        mouseUp?.post(tap: .cghidEventTap)
    }

    static func typeText(_ text: String) {
        guard let source = CGEventSource(stateID: .hidSystemState) else { return }
        for scalar in text.unicodeScalars {
            var utf16 = Array(String(scalar).utf16)
            let keyDown = CGEvent(keyboardEventSource: source, virtualKey: 0, keyDown: true)
            keyDown?.keyboardSetUnicodeString(stringLength: utf16.count, unicodeString: &utf16)
            keyDown?.post(tap: .cghidEventTap)
            let keyUp = CGEvent(keyboardEventSource: source, virtualKey: 0, keyDown: false)
            keyUp?.keyboardSetUnicodeString(stringLength: utf16.count, unicodeString: &utf16)
            keyUp?.post(tap: .cghidEventTap)
        }
    }

    static func keyCombo(keys: [String]) {
        guard let source = CGEventSource(stateID: .hidSystemState) else { return }
        let keyCodes = keys.compactMap { mapKey($0) }
        for code in keyCodes {
            let keyDown = CGEvent(keyboardEventSource: source, virtualKey: code, keyDown: true)
            keyDown?.post(tap: .cghidEventTap)
        }
        for code in keyCodes.reversed() {
            let keyUp = CGEvent(keyboardEventSource: source, virtualKey: code, keyDown: false)
            keyUp?.post(tap: .cghidEventTap)
        }
    }

    static func mapKey(_ key: String) -> CGKeyCode? {
        let lower = key.lowercased()
        switch lower {
        case "cmd", "command": return 55
        case "ctrl", "control": return 59
        case "alt", "option": return 58
        case "shift": return 56
        case "enter", "return": return 36
        case "tab": return 48
        case "esc", "escape": return 53
        case "space": return 49
        case "up": return 126
        case "down": return 125
        case "left": return 123
        case "right": return 124
        default:
            return keyCodeForLetter(lower)
        }
    }

    private static func keyCodeForLetter(_ key: String) -> CGKeyCode? {
        let map: [String: CGKeyCode] = [
            "a": 0, "s": 1, "d": 2, "f": 3, "h": 4, "g": 5, "z": 6, "x": 7, "c": 8, "v": 9,
            "b": 11, "q": 12, "w": 13, "e": 14, "r": 15, "y": 16, "t": 17, "1": 18, "2": 19,
            "3": 20, "4": 21, "6": 22, "5": 23, "=": 24, "9": 25, "7": 26, "-": 27, "8": 28,
            "0": 29, "]": 30, "o": 31, "u": 32, "[": 33, "i": 34, "p": 35, "l": 37, "j": 38,
            "'": 39, "k": 40, ";": 41, "\\": 42, ",": 43, "/": 44, "n": 45, "m": 46, ".": 47,
            "`": 50
        ]
        return map[key]
    }
}
