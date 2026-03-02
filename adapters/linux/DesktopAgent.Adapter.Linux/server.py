import os
import subprocess
import uuid
import shlex
import time
import re
from concurrent import futures
import grpc

import desktop_adapter_pb2 as pb2
import desktop_adapter_pb2_grpc as pb2_grpc

try:
    import pyatspi  # type: ignore
    HAS_ATSPI = True
except Exception:
    HAS_ATSPI = False


class AdapterState:
    def __init__(self):
        self.armed = False
        self.require_user_presence = False
        self.elements = {}
        self.windows = {}

    def arm(self, require_user_presence: bool):
        self.armed = True
        self.require_user_presence = require_user_presence

    def disarm(self):
        self.armed = False
        self.require_user_presence = False

    def remember_window(self, element):
        element_id = uuid.uuid4().hex
        self.windows[element_id] = element
        self.elements[element_id] = element
        return element_id

    def remember_element(self, element):
        element_id = uuid.uuid4().hex
        self.elements[element_id] = element
        return element_id

    def get_window(self, element_id):
        return self.windows.get(element_id)

    def get_element(self, element_id):
        return self.elements.get(element_id)


def not_armed_result():
    return pb2.ActionResult(success=False, message="Adapter is disarmed")


def not_supported(message="NotSupported"):
    return pb2.ActionResult(success=False, message=message)


class DesktopAdapterService(pb2_grpc.DesktopAdapterServicer):
    def __init__(self, state: AdapterState):
        self.state = state

    def Arm(self, request, context):
        self.state.arm(request.requireUserPresence)
        return pb2.Status(armed=True, requireUserPresence=self.state.require_user_presence, message="Armed")

    def Disarm(self, request, context):
        self.state.disarm()
        return pb2.Status(armed=False, requireUserPresence=False, message="Disarmed")

    def GetStatus(self, request, context):
        return pb2.Status(armed=self.state.armed, requireUserPresence=self.state.require_user_presence,
                          message="Armed" if self.state.armed else "Disarmed")

    def GetActiveWindow(self, request, context):
        window = get_active_window()
        if window is None:
            return pb2.WindowRef()
        window_id = self.state.remember_window(window)
        return window_ref(window_id, window)

    def ListWindows(self, request, context):
        windows = []
        for window in list_windows():
            window_id = self.state.remember_window(window)
            windows.append(window_ref(window_id, window))
        return pb2.WindowList(windows=windows)

    def GetUiTree(self, request, context):
        window = self.state.get_window(request.windowId)
        if window is None:
            return pb2.UiTree()
        tree = pb2.UiTree()
        tree.window.CopyFrom(window_ref(request.windowId, window))
        tree.root.CopyFrom(build_ui_node(self.state, window, depth=0, max_depth=5))
        return tree

    def FindElements(self, request, context):
        selector = request.selector
        root = self.state.get_window(selector.windowId) if selector.windowId else None
        if root is None:
            root = get_desktop_root()
        if root is None:
            return pb2.ElementList(elements=[])

        matches = []
        for element, ancestors in traverse(root, depth=0, max_depth=6):
            if matches_selector(selector, element, ancestors):
                element_id = self.state.remember_element(element)
                matches.append(element_ref(element_id, element, ancestors))

        if selector.index > 0 and selector.index < len(matches):
            matches = [matches[selector.index]]

        return pb2.ElementList(elements=matches)

    def InvokeElement(self, request, context):
        if not self.state.armed:
            return not_armed_result()
        element = self.state.get_element(request.elementId)
        if element is None:
            return pb2.ActionResult(success=False, message="Element not found")
        return invoke_element(element)

    def SetElementValue(self, request, context):
        if not self.state.armed:
            return not_armed_result()
        element = self.state.get_element(request.elementId)
        if element is None:
            return pb2.ActionResult(success=False, message="Element not found")
        return set_element_value(element, request.value)

    def ClickPoint(self, request, context):
        if not self.state.armed:
            return not_armed_result()
        if not shutil_which("xdotool"):
            return not_supported("xdotool not installed")
        try:
            subprocess.run(["xdotool", "mousemove", str(request.x), str(request.y), "click", "1"], check=True)
            return pb2.ActionResult(success=True, message="Clicked")
        except Exception as exc:
            return pb2.ActionResult(success=False, message=str(exc))

    def TypeText(self, request, context):
        if not self.state.armed:
            return not_armed_result()
        if not shutil_which("xdotool"):
            return not_supported("xdotool not installed")
        try:
            subprocess.run(["xdotool", "type", "--clearmodifiers", request.text], check=True)
            return pb2.ActionResult(success=True, message="Typed")
        except Exception as exc:
            return pb2.ActionResult(success=False, message=str(exc))

    def KeyCombo(self, request, context):
        if not self.state.armed:
            return not_armed_result()
        if not shutil_which("xdotool"):
            return not_supported("xdotool not installed")
        combo = "+".join([key.lower() for key in request.keys])
        try:
            subprocess.run(["xdotool", "key", "--clearmodifiers", combo], check=True)
            return pb2.ActionResult(success=True, message="Key combo sent")
        except Exception as exc:
            return pb2.ActionResult(success=False, message=str(exc))

    def OpenApp(self, request, context):
        if not self.state.armed:
            return not_armed_result()
        app = (request.appIdOrPath or "").strip()
        try:
            if not app:
                return pb2.ActionResult(success=False, message="App id missing")
            if os.path.exists(app):
                subprocess.Popen([app])
                return pb2.ActionResult(success=True, message="App launched")
            if shutil_which("gtk-launch") and looks_like_desktop_id(app):
                subprocess.Popen(["gtk-launch", app])
                return pb2.ActionResult(success=True, message="App launched")
            match = find_best_app_match(app)
            if match:
                if match.get("desktop_id") and shutil_which("gtk-launch"):
                    subprocess.Popen(["gtk-launch", match["desktop_id"]])
                    return pb2.ActionResult(success=True, message=f"App launched: {match['name']}")
                if match.get("exec"):
                    subprocess.Popen(match["exec"])
                    return pb2.ActionResult(success=True, message=f"App launched: {match['name']}")
            if shutil_which("xdg-open"):
                subprocess.Popen(["xdg-open", app])
                return pb2.ActionResult(success=True, message="App launched")
            suggestions = suggest_app_matches(app, 5)
            if suggestions:
                return pb2.ActionResult(success=False, message=f"No confident match for '{app}'. Top matches: {', '.join(suggestions)}")
            return not_supported("No launcher available")
        except Exception as exc:
            return pb2.ActionResult(success=False, message=str(exc))

    def CaptureScreen(self, request, context):
        region = request.region if request.HasField("region") else None
        png_bytes = try_capture_x11(region)
        if png_bytes:
            return pb2.ScreenshotResponse(png=png_bytes)
        return pb2.ScreenshotResponse()

    def GetClipboard(self, request, context):
        return pb2.ClipboardResponse(text=get_clipboard())

    def SetClipboard(self, request, context):
        if not self.state.armed:
            return not_armed_result()
        ok, message = set_clipboard(request.text)
        return pb2.ActionResult(success=ok, message=message)


def get_desktop_root():
    if not HAS_ATSPI:
        return None
    try:
        return pyatspi.Registry.getDesktop(0)
    except Exception:
        return None


def list_windows():
    if not HAS_ATSPI:
        return []
    desktop = get_desktop_root()
    if desktop is None:
        return []
    results = []
    try:
        for i in range(desktop.childCount):
            app = desktop.getChildAtIndex(i)
            for j in range(app.childCount):
                window = app.getChildAtIndex(j)
                role = get_role_name(window).lower()
                if role in {"frame", "dialog", "window"}:
                    results.append(window)
    except Exception:
        return results
    return results


def get_active_window():
    if not HAS_ATSPI:
        return None
    for window in list_windows():
        try:
            state = window.getState()
            if state.contains(pyatspi.STATE_ACTIVE) or state.contains(pyatspi.STATE_FOCUSED):
                return window
        except Exception:
            continue
    return None


def build_ui_node(state: AdapterState, element, depth: int, max_depth: int):
    node = pb2.UiNode()
    node.id = state.remember_element(element)
    node.role = get_role_name(element)
    node.name = get_name(element)
    node.automationId = ""
    node.className = ""
    bounds = get_bounds(element)
    if bounds:
        node.bounds.CopyFrom(rect_from(bounds))
    if depth < max_depth:
        for child in safe_children(element):
            node.children.append(build_ui_node(state, child, depth + 1, max_depth))
    return node


def traverse(root, depth: int, max_depth: int):
    queue = [(root, depth, [])]
    while queue:
        element, current_depth, ancestors = queue.pop(0)
        yield element, ancestors
        if current_depth >= max_depth:
            continue
        for child in safe_children(element):
            name = get_name(child)
            queue.append((child, current_depth + 1, ancestors + [name]))


def matches_selector(selector, element, ancestors):
    if selector.role:
        role = get_role_name(element)
        if selector.role.lower() not in role.lower():
            return False
    if selector.nameContains:
        if selector.nameContains.lower() not in get_name(element).lower():
            return False
    if selector.automationId:
        return False
    if selector.className:
        return False
    if selector.ancestorNameContains:
        if not any(selector.ancestorNameContains.lower() in name.lower() for name in ancestors):
            return False
    if selector.HasField("boundsHint"):
        bounds = get_bounds(element)
        if bounds:
            hint = (selector.boundsHint.x, selector.boundsHint.y, selector.boundsHint.width, selector.boundsHint.height)
            if not intersects(bounds, hint):
                return False
    return True


def invoke_element(element):
    try:
        action = element.queryAction()
        if action.nActions > 0:
            action.doAction(0)
            return pb2.ActionResult(success=True, message="Invoked")
    except Exception:
        pass
    return not_supported("Invoke not supported")


def set_element_value(element, value: str):
    try:
        editable = element.queryEditableText()
        editable.setTextContents(0, editable.characterCount, value)
        return pb2.ActionResult(success=True, message="Value set")
    except Exception:
        return not_supported("Set value not supported")


def element_ref(element_id: str, element, ancestors):
    ref = pb2.ElementRef()
    ref.id = element_id
    ref.role = get_role_name(element)
    ref.name = get_name(element)
    ref.automationId = ""
    ref.className = ""
    ref.pathHints = "/".join(ancestors)
    bounds = get_bounds(element)
    if bounds:
        ref.bounds.CopyFrom(rect_from(bounds))
    return ref


def window_ref(window_id: str, window):
    ref = pb2.WindowRef()
    ref.id = window_id
    ref.title = get_name(window)
    ref.appId = get_app_id(window)
    bounds = get_bounds(window)
    if bounds:
        ref.bounds.CopyFrom(rect_from(bounds))
    return ref


def get_role_name(element):
    try:
        return element.getRoleName()
    except Exception:
        try:
            return pyatspi.role_name(element.role)
        except Exception:
            return ""


def get_name(element):
    try:
        return element.name or ""
    except Exception:
        return ""


def get_app_id(element):
    try:
        app = element.getApplication()
        return app.name or ""
    except Exception:
        return ""


def safe_children(element):
    try:
        return [element.getChildAtIndex(i) for i in range(element.childCount)]
    except Exception:
        return []


def get_bounds(element):
    try:
        component = element.queryComponent()
        extents = component.getExtents(pyatspi.DESKTOP_COORDS)
        return extents.x, extents.y, extents.width, extents.height
    except Exception:
        return None


def rect_from(bounds):
    x, y, w, h = bounds
    return pb2.Rect(x=int(x), y=int(y), width=int(w), height=int(h))


def intersects(bounds, hint):
    x, y, w, h = bounds
    hx, hy, hw, hh = hint
    return not (x + w < hx or hx + hw < x or y + h < hy or hy + hh < y)


def try_capture_x11(region=None):
    if os.environ.get("DISPLAY") is None:
        return b""

    if not shutil_which("import"):
        return b""

    try:
        cmd = ["import", "-window", "root"]
        if region is not None:
            cmd.extend(["-crop", f"{region.width}x{region.height}+{region.x}+{region.y}"])
        cmd.append("png:-")
        proc = subprocess.run(cmd, capture_output=True, check=True)
        return proc.stdout
    except Exception:
        return b""


def get_clipboard():
    if shutil_which("wl-paste"):
        try:
            proc = subprocess.run(["wl-paste"], capture_output=True, check=True)
            return proc.stdout.decode("utf-8", errors="ignore")
        except Exception:
            return ""
    if shutil_which("xclip"):
        try:
            proc = subprocess.run(["xclip", "-selection", "clipboard", "-o"], capture_output=True, check=True)
            return proc.stdout.decode("utf-8", errors="ignore")
        except Exception:
            return ""
    return ""


def set_clipboard(text: str):
    if shutil_which("wl-copy"):
        try:
            subprocess.run(["wl-copy"], input=text.encode("utf-8"), check=True)
            return True, "Clipboard set"
        except Exception as exc:
            return False, str(exc)
    if shutil_which("xclip"):
        try:
            subprocess.run(["xclip", "-selection", "clipboard"], input=text.encode("utf-8"), check=True)
            return True, "Clipboard set"
        except Exception as exc:
            return False, str(exc)
    return False, "No clipboard tool available"


APP_CACHE = {"time": 0.0, "apps": []}


def list_installed_apps():
    now = time.time()
    if now - APP_CACHE["time"] < 300 and APP_CACHE["apps"]:
        return APP_CACHE["apps"]

    apps = []
    home = os.path.expanduser("~")
    dirs = [
        "/usr/share/applications",
        "/usr/local/share/applications",
        os.path.join(home, ".local", "share", "applications"),
    ]
    for directory in dirs:
        if not os.path.isdir(directory):
            continue
        for filename in os.listdir(directory):
            if not filename.endswith(".desktop"):
                continue
            path = os.path.join(directory, filename)
            entry = parse_desktop_file(path)
            if entry:
                apps.append(entry)

    APP_CACHE["apps"] = apps
    APP_CACHE["time"] = now
    return apps


def parse_desktop_file(path: str):
    try:
        in_entry = False
        name = None
        exec_cmd = None
        no_display = False
        with open(path, "r", encoding="utf-8", errors="ignore") as handle:
            for raw in handle:
                line = raw.strip()
                if line.startswith("[") and line.endswith("]"):
                    in_entry = line.lower() == "[desktop entry]"
                    continue
                if not in_entry or line.startswith("#"):
                    continue
                if line.lower().startswith("nodisplay="):
                    no_display = line.lower().endswith("true")
                if line.lower().startswith("hidden="):
                    no_display = no_display or line.lower().endswith("true")
                if line.lower().startswith("name=") and name is None:
                    name = line.split("=", 1)[1].strip()
                if line.lower().startswith("exec=") and exec_cmd is None:
                    exec_cmd = line.split("=", 1)[1].strip()

        if no_display or not name or not exec_cmd:
            return None

        exec_list = clean_exec(exec_cmd)
        if not exec_list:
            return None

        desktop_id = os.path.splitext(os.path.basename(path))[0]
        return {"name": name, "exec": exec_list, "desktop_id": desktop_id}
    except Exception:
        return None


def clean_exec(exec_cmd: str):
    try:
        tokens = shlex.split(exec_cmd)
    except Exception:
        tokens = exec_cmd.split()
    tokens = [t for t in tokens if not t.startswith("%")]
    if not tokens:
        return None
    if tokens[0].lower() == "env":
        idx = 1
        while idx < len(tokens) and "=" in tokens[idx]:
            idx += 1
        tokens = tokens[idx:]
    return tokens if tokens else None


def looks_like_desktop_id(value: str):
    return "." in value and not os.path.sep in value


def find_best_app_match(query: str):
    normalized = normalize_text(query)
    if not normalized:
        return None
    apps = list_installed_apps()
    best = (None, 0.0)
    for app in apps:
        score = score_match(normalized, app["name"])
        if score > best[1]:
            best = (app, score)
    return best[0] if best[1] >= 0.72 else None


def suggest_app_matches(query: str, max_count: int):
    normalized = normalize_text(query)
    if not normalized:
        return []
    apps = list_installed_apps()
    scored = []
    for app in apps:
        score = score_match(normalized, app["name"])
        if score >= 0.35:
            scored.append((app["name"], score))
    scored.sort(key=lambda x: (-x[1], x[0].lower()))
    results = []
    for name, _ in scored:
        if name not in results:
            results.append(name)
        if len(results) >= max_count:
            break
    return results


def score_match(normalized_query: str, candidate: str):
    normalized_candidate = normalize_text(candidate)
    if not normalized_candidate:
        return 0.0
    if normalized_candidate == normalized_query:
        return 1.0
    if normalized_candidate.startswith(normalized_query):
        return 0.9
    if normalized_query in normalized_candidate:
        return 0.8
    query_tokens = tokenize(normalized_query)
    cand_tokens = tokenize(normalized_candidate)
    if not query_tokens or not cand_tokens:
        return 0.0
    intersection = len(query_tokens.intersection(cand_tokens))
    overlap = intersection / float(max(len(query_tokens), len(cand_tokens)))
    all_tokens = intersection == len(query_tokens)
    token_score = 0.85 if all_tokens else overlap
    acronym = initialism(normalized_candidate)
    if len(normalized_query) <= 4 and acronym.startswith(normalized_query):
        token_score = max(token_score, 0.75)
    return token_score


def normalize_text(text: str):
    if not text:
        return ""
    value = text.strip().strip('"').strip("'").lower()
    value = value.replace("per favore", "").replace("perfavore", "").replace("please", "")
    value = value.replace("application", "").replace("applicazione", "")
    value = value.replace("programma", "").replace("program", "").replace("app", "")
    value = re.sub(r"[^a-z0-9 ]", " ", value)
    value = re.sub(r"\\s+", " ", value)
    return value.strip()


def tokenize(text: str):
    return set([t for t in text.split(" ") if t])


def initialism(text: str):
    parts = [p for p in text.split(" ") if p]
    return "".join([p[0] for p in parts])


def shutil_which(cmd):
    for path in os.environ.get("PATH", "").split(os.pathsep):
        candidate = os.path.join(path, cmd)
        if os.path.isfile(candidate) and os.access(candidate, os.X_OK):
            return candidate
    return None


def serve():
    port = os.environ.get("DESKTOP_AGENT_PORT", "51877")
    server = grpc.server(futures.ThreadPoolExecutor(max_workers=4))
    state = AdapterState()
    pb2_grpc.add_DesktopAdapterServicer_to_server(DesktopAdapterService(state), server)
    server.add_insecure_port(f"0.0.0.0:{port}")
    server.start()
    print(f"DesktopAgent.Adapter.Linux listening on {port}")
    server.wait_for_termination()


if __name__ == "__main__":
    serve()
