using System.Drawing;
using System.Drawing.Imaging;
using DesktopAgent.Proto;
using FlaUI.Core.AutomationElements;
using FlaUI.Core.Definitions;
using FlaUI.Core.Input;
using FlaUI.Core.WindowsAPI;
using FlaUI.UIA3;
using Grpc.Core;
using TextCopy;
using Status = DesktopAgent.Proto.Status;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Diagnostics;
using System.Windows.Forms;

namespace DesktopAgent.Adapter.Windows;

public sealed class DesktopAdapterService : DesktopAdapter.DesktopAdapterBase
{
    private readonly AdapterState _state;

    public DesktopAdapterService(AdapterState state)
    {
        _state = state;
    }

    public override Task<Status> Arm(ArmRequest request, ServerCallContext context)
    {
        _state.Arm(request.RequireUserPresence);
        return Task.FromResult(new Status { Armed = true, RequireUserPresence = _state.RequireUserPresence, Message = "Armed" });
    }

    public override Task<Status> Disarm(Empty request, ServerCallContext context)
    {
        _state.Disarm();
        return Task.FromResult(new Status { Armed = false, RequireUserPresence = false, Message = "Disarmed" });
    }

    public override Task<Status> GetStatus(Empty request, ServerCallContext context)
    {
        return Task.FromResult(new Status { Armed = _state.Armed, RequireUserPresence = _state.RequireUserPresence, Message = _state.Armed ? "Armed" : "Disarmed" });
    }

    public override Task<WindowRef> GetActiveWindow(Empty request, ServerCallContext context)
    {
        var handle = NativeMethods.GetForegroundWindow();
        if (handle == IntPtr.Zero)
        {
            return Task.FromResult(new WindowRef());
        }

        var element = _state.Automation.FromHandle(handle);
        var id = _state.RememberWindow(element);
        return Task.FromResult(ToWindowRef(id, element));
    }

    public override Task<WindowList> ListWindows(Empty request, ServerCallContext context)
    {
        var desktop = _state.Automation.GetDesktop();
        var windows = desktop.FindAllChildren(cf => cf.ByControlType(ControlType.Window));
        var list = new WindowList();
        foreach (var window in windows)
        {
            var id = _state.RememberWindow(window);
            list.Windows.Add(ToWindowRef(id, window));
        }
        return Task.FromResult(list);
    }

    public override Task<UiTree> GetUiTree(WindowRequest request, ServerCallContext context)
    {
        var window = ResolveWindow(request.WindowId);
        if (window == null)
        {
            return Task.FromResult(new UiTree());
        }

        var rootNode = BuildUiNode(window, depthLimit: 5);
        return Task.FromResult(new UiTree { Window = ToWindowRef(request.WindowId, window), Root = rootNode });
    }

    public override Task<ElementList> FindElements(FindRequest request, ServerCallContext context)
    {
        var selector = request.Selector ?? new Selector();
        var scope = ResolveWindow(selector.WindowId) ?? _state.Automation.GetDesktop();
        var elements = new List<ElementRef>();
        var requestedIndex = selector.Index;
        var stopAtCount = requestedIndex > 0 ? requestedIndex + 1 : int.MaxValue;

        foreach (var element in Enumerate(scope, 6))
        {
            if (Matches(selector, element))
            {
                var id = _state.RememberElement(element);
                elements.Add(ToElementRef(id, element));
                if (elements.Count >= stopAtCount)
                {
                    break;
                }
            }
        }

        if (selector.Index > 0 && selector.Index < elements.Count)
        {
            elements = new List<ElementRef> { elements[selector.Index] };
        }

        var list = new ElementList();
        list.Elements.AddRange(elements);
        return Task.FromResult(list);
    }

    public override Task<ActionResult> InvokeElement(ElementRequest request, ServerCallContext context)
    {
        if (!EnsureArmed(out var result))
        {
            return Task.FromResult(result);
        }

        var element = _state.GetElement(request.ElementId);
        if (element == null)
        {
            return Task.FromResult(new ActionResult { Success = false, Message = "Element not found" });
        }

        try
        {
            if (element.Patterns.Invoke.IsSupported)
            {
                element.Patterns.Invoke.Pattern.Invoke();
                return Task.FromResult(new ActionResult { Success = true, Message = "Invoked" });
            }

            return Task.FromResult(new ActionResult { Success = false, Message = "Invoke not supported" });
        }
        catch (Exception ex)
        {
            return Task.FromResult(new ActionResult { Success = false, Message = ex.Message });
        }
    }

    public override Task<ActionResult> SetElementValue(SetValueRequest request, ServerCallContext context)
    {
        if (!EnsureArmed(out var result))
        {
            return Task.FromResult(result);
        }

        var element = _state.GetElement(request.ElementId);
        if (element == null)
        {
            return Task.FromResult(new ActionResult { Success = false, Message = "Element not found" });
        }

        try
        {
            if (element.Patterns.Value.IsSupported)
            {
                element.Patterns.Value.Pattern.SetValue(request.Value ?? string.Empty);
                return Task.FromResult(new ActionResult { Success = true, Message = "Value set" });
            }

            return Task.FromResult(new ActionResult { Success = false, Message = "Value pattern not supported" });
        }
        catch (Exception ex)
        {
            return Task.FromResult(new ActionResult { Success = false, Message = ex.Message });
        }
    }

    public override Task<ActionResult> ClickPoint(ClickRequest request, ServerCallContext context)
    {
        if (!EnsureArmed(out var result))
        {
            return Task.FromResult(result);
        }

        try
        {
            Mouse.Click(new System.Drawing.Point(request.X, request.Y));
            return Task.FromResult(new ActionResult { Success = true, Message = "Clicked" });
        }
        catch (Exception ex)
        {
            return Task.FromResult(new ActionResult { Success = false, Message = ex.Message });
        }
    }

    public override Task<ActionResult> TypeText(TypeTextRequest request, ServerCallContext context)
    {
        if (!EnsureArmed(out var result))
        {
            return Task.FromResult(result);
        }

        try
        {
            Keyboard.Type(request.Text ?? string.Empty);
            return Task.FromResult(new ActionResult { Success = true, Message = "Typed" });
        }
        catch (Exception ex)
        {
            return Task.FromResult(new ActionResult { Success = false, Message = ex.Message });
        }
    }

    public override Task<ActionResult> KeyCombo(KeyComboRequest request, ServerCallContext context)
    {
        if (!EnsureArmed(out var result))
        {
            return Task.FromResult(result);
        }

        try
        {
            foreach (var key in request.Keys)
            {
                if (TryMapKey(key, out var virtualKey))
                {
                    Keyboard.Press(virtualKey);
                }
            }

            foreach (var key in request.Keys.Reverse())
            {
                if (TryMapKey(key, out var virtualKey))
                {
                    Keyboard.Release(virtualKey);
                }
            }

            return Task.FromResult(new ActionResult { Success = true, Message = "Key combo sent" });
        }
        catch (Exception ex)
        {
            return Task.FromResult(new ActionResult { Success = false, Message = ex.Message });
        }
    }

    public override Task<ActionResult> OpenApp(OpenAppRequest request, ServerCallContext context)
    {
        if (!EnsureArmed(out var result))
        {
            return Task.FromResult(result);
        }

        try
        {
            var appIdOrPath = (request.AppIdOrPath ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(appIdOrPath))
            {
                return Task.FromResult(new ActionResult { Success = false, Message = "App id missing" });
            }

            if (LooksLikePath(appIdOrPath))
            {
                var info = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = appIdOrPath,
                    UseShellExecute = true
                };
                var process = System.Diagnostics.Process.Start(info);
                TryFocusProcess(process, TimeSpan.FromSeconds(4));
                return Task.FromResult(new ActionResult { Success = true, Message = "App launched" });
            }

            var candidates = GetInstalledApps();
            var match = FindBestMatch(appIdOrPath, candidates);
            if (match != null)
            {
                var info = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = match.Path,
                    UseShellExecute = true
                };
                var process = System.Diagnostics.Process.Start(info);
                TryFocusProcess(process, TimeSpan.FromSeconds(4));
                if (!TryFocusByHint(match.Name, TimeSpan.FromSeconds(4)))
                {
                    TryFocusByHint(appIdOrPath, TimeSpan.FromSeconds(2));
                }
                return Task.FromResult(new ActionResult { Success = true, Message = $"App launched: {match.Name}" });
            }

            var suggestions = SuggestMatches(appIdOrPath, candidates, 5);
            if (suggestions.Count > 0)
            {
                return Task.FromResult(new ActionResult
                {
                    Success = false,
                    Message = $"No confident match for '{appIdOrPath}'. Top matches: {string.Join(", ", suggestions)}"
                });
            }

            return Task.FromResult(new ActionResult { Success = false, Message = $"No app match for '{appIdOrPath}'" });
        }
        catch (Exception ex)
        {
            return Task.FromResult(new ActionResult { Success = false, Message = ex.Message });
        }
    }

    private sealed record AppEntry(string Name, string Path);

    private static readonly object AppCacheLock = new();
    private static DateTime AppCacheTimeUtc = DateTime.MinValue;
    private static List<AppEntry> AppCache = new();

    private static IReadOnlyList<AppEntry> GetInstalledApps()
    {
        lock (AppCacheLock)
        {
            if ((DateTime.UtcNow - AppCacheTimeUtc) < TimeSpan.FromMinutes(5) && AppCache.Count > 0)
            {
                return AppCache;
            }

            AppCache = EnumerateStartMenuApps();
            AppCacheTimeUtc = DateTime.UtcNow;
            return AppCache;
        }
    }

    private static List<AppEntry> EnumerateStartMenuApps()
    {
        var results = new List<AppEntry>();
        var folders = new[]
        {
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonStartMenu), "Programs"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.StartMenu), "Programs")
        };

        foreach (var folder in folders)
        {
            if (!Directory.Exists(folder))
            {
                continue;
            }

            foreach (var file in Directory.EnumerateFiles(folder, "*.*", SearchOption.AllDirectories))
            {
                var ext = Path.GetExtension(file);
                if (!ext.Equals(".lnk", StringComparison.OrdinalIgnoreCase)
                    && !ext.Equals(".exe", StringComparison.OrdinalIgnoreCase)
                    && !ext.Equals(".appref-ms", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var name = Path.GetFileNameWithoutExtension(file);
                if (string.IsNullOrWhiteSpace(name))
                {
                    continue;
                }

                results.Add(new AppEntry(name, file));
            }
        }

        return results;
    }

    private static AppEntry? FindBestMatch(string query, IReadOnlyList<AppEntry> apps)
    {
        var normalized = Normalize(query);
        if (string.IsNullOrWhiteSpace(normalized))
        {
            return null;
        }

        var best = (Entry: (AppEntry?)null, Score: 0.0);
        foreach (var app in apps)
        {
            var score = ScoreMatch(normalized, app.Name);
            if (score > best.Score)
            {
                best = (app, score);
            }
        }

        return best.Score >= 0.72 ? best.Entry : null;
    }

    private static List<string> SuggestMatches(string query, IReadOnlyList<AppEntry> apps, int max)
    {
        var normalized = Normalize(query);
        if (string.IsNullOrWhiteSpace(normalized))
        {
            return new List<string>();
        }

        return apps
            .Select(app => (app.Name, Score: ScoreMatch(normalized, app.Name)))
            .Where(x => x.Score >= 0.35)
            .OrderByDescending(x => x.Score)
            .Take(max)
            .Select(x => x.Name)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static double ScoreMatch(string normalizedQuery, string candidateName)
    {
        var normalizedCandidate = Normalize(candidateName);
        if (normalizedCandidate.Length == 0)
        {
            return 0.0;
        }

        if (normalizedCandidate.Equals(normalizedQuery, StringComparison.OrdinalIgnoreCase))
        {
            return 1.0;
        }

        if (normalizedCandidate.StartsWith(normalizedQuery, StringComparison.OrdinalIgnoreCase))
        {
            return 0.9;
        }

        if (normalizedCandidate.Contains(normalizedQuery, StringComparison.OrdinalIgnoreCase))
        {
            return 0.8;
        }

        var queryTokens = Tokenize(normalizedQuery);
        var candTokens = Tokenize(normalizedCandidate);
        if (queryTokens.Count == 0 || candTokens.Count == 0)
        {
            return 0.0;
        }

        var intersection = queryTokens.Intersect(candTokens, StringComparer.OrdinalIgnoreCase).Count();
        var overlap = intersection / (double)Math.Max(queryTokens.Count, candTokens.Count);
        var allTokensPresent = intersection == queryTokens.Count;
        var tokenScore = allTokensPresent ? 0.85 : overlap;

        var acronym = GetInitialism(normalizedCandidate);
        if (normalizedQuery.Length <= 4 && acronym.StartsWith(normalizedQuery, StringComparison.OrdinalIgnoreCase))
        {
            tokenScore = Math.Max(tokenScore, 0.75);
        }

        return tokenScore;
    }

    private static string Normalize(string input)
    {
        var value = input.Trim().Trim('"', '\'');
        value = Regex.Replace(value, "\\s+", " ");
        var lower = value.ToLowerInvariant();
        lower = lower.Replace("per favore", string.Empty)
                     .Replace("perfavore", string.Empty)
                     .Replace("please", string.Empty)
                     .Replace("application", string.Empty)
                     .Replace("applicazione", string.Empty)
                     .Replace("programma", string.Empty)
                     .Replace("program", string.Empty)
                     .Replace("app", string.Empty)
                     .Replace("il ", string.Empty)
                     .Replace("la ", string.Empty)
                     .Replace("lo ", string.Empty);
        lower = Regex.Replace(lower, "[^a-z0-9 ]", " ");
        return Regex.Replace(lower, "\\s+", " ").Trim();
    }

    private static HashSet<string> Tokenize(string input)
    {
        return input.Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
    }

    private static string GetInitialism(string input)
    {
        var parts = input.Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (parts.Length == 0)
        {
            return string.Empty;
        }

        var chars = parts.Select(p => p[0]);
        return new string(chars.ToArray());
    }

    private static bool LooksLikePath(string input)
    {
        if (input.Contains(Path.DirectorySeparatorChar)
            || input.Contains(Path.AltDirectorySeparatorChar)
            || input.Contains(":\\", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        return input.EndsWith(".exe", StringComparison.OrdinalIgnoreCase)
            || input.EndsWith(".lnk", StringComparison.OrdinalIgnoreCase)
            || input.EndsWith(".appref-ms", StringComparison.OrdinalIgnoreCase);
    }

    public override Task<ScreenshotResponse> CaptureScreen(ScreenshotRequest request, ServerCallContext context)
    {
        try
        {
            Rectangle bounds;
            if (request.Region != null && request.Region.Width > 0 && request.Region.Height > 0)
            {
                bounds = new Rectangle(request.Region.X, request.Region.Y, request.Region.Width, request.Region.Height);
            }
            else if (TryParseScreenIndexHint(request.WindowId, out var screenIndex))
            {
                var allScreens = Screen.AllScreens;
                if (screenIndex < 0 || screenIndex >= allScreens.Length)
                {
                    return Task.FromResult(new ScreenshotResponse());
                }

                bounds = allScreens[screenIndex].Bounds;
            }
            else if (!string.IsNullOrWhiteSpace(request.WindowId) && TryGetWindowBounds(request.WindowId, out var windowBounds))
            {
                bounds = windowBounds;
            }
            else
            {
                bounds = Screen.PrimaryScreen?.Bounds ?? new Rectangle(0, 0, 1024, 768);
            }

            using var bitmap = new Bitmap(bounds.Width, bounds.Height);
            using var graphics = Graphics.FromImage(bitmap);
            graphics.CopyFromScreen(bounds.Location, Point.Empty, bounds.Size);
            using var ms = new MemoryStream();
            bitmap.Save(ms, ImageFormat.Png);
            return Task.FromResult(new ScreenshotResponse { Png = Google.Protobuf.ByteString.CopyFrom(ms.ToArray()), Width = bounds.Width, Height = bounds.Height });
        }
        catch
        {
            return Task.FromResult(new ScreenshotResponse());
        }
    }

    private static bool TryParseScreenIndexHint(string? windowId, out int screenIndex)
    {
        screenIndex = -1;
        if (string.IsNullOrWhiteSpace(windowId))
        {
            return false;
        }

        const string prefix = "__screen:";
        if (!windowId.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        var raw = windowId[prefix.Length..].Trim();
        if (!int.TryParse(raw, out screenIndex))
        {
            screenIndex = -1;
        }

        return true;
    }

    private bool TryGetWindowBounds(string windowId, out Rectangle bounds)
    {
        bounds = default;
        var window = ResolveWindow(windowId);
        if (window == null)
        {
            return false;
        }

        var rect = SafeBoundingRectangle(window);
        if (rect.Width <= 0 || rect.Height <= 0)
        {
            return false;
        }

        bounds = new Rectangle(rect.Left, rect.Top, rect.Width, rect.Height);
        return true;
    }

    public override async Task<ClipboardResponse> GetClipboard(Empty request, ServerCallContext context)
    {
        try
        {
            var text = await ClipboardService.GetTextAsync();
            return new ClipboardResponse { Text = text ?? string.Empty };
        }
        catch
        {
            return new ClipboardResponse { Text = string.Empty };
        }
    }

    public override async Task<ActionResult> SetClipboard(SetClipboardRequest request, ServerCallContext context)
    {
        if (!EnsureArmed(out var result))
        {
            return result;
        }

        try
        {
            await ClipboardService.SetTextAsync(request.Text ?? string.Empty);
            return new ActionResult { Success = true, Message = "Clipboard set" };
        }
        catch (Exception ex)
        {
            return new ActionResult { Success = false, Message = ex.Message };
        }
    }

    private bool EnsureArmed(out ActionResult result)
    {
        if (_state.Armed)
        {
            result = new ActionResult { Success = true, Message = "Armed" };
            return true;
        }

        result = new ActionResult { Success = false, Message = "Adapter is disarmed" };
        return false;
    }

    private AutomationElement? ResolveWindow(string windowId)
    {
        if (string.IsNullOrWhiteSpace(windowId))
        {
            return null;
        }

        return _state.GetWindow(windowId);
    }

    private static WindowRef ToWindowRef(string id, AutomationElement element)
    {
        var rect = SafeBoundingRectangle(element);
        var processId = element.Properties.ProcessId.ValueOrDefault;
        var appId = string.Empty;
        if (processId > 0)
        {
            try
            {
                appId = System.Diagnostics.Process.GetProcessById(processId).ProcessName;
            }
            catch
            {
                appId = processId.ToString();
            }
        }

        return new WindowRef
        {
            Id = id,
            Title = SafeName(element),
            AppId = appId,
            Bounds = new Rect { X = rect.Left, Y = rect.Top, Width = rect.Width, Height = rect.Height }
        };
    }

    private static ElementRef ToElementRef(string id, AutomationElement element)
    {
        var rect = SafeBoundingRectangle(element);
        return new ElementRef
        {
            Id = id,
            Role = SafeControlType(element),
            Name = SafeName(element),
            AutomationId = SafeAutomationId(element),
            ClassName = SafeClassName(element),
            Bounds = new Rect { X = rect.Left, Y = rect.Top, Width = rect.Width, Height = rect.Height }
        };
    }

    private UiNode BuildUiNode(AutomationElement element, int depthLimit)
    {
        var rect = SafeBoundingRectangle(element);
        var node = new UiNode
        {
            Id = _state.RememberElement(element),
            Role = SafeControlType(element),
            Name = SafeName(element),
            AutomationId = SafeAutomationId(element),
            ClassName = SafeClassName(element),
            Bounds = new Rect { X = rect.Left, Y = rect.Top, Width = rect.Width, Height = rect.Height }
        };

        if (depthLimit <= 0)
        {
            return node;
        }

        foreach (var child in SafeChildren(element))
        {
            node.Children.Add(BuildUiNode(child, depthLimit - 1));
        }

        return node;
    }

    private static IEnumerable<AutomationElement> Enumerate(AutomationElement root, int depthLimit)
    {
        var queue = new Queue<(AutomationElement Element, int Depth)>();
        queue.Enqueue((root, 0));
        while (queue.Count > 0)
        {
            var (element, depth) = queue.Dequeue();
            yield return element;
            if (depth >= depthLimit) continue;
            foreach (var child in SafeChildren(element))
            {
                queue.Enqueue((child, depth + 1));
            }
        }
    }

    private static bool Matches(Selector selector, AutomationElement element)
    {
        if (!string.IsNullOrWhiteSpace(selector.Role))
        {
            var role = SafeControlType(element);
            if (!role.Contains(selector.Role, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }
        }

        if (!string.IsNullOrWhiteSpace(selector.NameContains))
        {
            var name = SafeName(element);
            if (!name.Contains(selector.NameContains, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }
        }

        if (!string.IsNullOrWhiteSpace(selector.AutomationId))
        {
            var automationId = SafeAutomationId(element);
            if (!string.Equals(automationId, selector.AutomationId, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }
        }

        if (!string.IsNullOrWhiteSpace(selector.ClassName))
        {
            var className = SafeClassName(element);
            if (!string.Equals(className, selector.ClassName, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }
        }

        if (selector.BoundsHint != null && selector.BoundsHint.Width > 0 && selector.BoundsHint.Height > 0)
        {
            var rect = SafeBoundingRectangle(element);
            if (!rect.IntersectsWith(new System.Drawing.Rectangle(selector.BoundsHint.X, selector.BoundsHint.Y, selector.BoundsHint.Width, selector.BoundsHint.Height)))
            {
                return false;
            }
        }

        return true;
    }

    private static IEnumerable<AutomationElement> SafeChildren(AutomationElement element)
    {
        try
        {
            return element.FindAllChildren();
        }
        catch
        {
            return Array.Empty<AutomationElement>();
        }
    }

    private static bool TryMapKey(string key, out VirtualKeyShort vk)
    {
        vk = key.ToLowerInvariant() switch
        {
            "ctrl" or "control" => VirtualKeyShort.CONTROL,
            "alt" => VirtualKeyShort.LMENU,
            "shift" => VirtualKeyShort.SHIFT,
            "win" or "windows" => VirtualKeyShort.LWIN,
            "enter" => VirtualKeyShort.RETURN,
            "tab" => VirtualKeyShort.TAB,
            "esc" or "escape" => VirtualKeyShort.ESCAPE,
            "backspace" => VirtualKeyShort.BACK,
            "delete" or "del" => VirtualKeyShort.DELETE,
            "home" => VirtualKeyShort.HOME,
            "end" => VirtualKeyShort.END,
            "pageup" or "pgup" => VirtualKeyShort.PRIOR,
            "pagedown" or "pgdn" => VirtualKeyShort.NEXT,
            "space" => VirtualKeyShort.SPACE,
            "up" => VirtualKeyShort.UP,
            "down" => VirtualKeyShort.DOWN,
            "left" => VirtualKeyShort.LEFT,
            "right" => VirtualKeyShort.RIGHT,
            _ => VirtualKeyShort.SPACE
        };

        if (vk != VirtualKeyShort.SPACE)
        {
            return true;
        }

        if (TryMapSingleCharKey(key, out var singleKey))
        {
            vk = singleKey;
            return true;
        }

        if (TryMapFunctionKey(key, out var functionKey))
        {
            vk = functionKey;
            return true;
        }

        return false;
    }

    private static bool TryMapSingleCharKey(string key, out VirtualKeyShort vk)
    {
        vk = default;
        if (string.IsNullOrWhiteSpace(key) || key.Length != 1)
        {
            return false;
        }

        var ch = char.ToUpperInvariant(key[0]);
        if (ch is >= 'A' and <= 'Z')
        {
            vk = (VirtualKeyShort)ch;
            return true;
        }

        if (ch is >= '0' and <= '9')
        {
            vk = (VirtualKeyShort)ch;
            return true;
        }

        return false;
    }

    private static bool TryMapFunctionKey(string key, out VirtualKeyShort vk)
    {
        vk = default;
        if (!Regex.IsMatch(key, "^f([1-9]|1[0-9]|2[0-4])$", RegexOptions.IgnoreCase))
        {
            return false;
        }

        if (!int.TryParse(key[1..], out var index))
        {
            return false;
        }

        vk = (VirtualKeyShort)(0x70 + index - 1);
        return true;
    }

    private static string SafeName(AutomationElement element)
    {
        try
        {
            return element.Properties.Name.ValueOrDefault ?? string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    private static string SafeAutomationId(AutomationElement element)
    {
        try
        {
            return element.Properties.AutomationId.ValueOrDefault ?? string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    private static string SafeClassName(AutomationElement element)
    {
        try
        {
            return element.Properties.ClassName.ValueOrDefault ?? string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    private static string SafeControlType(AutomationElement element)
    {
        try
        {
            return element.Properties.ControlType.ValueOrDefault.ToString();
        }
        catch
        {
            return string.Empty;
        }
    }

    private static System.Drawing.Rectangle SafeBoundingRectangle(AutomationElement element)
    {
        try
        {
            return element.Properties.BoundingRectangle.ValueOrDefault;
        }
        catch
        {
            return System.Drawing.Rectangle.Empty;
        }
    }

    private static void TryFocusProcess(Process? process, TimeSpan timeout)
    {
        if (process == null)
        {
            return;
        }

        var end = DateTime.UtcNow + timeout;
        while (DateTime.UtcNow < end)
        {
            try
            {
                process.Refresh();
                var handle = process.MainWindowHandle;
                if (handle != IntPtr.Zero)
                {
                    NativeMethods.ShowWindow(handle, NativeMethods.SW_RESTORE);
                    NativeMethods.SetForegroundWindow(handle);
                    return;
                }
            }
            catch
            {
                return;
            }

            Thread.Sleep(120);
        }
    }

    private static bool TryFocusByHint(string hint, TimeSpan timeout)
    {
        var normalized = Normalize(hint);
        if (string.IsNullOrWhiteSpace(normalized))
        {
            return false;
        }

        var end = DateTime.UtcNow + timeout;
        while (DateTime.UtcNow < end)
        {
            var candidates = new List<IntPtr>();
            NativeMethods.EnumWindows((hWnd, lParam) =>
            {
                if (!NativeMethods.IsWindowVisible(hWnd))
                {
                    return true;
                }

                var title = GetWindowTitle(hWnd);
                if (!string.IsNullOrWhiteSpace(title) && Normalize(title).Contains(normalized, StringComparison.OrdinalIgnoreCase))
                {
                    candidates.Add(hWnd);
                }

                return true;
            }, IntPtr.Zero);

            var target = candidates.FirstOrDefault();
            if (target != IntPtr.Zero)
            {
                NativeMethods.ShowWindow(target, NativeMethods.SW_RESTORE);
                NativeMethods.SetForegroundWindow(target);
                return true;
            }

            Thread.Sleep(120);
        }

        return false;
    }

    private static string GetWindowTitle(IntPtr hWnd)
    {
        var sb = new System.Text.StringBuilder(256);
        var len = NativeMethods.GetWindowText(hWnd, sb, sb.Capacity);
        if (len <= 0)
        {
            return string.Empty;
        }

        return sb.ToString();
    }

    private static class NativeMethods
    {
        public const int SW_RESTORE = 9;

        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        public static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        public static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();

        public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
    }
}
