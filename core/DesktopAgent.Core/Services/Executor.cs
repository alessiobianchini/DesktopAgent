using DesktopAgent.Core.Abstractions;
using DesktopAgent.Core.Config;
using DesktopAgent.Core.Models;
using DesktopAgent.Proto;
using Microsoft.Extensions.Logging;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace DesktopAgent.Core.Services;

public sealed class Executor : IExecutor
{
    private readonly IDesktopAdapterClient _client;
    private readonly IContextProvider _contextProvider;
    private readonly IAppResolver _appResolver;
    private readonly IPolicyEngine _policyEngine;
    private readonly IRateLimiter _rateLimiter;
    private readonly IAuditLog _auditLog;
    private readonly IUserConfirmation _confirmation;
    private readonly IKillSwitch _killSwitch;
    private readonly AgentConfig _config;
    private readonly IOcrEngine _ocrEngine;
    private readonly ILogger<Executor> _logger;
    private static readonly object ClipboardHistoryLock = new();
    private static readonly List<ClipboardHistoryEntry> ClipboardHistory = new();
    private readonly object _recordingSync = new();
    private Process? _recordingProcess;
    private string? _recordingOutputPath;
    private DateTimeOffset? _recordingStartedAtUtc;
    private bool _recordingAudio;
    private string? _recordingAudioBackend;
    private const string CapturePerScreenMode = "mode:per-screen";
    private const string CaptureSingleMode = "mode:single";
    private const int MaxScreenCaptureCount = 16;

    public Executor(
        IDesktopAdapterClient client,
        IContextProvider contextProvider,
        IAppResolver appResolver,
        IPolicyEngine policyEngine,
        IRateLimiter rateLimiter,
        IAuditLog auditLog,
        IUserConfirmation confirmation,
        IKillSwitch killSwitch,
        AgentConfig config,
        IOcrEngine ocrEngine,
        ILogger<Executor> logger)
    {
        _client = client;
        _contextProvider = contextProvider;
        _appResolver = appResolver;
        _policyEngine = policyEngine;
        _rateLimiter = rateLimiter;
        _auditLog = auditLog;
        _confirmation = confirmation;
        _killSwitch = killSwitch;
        _config = config;
        _ocrEngine = ocrEngine;
        _logger = logger;
    }

    public async Task<ExecutionResult> ExecutePlanAsync(ActionPlan plan, bool dryRun, CancellationToken cancellationToken)
    {
        var result = new ExecutionResult { Success = true, Message = "Completed" };
        var binding = new RuntimeBinding();

        for (var i = 0; i < plan.Steps.Count; i++)
        {
            var step = plan.Steps[i];
            if (_killSwitch.IsTripped)
            {
                var msg = $"Kill switch tripped: {_killSwitch.Reason}";
                result.Success = false;
                result.Message = msg;
                result.Steps.Add(new StepResult { Index = i, Type = step.Type, Success = false, Message = msg });
                await WriteAuditAsync("kill", msg, step, cancellationToken);
                break;
            }

            if (!_rateLimiter.TryAcquire())
            {
                var msg = "Rate limit exceeded";
                result.Success = false;
                result.Message = msg;
                result.Steps.Add(new StepResult { Index = i, Type = step.Type, Success = false, Message = msg });
                await WriteAuditAsync("rate_limit", msg, step, cancellationToken);
                break;
            }

            var activeWindow = await _client.GetActiveWindowAsync(cancellationToken);
            EnsureDefaultBinding(step, activeWindow, binding);

            var bindingDecision = EvaluateContextBinding(step, activeWindow, binding);
            if (!bindingDecision.Allowed)
            {
                result.Success = false;
                result.Message = bindingDecision.Reason;
                result.Steps.Add(new StepResult { Index = i, Type = step.Type, Success = false, Message = bindingDecision.Reason });
                await WriteAuditAsync("context_block", bindingDecision.Reason, step, cancellationToken);
                break;
            }

            var decision = _policyEngine.Evaluate(step, activeWindow);
            if (!decision.Allowed)
            {
                result.Success = false;
                result.Message = decision.Reason;
                result.Steps.Add(new StepResult { Index = i, Type = step.Type, Success = false, Message = decision.Reason });
                await WriteAuditAsync("policy_block", decision.Reason, step, cancellationToken);
                break;
            }

            if (decision.RequiresConfirmation)
            {
                var confirm = await _confirmation.ConfirmAsync(decision.Reason, cancellationToken);
                if (!confirm)
                {
                    var msg = "User declined confirmation";
                    result.Success = false;
                    result.Message = msg;
                    result.Steps.Add(new StepResult { Index = i, Type = step.Type, Success = false, Message = msg });
                    await WriteAuditAsync("user_declined", msg, step, cancellationToken);
                    break;
                }
            }

            if (dryRun)
            {
                result.Steps.Add(new StepResult { Index = i, Type = step.Type, Success = true, Message = "Dry-run" });
                await WriteAuditAsync("dry_run", "Dry-run", step, cancellationToken);
                continue;
            }

            var stepResult = await ExecuteStepAsync(step, cancellationToken);
            stepResult.Index = i;
            stepResult.Type = step.Type;
            result.Steps.Add(stepResult);

            await WriteAuditAsync(stepResult.Success ? "action" : "action_failed", stepResult.Message, step, cancellationToken);

            if (!stepResult.Success)
            {
                result.Success = false;
                result.Message = stepResult.Message;
                break;
            }

            await RefreshBindingAfterStepAsync(step, binding, cancellationToken);
        }

        return result;
    }

    private async Task<StepResult> ExecuteStepAsync(PlanStep step, CancellationToken cancellationToken)
    {
        try
        {
            switch (step.Type)
            {
                case ActionType.Find:
                    return await ExecuteFindAsync(step, cancellationToken);
                case ActionType.Click:
                    return await ExecuteClickAsync(step, cancellationToken);
                case ActionType.DoubleClick:
                    return await ExecuteDoubleClickAsync(step, cancellationToken);
                case ActionType.RightClick:
                    return await ExecuteRightClickAsync(step, cancellationToken);
                case ActionType.Drag:
                    return await ExecuteDragAsync(step, cancellationToken);
                case ActionType.TypeText:
                    return await ExecuteActionWithPostCheckAsync(
                        step,
                        () => _client.TypeTextAsync(step.Text ?? string.Empty, cancellationToken),
                        "Typed",
                        cancellationToken);
                case ActionType.OpenApp:
                    return await ExecuteOpenAppAsync(step, cancellationToken);
                case ActionType.KeyCombo:
                    return await ExecuteSimpleResultAsync(() => _client.KeyComboAsync(step.Keys ?? new List<string>(), cancellationToken), "Key combo");
                case ActionType.Invoke:
                    return await ExecuteActionWithPostCheckAsync(
                        step,
                        () => _client.InvokeElementAsync(step.ElementId ?? string.Empty, cancellationToken),
                        "Invoked",
                        cancellationToken);
                case ActionType.SetValue:
                    return await ExecuteActionWithPostCheckAsync(
                        step,
                        () => _client.SetElementValueAsync(step.ElementId ?? string.Empty, step.Text ?? string.Empty, cancellationToken),
                        "Value set",
                        cancellationToken);
                case ActionType.ReadText:
                    return await ExecuteReadTextAsync(cancellationToken);
                case ActionType.CaptureScreen:
                    return await ExecuteCaptureAsync(step, cancellationToken);
                case ActionType.RecordScreen:
                    return await ExecuteRecordScreenAsync(step, cancellationToken);
                case ActionType.StartScreenRecording:
                    return await ExecuteStartScreenRecordingAsync(step, cancellationToken);
                case ActionType.StopScreenRecording:
                    return await ExecuteStopScreenRecordingAsync(cancellationToken);
                case ActionType.GetClipboard:
                    var clip = await _client.GetClipboardAsync(cancellationToken);
                    AddClipboardHistory("get", clip.Text);
                    return new StepResult { Success = true, Message = "Clipboard read", Data = clip.Text };
                case ActionType.SetClipboard:
                    var clipText = step.Text ?? string.Empty;
                    var clipResult = await ExecuteSimpleResultAsync(() => _client.SetClipboardAsync(clipText, cancellationToken), "Clipboard set");
                    if (clipResult.Success)
                    {
                        AddClipboardHistory("set", clipText);
                    }
                    return clipResult;
                case ActionType.OpenUrl:
                    return await ExecuteOpenUrlAsync(step, cancellationToken);
                case ActionType.FileWrite:
                    return await ExecuteFileWriteAsync(step, append: false, cancellationToken);
                case ActionType.FileAppend:
                    return await ExecuteFileWriteAsync(step, append: true, cancellationToken);
                case ActionType.FileRead:
                    return await ExecuteFileReadAsync(step, cancellationToken);
                case ActionType.FileList:
                    return await ExecuteFileListAsync(step, cancellationToken);
                case ActionType.ClipboardHistory:
                    return new StepResult { Success = true, Message = "Clipboard history", Data = GetClipboardHistory() };
                case ActionType.Notify:
                    return await ExecuteNotifyAsync(step, cancellationToken);
                case ActionType.MouseJiggle:
                    return await ExecuteMouseJiggleAsync(step, cancellationToken);
                case ActionType.WaitForText:
                    return await ExecuteWaitForTextAsync(step, cancellationToken);
                case ActionType.VolumeUp:
                    return await ExecuteVolumeAsync(up: true, mute: false, step, cancellationToken);
                case ActionType.VolumeDown:
                    return await ExecuteVolumeAsync(up: false, mute: false, step, cancellationToken);
                case ActionType.VolumeMute:
                    return await ExecuteVolumeAsync(up: false, mute: true, step, cancellationToken);
                case ActionType.BrightnessUp:
                    return await ExecuteBrightnessAsync(up: true, step, cancellationToken);
                case ActionType.BrightnessDown:
                    return await ExecuteBrightnessAsync(up: false, step, cancellationToken);
                case ActionType.LockScreen:
                    return await ExecuteLockScreenAsync(cancellationToken);
                case ActionType.WaitFor:
                    if (step.WaitFor.HasValue)
                    {
                        await Task.Delay(step.WaitFor.Value, cancellationToken);
                    }
                    return new StepResult { Success = true, Message = "Waited" };
                default:
                    return new StepResult { Success = false, Message = "Unknown action" };
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Step execution failed");
            return new StepResult { Success = false, Message = ex.Message };
        }
    }

    private async Task<StepResult> ExecuteFindAsync(PlanStep step, CancellationToken cancellationToken)
    {
        var selector = BuildEffectiveSelector(step);
        var elements = await FindElementsWithRetryAsync(selector, cancellationToken);
        elements = RankElements(elements, selector, step.Type);
        if (elements.Count > 0)
        {
            return new StepResult { Success = true, Message = $"Found {elements.Count} elements", Data = elements };
        }

        if (!_config.OcrEnabled)
        {
            return new StepResult { Success = false, Message = "No elements found and OCR disabled" };
        }

        var find = await _contextProvider.FindByTextAsync(selector.NameContains ?? step.Text ?? string.Empty, cancellationToken);
        if (find.OcrMatches.Count > 0)
        {
            var ranked = RankOcrMatches(find.OcrMatches, selector.NameContains ?? step.Text ?? string.Empty);
            return new StepResult { Success = true, Message = $"OCR matched {ranked.Count} regions", Data = ranked };
        }

        return new StepResult { Success = false, Message = "No elements found" };
    }

    private async Task<StepResult> ExecuteClickAsync(PlanStep step, CancellationToken cancellationToken)
    {
        if (step.Point != null)
        {
            var res = await _client.ClickPointAsync(step.Point.X, step.Point.Y, cancellationToken);
            return new StepResult { Success = res.Success, Message = res.Message };
        }

        var selector = BuildEffectiveSelector(step);
        var elements = await FindElementsWithRetryAsync(selector, cancellationToken);
        elements = RankElements(elements, selector, step.Type);
        if (elements.Count > 0)
        {
            var target = elements[0];
            var preWindow = await _client.GetActiveWindowAsync(cancellationToken);
            var (x, y) = Center(target.Bounds);
            var res = await _client.ClickPointAsync(x, y, cancellationToken);
            if (!res.Success)
            {
                return new StepResult { Success = false, Message = res.Message, Data = target };
            }

            var post = await VerifyPostConditionAsync(step, selector, preWindow, elements.Count, target.Role, cancellationToken);
            return BuildPostCheckedResult(res.Success, res.Message, target, post, _config.PostCheckStrict);
        }

        if (_config.OcrEnabled)
        {
            var find = await _contextProvider.FindByTextAsync(selector.NameContains ?? step.Text ?? string.Empty, cancellationToken);
            if (find.OcrMatches.Count > 0)
            {
                var region = SelectBestOcrMatch(find.OcrMatches, selector.NameContains ?? step.Text ?? string.Empty);
                var preWindow = await _client.GetActiveWindowAsync(cancellationToken);
                var (x, y) = Center(region.Bounds);
                var res = await _client.ClickPointAsync(x, y, cancellationToken);
                if (!res.Success)
                {
                    return new StepResult { Success = false, Message = res.Message, Data = region };
                }

                var post = await VerifyPostConditionAsync(step, selector, preWindow, 0, null, cancellationToken);
                return BuildPostCheckedResult(res.Success, res.Message, region, post, _config.PostCheckStrict);
            }
        }

        return new StepResult { Success = false, Message = "No target to click" };
    }

    private async Task<StepResult> ExecuteDoubleClickAsync(PlanStep step, CancellationToken cancellationToken)
    {
        var resolved = await ResolveActionPointAsync(step, cancellationToken);
        if (!resolved.Success)
        {
            return new StepResult { Success = false, Message = resolved.Error };
        }

        var first = await _client.ClickPointAsync(resolved.X, resolved.Y, cancellationToken);
        if (!first.Success)
        {
            return new StepResult { Success = false, Message = first.Message };
        }

        await Task.Delay(90, cancellationToken);
        var second = await _client.ClickPointAsync(resolved.X, resolved.Y, cancellationToken);
        if (!second.Success)
        {
            return new StepResult { Success = false, Message = second.Message };
        }

        var post = await VerifyPostConditionAsync(step, resolved.Selector, resolved.PreWindow, resolved.PreCount, resolved.TargetRole, cancellationToken);
        return BuildPostCheckedResult(true, "Double clicked", resolved.Data, post, _config.PostCheckStrict);
    }

    private async Task<StepResult> ExecuteRightClickAsync(PlanStep step, CancellationToken cancellationToken)
    {
        var resolved = await ResolveActionPointAsync(step, cancellationToken);
        if (!resolved.Success)
        {
            return new StepResult { Success = false, Message = resolved.Error };
        }

        if (OperatingSystem.IsWindows())
        {
            if (!TryMoveCursor(resolved.X, resolved.Y, out var moveError))
            {
                return new StepResult { Success = false, Message = moveError ?? "Failed moving cursor" };
            }

            MouseEvent(MouseEventRightDown, 0, 0, 0, UIntPtr.Zero);
            MouseEvent(MouseEventRightUp, 0, 0, 0, UIntPtr.Zero);
            return new StepResult { Success = true, Message = "Right clicked", Data = resolved.Data };
        }

        if (OperatingSystem.IsLinux())
        {
            var command = $"mousemove --sync {resolved.X} {resolved.Y} click 3";
            if (TryRunProcess("xdotool", command, out _, out var xdotoolError))
            {
                return new StepResult { Success = true, Message = "Right clicked", Data = resolved.Data };
            }

            return new StepResult { Success = false, Message = xdotoolError ?? "Right click failed (xdotool missing?)" };
        }

        var focus = await _client.ClickPointAsync(resolved.X, resolved.Y, cancellationToken);
        if (!focus.Success)
        {
            return new StepResult { Success = false, Message = focus.Message };
        }

        var fallback = await _client.KeyComboAsync(new List<string> { "shift", "f10" }, cancellationToken);
        return new StepResult
        {
            Success = fallback.Success,
            Message = fallback.Success ? "Right clicked (fallback)" : fallback.Message,
            Data = resolved.Data
        };
    }

    private async Task<StepResult> ExecuteDragAsync(PlanStep step, CancellationToken cancellationToken)
    {
        var sourceStep = new PlanStep
        {
            Type = ActionType.Drag,
            Selector = step.Selector,
            Text = step.Text,
            ExpectedAppId = step.ExpectedAppId,
            ExpectedWindowId = step.ExpectedWindowId
        };
        var sourceResolved = await ResolveActionPointAsync(sourceStep, cancellationToken);
        if (!sourceResolved.Success)
        {
            return new StepResult { Success = false, Message = $"Drag source not found: {sourceResolved.Error}" };
        }

        var targetSelector = new Selector { NameContains = step.Target ?? string.Empty };
        if (string.IsNullOrWhiteSpace(targetSelector.NameContains))
        {
            return new StepResult { Success = false, Message = "Drag target missing" };
        }

        var targetStep = new PlanStep
        {
            Type = ActionType.Drag,
            Selector = targetSelector,
            ExpectedAppId = step.ExpectedAppId,
            ExpectedWindowId = step.ExpectedWindowId
        };
        var targetResolved = await ResolveActionPointAsync(targetStep, cancellationToken);
        if (!targetResolved.Success)
        {
            return new StepResult { Success = false, Message = $"Drag target not found: {targetResolved.Error}" };
        }

        if (OperatingSystem.IsWindows())
        {
            if (!TryMoveCursor(sourceResolved.X, sourceResolved.Y, out var moveError))
            {
                return new StepResult { Success = false, Message = moveError ?? "Failed moving cursor to source" };
            }

            MouseEvent(MouseEventLeftDown, 0, 0, 0, UIntPtr.Zero);
            var steps = 12;
            for (var i = 1; i <= steps; i++)
            {
                var x = sourceResolved.X + ((targetResolved.X - sourceResolved.X) * i / steps);
                var y = sourceResolved.Y + ((targetResolved.Y - sourceResolved.Y) * i / steps);
                if (!TryMoveCursor(x, y, out _))
                {
                    MouseEvent(MouseEventLeftUp, 0, 0, 0, UIntPtr.Zero);
                    return new StepResult { Success = false, Message = "Drag failed during mouse move" };
                }

                await Task.Delay(12, cancellationToken);
            }

            MouseEvent(MouseEventLeftUp, 0, 0, 0, UIntPtr.Zero);
        }
        else if (OperatingSystem.IsLinux())
        {
            var command = $"mousemove --sync {sourceResolved.X} {sourceResolved.Y} mousedown 1 mousemove --sync {targetResolved.X} {targetResolved.Y} mouseup 1";
            if (!TryRunProcess("xdotool", command, out _, out var xdotoolError))
            {
                return new StepResult { Success = false, Message = xdotoolError ?? "Drag failed (xdotool missing?)" };
            }
        }
        else
        {
            return new StepResult { Success = false, Message = "Drag not supported on this OS in current build" };
        }

        var post = await VerifyPostConditionAsync(step, sourceResolved.Selector, sourceResolved.PreWindow, sourceResolved.PreCount, sourceResolved.TargetRole, cancellationToken);
        return BuildPostCheckedResult(true, "Dragged", new { from = sourceResolved.Data, to = targetResolved.Data }, post, _config.PostCheckStrict);
    }

    private async Task<StepResult> ExecuteWaitForTextAsync(PlanStep step, CancellationToken cancellationToken)
    {
        var text = step.Text
                   ?? step.Selector?.NameContains
                   ?? string.Empty;
        if (string.IsNullOrWhiteSpace(text))
        {
            return new StepResult { Success = false, Message = "Missing text to wait for" };
        }

        var timeout = step.WaitFor.GetValueOrDefault(TimeSpan.FromSeconds(30));
        if (timeout <= TimeSpan.Zero)
        {
            timeout = TimeSpan.FromSeconds(1);
        }

        if (timeout > TimeSpan.FromMinutes(10))
        {
            timeout = TimeSpan.FromMinutes(10);
        }

        var started = DateTimeOffset.UtcNow;
        var deadline = started + timeout;
        while (DateTimeOffset.UtcNow < deadline)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var found = await _contextProvider.FindByTextAsync(text, cancellationToken);
            if (found.Elements.Count > 0 || found.OcrMatches.Count > 0)
            {
                return new StepResult
                {
                    Success = true,
                    Message = "Text detected",
                    Data = new { text, uiMatches = found.Elements.Count, ocrMatches = found.OcrMatches.Count }
                };
            }

            await Task.Delay(250, cancellationToken);
        }

        return new StepResult
        {
            Success = false,
            Message = $"Timeout waiting for text: {text}",
            Data = new { text, seconds = Math.Round(timeout.TotalSeconds, 1) }
        };
    }

    private async Task<StepResult> ExecuteReadTextAsync(CancellationToken cancellationToken)
    {
        if (!_config.OcrEnabled)
        {
            return new StepResult { Success = false, Message = "OCR disabled" };
        }

        var screenshot = await _client.CaptureScreenAsync(new ScreenshotRequest(), cancellationToken);
        if (screenshot.Png.IsEmpty)
        {
            return new StepResult { Success = false, Message = "Screenshot failed" };
        }

        var regions = await _ocrEngine.ReadTextAsync(screenshot.Png.ToByteArray(), cancellationToken);
        return new StepResult { Success = true, Message = "Read text", Data = regions };
    }

    private async Task<StepResult> ExecuteCaptureAsync(PlanStep step, CancellationToken cancellationToken)
    {
        if (IsSingleScreenCapture(step))
        {
            return await ExecuteSingleCaptureAsync(step, cancellationToken);
        }

        if (WantsPerScreenCapture(step))
        {
            return await ExecuteCapturePerScreenAsync(step, cancellationToken);
        }

        return await ExecuteSingleCaptureAsync(step, cancellationToken);
    }

    private async Task<StepResult> ExecuteSingleCaptureAsync(PlanStep step, CancellationToken cancellationToken)
    {
        var screenshot = await _client.CaptureScreenAsync(new ScreenshotRequest(), cancellationToken);
        if (screenshot.Png.IsEmpty)
        {
            return new StepResult { Success = false, Message = "Screenshot failed" };
        }

        try
        {
            var path = BuildMediaOutputPath(step.Target, "snapshot", ".png");
            await File.WriteAllBytesAsync(path, screenshot.Png.ToByteArray(), cancellationToken);
            return new StepResult
            {
                Success = true,
                Message = $"Snapshot saved: {path}",
                Data = new { path, screenshot.Width, screenshot.Height, bytes = screenshot.Png.Length }
            };
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Saving screenshot failed");
            return new StepResult { Success = false, Message = $"Screenshot save failed: {ex.Message}" };
        }
    }

    private async Task<StepResult> ExecuteCapturePerScreenAsync(PlanStep step, CancellationToken cancellationToken)
    {
        var supportsIndexedCapture = await SupportsIndexedScreenCaptureAsync(cancellationToken);
        if (!supportsIndexedCapture)
        {
            var fallback = await ExecuteSingleCaptureAsync(step, cancellationToken);
            if (fallback.Success)
            {
                fallback.Message += " (adapter does not support per-screen capture yet)";
            }

            return fallback;
        }

        var outputTemplate = BuildMediaOutputPath(step.Target, "snapshot", ".png");
        var directory = Path.GetDirectoryName(outputTemplate) ?? AppContext.BaseDirectory;
        var baseName = Path.GetFileNameWithoutExtension(outputTemplate);
        var saved = new List<object>();

        try
        {
            for (var index = 0; index < MaxScreenCaptureCount; index++)
            {
                var request = new ScreenshotRequest { WindowId = $"__screen:{index}" };
                var screenshot = await _client.CaptureScreenAsync(request, cancellationToken);
                if (screenshot.Png.IsEmpty)
                {
                    break;
                }

                var path = Path.Combine(directory, $"{baseName}-screen{index + 1}.png");
                await File.WriteAllBytesAsync(path, screenshot.Png.ToByteArray(), cancellationToken);
                saved.Add(new { screen = index + 1, path, screenshot.Width, screenshot.Height, bytes = screenshot.Png.Length });
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Saving per-screen screenshots failed");
            return new StepResult { Success = false, Message = $"Screenshot save failed: {ex.Message}" };
        }

        if (saved.Count == 0)
        {
            return new StepResult { Success = false, Message = "Per-screen screenshot failed" };
        }

        return new StepResult
        {
            Success = true,
            Message = $"Snapshots saved for {saved.Count} screen(s)",
            Data = saved
        };
    }

    private async Task<bool> SupportsIndexedScreenCaptureAsync(CancellationToken cancellationToken)
    {
        var probe = await _client.CaptureScreenAsync(new ScreenshotRequest { WindowId = "__screen:-1" }, cancellationToken);
        return probe.Png.IsEmpty;
    }

    private static bool WantsPerScreenCapture(PlanStep step)
    {
        if (string.IsNullOrWhiteSpace(step.Text))
        {
            return true;
        }

        return string.Equals(step.Text, CapturePerScreenMode, StringComparison.OrdinalIgnoreCase)
               || string.Equals(step.Text, "per-screen", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsSingleScreenCapture(PlanStep step)
    {
        return string.Equals(step.Text, CaptureSingleMode, StringComparison.OrdinalIgnoreCase)
               || string.Equals(step.Text, "single-screen", StringComparison.OrdinalIgnoreCase);
    }

    private async Task<StepResult> ExecuteRecordScreenAsync(PlanStep step, CancellationToken cancellationToken)
    {
        var duration = step.WaitFor ?? TimeSpan.FromSeconds(30);
        if (duration < TimeSpan.FromSeconds(1))
        {
            duration = TimeSpan.FromSeconds(1);
        }
        if (duration > TimeSpan.FromHours(2))
        {
            duration = TimeSpan.FromHours(2);
        }

        var includeAudioRequested = !string.Equals(step.Text, "audio:off", StringComparison.OrdinalIgnoreCase);
        var outputPath = BuildMediaOutputPath(step.Target, "recording", ".mp4");

        var seconds = Math.Max(1, (int)Math.Ceiling(duration.TotalSeconds));
        var ffmpegPath = ResolveFfmpegPath();

        var timeout = TimeSpan.FromSeconds(seconds + 30);
        var audioBackends = includeAudioRequested ? ResolveAvailableAudioInputs(ffmpegPath) : new List<FfmpegAudioInput>();
        var audioErrors = new List<string>();
        if (includeAudioRequested)
        {
            if (audioBackends.Count == 0)
            {
                audioErrors.Add(GetNoSupportedAudioBackendMessage());
            }
            else
            {
                foreach (var backend in audioBackends)
                {
                    var withAudioArgs = BuildRecordingArguments(ffmpegPath, outputPath, seconds, backend, out var withAudioError);
                    if (string.IsNullOrWhiteSpace(withAudioArgs))
                    {
                        audioErrors.Add($"{backend.DisplayName}: {withAudioError}");
                        continue;
                    }

                    var withAudio = await RunProcessAsync(ffmpegPath, withAudioArgs, timeout, cancellationToken);
                    if (withAudio.Success)
                    {
                        return new StepResult
                        {
                            Success = true,
                            Message = $"Screen recording saved (audio {backend.DisplayName}): {outputPath}",
                            Data = new { outputPath, seconds, audio = true, audioBackend = backend.DisplayName }
                        };
                    }

                    var failure = $"{backend.DisplayName}: {TrimForMessage(withAudio.Error)}";
                    audioErrors.Add(failure);
                    _logger.LogWarning("Screen recording with audio backend {Backend} failed. Error: {Error}", backend.DisplayName, withAudio.Error);
                }
            }
        }

        var videoOnlyArgs = BuildRecordingArguments(ffmpegPath, outputPath, seconds, audioInput: null, out var videoOnlyArgsError);
        if (string.IsNullOrWhiteSpace(videoOnlyArgs))
        {
            return new StepResult
            {
                Success = false,
                Message = $"Screen recording failed: {videoOnlyArgsError ?? "No supported video capture backend"}",
                Data = new { outputPath, seconds, includeAudioRequested }
            };
        }

        var videoOnly = await RunProcessAsync(ffmpegPath, videoOnlyArgs, timeout, cancellationToken);
        if (videoOnly.Success)
        {
            var reason = includeAudioRequested
                ? TrimForMessage(string.Join(" | ", audioErrors.Where(static s => !string.IsNullOrWhiteSpace(s))))
                : null;
            return new StepResult
            {
                Success = true,
                Message = includeAudioRequested
                    ? $"Screen recording saved (audio unavailable, recorded video-only): {outputPath}{(string.IsNullOrWhiteSpace(reason) ? string.Empty : $" | reason: {reason}")}"
                    : $"Screen recording saved: {outputPath}",
                Data = new { outputPath, seconds, audio = false, audioReason = reason, audioBackendsTried = audioBackends.Select(static x => x.DisplayName).ToArray() }
            };
        }

        return new StepResult
        {
            Success = false,
            Message = $"Screen recording failed: {videoOnly.Error ?? "ffmpeg not available or capture failed"}",
            Data = new { outputPath, seconds, includeAudioRequested }
        };
    }

    private Task<StepResult> ExecuteStartScreenRecordingAsync(PlanStep step, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var includeAudioRequested = !string.Equals(step.Text, "audio:off", StringComparison.OrdinalIgnoreCase);
        var outputPath = BuildMediaOutputPath(step.Target, "recording-live", ".mp4");

        lock (_recordingSync)
        {
            if (_recordingProcess is { HasExited: false })
            {
                var runningFor = _recordingStartedAtUtc.HasValue
                    ? (DateTimeOffset.UtcNow - _recordingStartedAtUtc.Value).TotalSeconds
                    : 0;
                return Task.FromResult(new StepResult
                {
                    Success = false,
                    Message = "A screen recording is already running.",
                    Data = new { outputPath = _recordingOutputPath, seconds = Math.Max(0, (int)runningFor), audio = _recordingAudio }
                });
            }

            if (_recordingProcess is { HasExited: true })
            {
                _recordingProcess.Dispose();
                _recordingProcess = null;
                _recordingOutputPath = null;
                _recordingStartedAtUtc = null;
                _recordingAudio = false;
                _recordingAudioBackend = null;
            }
        }

        var ffmpegPath = ResolveFfmpegPath();
        if (TryStartRecordingProcess(ffmpegPath, outputPath, includeAudioRequested, out var audioUsed, out var audioBackend, out var error))
        {
            lock (_recordingSync)
            {
                _recordingAudio = audioUsed;
                _recordingAudioBackend = audioBackend;
            }

            return Task.FromResult(new StepResult
            {
                Success = true,
                Message = audioUsed
                    ? $"Screen recording started (with audio {audioBackend ?? "unknown"}): {outputPath}"
                    : $"Screen recording started (video-only): {outputPath}",
                Data = new { outputPath, audio = audioUsed, audioBackend, error }
            });
        }

        return Task.FromResult(new StepResult
        {
            Success = false,
            Message = $"Start recording failed: {error ?? "ffmpeg not available or capture failed"}",
            Data = new { outputPath, includeAudioRequested }
        });
    }

    private Task<StepResult> ExecuteStopScreenRecordingAsync(CancellationToken cancellationToken)
    {
        Process? process;
        string? outputPath;
        DateTimeOffset? startedAtUtc;
        bool audio;
        string? audioBackend;

        lock (_recordingSync)
        {
            process = _recordingProcess;
            outputPath = _recordingOutputPath;
            startedAtUtc = _recordingStartedAtUtc;
            audio = _recordingAudio;
            audioBackend = _recordingAudioBackend;
            _recordingProcess = null;
            _recordingOutputPath = null;
            _recordingStartedAtUtc = null;
            _recordingAudio = false;
            _recordingAudioBackend = null;
        }

        if (process == null || process.HasExited)
        {
            return Task.FromResult(new StepResult
            {
                Success = false,
                Message = "No active recording to stop."
            });
        }

        try
        {
            process.Kill(entireProcessTree: true);
            process.WaitForExit(5000);
        }
        catch (Exception ex)
        {
            return Task.FromResult(new StepResult { Success = false, Message = $"Stop recording failed: {ex.Message}" });
        }
        finally
        {
            process.Dispose();
        }

        var elapsedSeconds = startedAtUtc.HasValue
            ? Math.Max(0, (int)(DateTimeOffset.UtcNow - startedAtUtc.Value).TotalSeconds)
            : 0;

        long? bytes = null;
        if (!string.IsNullOrWhiteSpace(outputPath) && File.Exists(outputPath))
        {
            bytes = new FileInfo(outputPath).Length;
        }

        return Task.FromResult(new StepResult
        {
            Success = true,
            Message = string.IsNullOrWhiteSpace(outputPath)
                ? "Screen recording stopped."
                : $"Screen recording stopped: {outputPath}",
            Data = new { outputPath, seconds = elapsedSeconds, audio, audioBackend, bytes }
        });
    }

    private async Task<StepResult> ExecuteOpenAppAsync(PlanStep step, CancellationToken cancellationToken)
    {
        var requested = (step.AppIdOrPath ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(requested))
        {
            return new StepResult { Success = false, Message = "Missing app id or path" };
        }

        var target = requested;
        if (_appResolver.TryResolveApp(requested, out var resolved))
        {
            target = resolved;
        }

        var result = await _client.OpenAppAsync(target, cancellationToken);
        if (result.Success)
        {
            var settleDelayMs = Math.Max(0, _config.OpenAppSettleDelayMs);
            if (settleDelayMs > 0)
            {
                await WaitForAppActivationAsync(requested, target, settleDelayMs, cancellationToken);
            }

            return new StepResult
            {
                Success = true,
                Message = string.Equals(target, requested, StringComparison.OrdinalIgnoreCase)
                    ? "Opened app"
                    : $"Opened app (resolved \"{requested}\" -> \"{target}\")"
            };
        }

        var suggestions = _appResolver.Suggest(requested, 3)
            .Select(match => match.Entry.Name)
            .Where(name => !string.IsNullOrWhiteSpace(name))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        var suggestionText = suggestions.Count > 0
            ? $" Suggestions: {string.Join(", ", suggestions)}."
            : string.Empty;

        if (!string.Equals(target, requested, StringComparison.OrdinalIgnoreCase))
        {
            return new StepResult
            {
                Success = false,
                Message = $"{result.Message} (resolved \"{requested}\" -> \"{target}\").{suggestionText}"
            };
        }

        return new StepResult
        {
            Success = false,
            Message = $"{result.Message}{suggestionText}"
        };
    }

    private async Task WaitForAppActivationAsync(string requestedTarget, string resolvedTarget, int maxWaitMs, CancellationToken cancellationToken)
    {
        if (maxWaitMs <= 0)
        {
            return;
        }

        var expected = BuildExpectedAppTokens(requestedTarget, resolvedTarget);
        if (expected.Count == 0)
        {
            await Task.Delay(maxWaitMs, cancellationToken);
            return;
        }

        var started = DateTime.UtcNow;
        var deadline = started.AddMilliseconds(maxWaitMs);
        var pollMs = Math.Clamp(maxWaitMs / 8, 25, 120);

        while (DateTime.UtcNow < deadline)
        {
            var activeWindow = await _client.GetActiveWindowAsync(cancellationToken);
            if (activeWindow != null && expected.Any(token => AppIdMatches(activeWindow.AppId, token)))
            {
                return;
            }

            var remaining = (int)(deadline - DateTime.UtcNow).TotalMilliseconds;
            if (remaining <= 0)
            {
                break;
            }

            await Task.Delay(Math.Min(pollMs, remaining), cancellationToken);
        }
    }

    private static List<string> BuildExpectedAppTokens(string requestedTarget, string resolvedTarget)
    {
        var expected = new List<string>();
        TryAddAppToken(expected, requestedTarget);
        TryAddAppToken(expected, resolvedTarget);
        return expected;
    }

    private static void TryAddAppToken(List<string> tokens, string? value)
    {
        var normalized = NormalizeAppToken(value ?? string.Empty);
        if (string.IsNullOrWhiteSpace(normalized))
        {
            return;
        }

        if (!tokens.Any(existing => string.Equals(existing, normalized, StringComparison.OrdinalIgnoreCase)))
        {
            tokens.Add(normalized);
        }
    }

    private Task<StepResult> ExecuteOpenUrlAsync(PlanStep step, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var raw = (step.Target ?? step.Text ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(raw))
        {
            return Task.FromResult(new StepResult { Success = false, Message = "Missing URL target" });
        }

        if (!TryNormalizeUrl(raw, out var normalized))
        {
            return Task.FromResult(new StepResult { Success = false, Message = "Invalid URL" });
        }

        var requestedApp = (step.AppIdOrPath ?? string.Empty).Trim();
        if (!string.IsNullOrWhiteSpace(requestedApp))
        {
            var targetApp = requestedApp;
            if (_appResolver.TryResolveApp(requestedApp, out var resolvedApp))
            {
                targetApp = resolvedApp;
            }

            try
            {
                LaunchUrlInBrowserApp(targetApp, normalized);

                var details = string.Equals(targetApp, requestedApp, StringComparison.OrdinalIgnoreCase)
                    ? string.Empty
                    : $" (resolved \"{requestedApp}\" -> \"{targetApp}\")";
                return Task.FromResult(new StepResult { Success = true, Message = $"Opened URL in app: {normalized}{details}" });
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "OpenUrl with app target failed, fallback to system browser");
            }
        }

        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = normalized,
                UseShellExecute = true
            });
            return Task.FromResult(new StepResult { Success = true, Message = $"Opened URL: {normalized}" });
        }
        catch (Exception ex)
        {
            return Task.FromResult(new StepResult { Success = false, Message = ex.Message });
        }
    }

    private static void LaunchUrlInBrowserApp(string targetApp, string normalizedUrl)
    {
        var normalizedToken = NormalizeAppToken(targetApp);
        var (fileName, args, useShellExecute) = BuildBrowserLaunchCommand(targetApp, normalizedUrl, normalizedToken);
        Process.Start(new ProcessStartInfo
        {
            FileName = fileName,
            Arguments = args,
            UseShellExecute = useShellExecute
        });
    }

    private static (string FileName, string Arguments, bool UseShellExecute) BuildBrowserLaunchCommand(string targetApp, string url, string normalizedToken)
    {
        var quotedUrl = QuoteArgument(url);
        if (normalizedToken.Contains("chrome", StringComparison.OrdinalIgnoreCase)
            || normalizedToken.Contains("msedge", StringComparison.OrdinalIgnoreCase)
            || normalizedToken.Contains("edge", StringComparison.OrdinalIgnoreCase)
            || normalizedToken.Contains("brave", StringComparison.OrdinalIgnoreCase)
            || normalizedToken.Contains("opera", StringComparison.OrdinalIgnoreCase))
        {
            return (targetApp, $"--new-tab {quotedUrl}", false);
        }

        if (normalizedToken.Contains("firefox", StringComparison.OrdinalIgnoreCase))
        {
            return (targetApp, $"-new-tab {quotedUrl}", false);
        }

        if (normalizedToken.Contains("safari", StringComparison.OrdinalIgnoreCase))
        {
            return (targetApp, quotedUrl, true);
        }

        return (targetApp, quotedUrl, true);
    }

    private static string QuoteArgument(string value)
    {
        if (string.IsNullOrEmpty(value))
        {
            return "\"\"";
        }

        if (value.Contains('"'))
        {
            value = value.Replace("\"", "\\\"");
        }

        return $"\"{value}\"";
    }

    private Task<StepResult> ExecuteFileWriteAsync(PlanStep step, bool append, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var target = (step.Target ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(target))
        {
            return Task.FromResult(new StepResult { Success = false, Message = "Missing file path target" });
        }

        if (!TryResolveAllowedPath(target, out var fullPath, out var reason))
        {
            return Task.FromResult(new StepResult { Success = false, Message = reason });
        }

        try
        {
            var dir = Path.GetDirectoryName(fullPath);
            if (!string.IsNullOrWhiteSpace(dir))
            {
                Directory.CreateDirectory(dir);
            }

            var content = step.Text ?? string.Empty;
            if (append)
            {
                File.AppendAllText(fullPath, content);
            }
            else
            {
                File.WriteAllText(fullPath, content);
            }

            return Task.FromResult(new StepResult
            {
                Success = true,
                Message = append ? "File appended" : "File written",
                Data = new { path = fullPath, bytes = content.Length }
            });
        }
        catch (Exception ex)
        {
            return Task.FromResult(new StepResult { Success = false, Message = ex.Message });
        }
    }

    private Task<StepResult> ExecuteFileReadAsync(PlanStep step, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var target = (step.Target ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(target))
        {
            return Task.FromResult(new StepResult { Success = false, Message = "Missing file path target" });
        }

        if (!TryResolveAllowedPath(target, out var fullPath, out var reason))
        {
            return Task.FromResult(new StepResult { Success = false, Message = reason });
        }

        try
        {
            if (!File.Exists(fullPath))
            {
                return Task.FromResult(new StepResult { Success = false, Message = "File not found" });
            }

            var text = File.ReadAllText(fullPath);
            return Task.FromResult(new StepResult
            {
                Success = true,
                Message = "File read",
                Data = new { path = fullPath, content = text, bytes = text.Length }
            });
        }
        catch (Exception ex)
        {
            return Task.FromResult(new StepResult { Success = false, Message = ex.Message });
        }
    }

    private Task<StepResult> ExecuteFileListAsync(PlanStep step, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var target = (step.Target ?? ".").Trim();
        if (!TryResolveAllowedPath(target, out var fullPath, out var reason))
        {
            return Task.FromResult(new StepResult { Success = false, Message = reason });
        }

        try
        {
            if (!Directory.Exists(fullPath))
            {
                return Task.FromResult(new StepResult { Success = false, Message = "Directory not found" });
            }

            var searchQuery = (step.Text ?? string.Empty).Trim();
            if (!string.IsNullOrWhiteSpace(searchQuery))
            {
                var matches = SearchFiles(fullPath, searchQuery, maxResults: 200, cancellationToken)
                    .Select(path => new
                    {
                        name = Path.GetFileName(path),
                        path
                    })
                    .ToList();

                return Task.FromResult(new StepResult
                {
                    Success = true,
                    Message = $"Found {matches.Count} file(s)",
                    Data = new { path = fullPath, query = searchQuery, entries = matches }
                });
            }

            var entries = Directory.EnumerateFileSystemEntries(fullPath)
                .Select(Path.GetFileName)
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .Take(200)
                .ToList();
            return Task.FromResult(new StepResult
            {
                Success = true,
                Message = $"Listed {entries.Count} entries",
                Data = new { path = fullPath, entries }
            });
        }
        catch (Exception ex)
        {
            return Task.FromResult(new StepResult { Success = false, Message = ex.Message });
        }
    }

    private static IEnumerable<string> SearchFiles(string rootPath, string query, int maxResults, CancellationToken cancellationToken)
    {
        var comparer = StringComparison.OrdinalIgnoreCase;
        var wildcard = query.Contains('*') || query.Contains('?')
            ? new Regex("^" + Regex.Escape(query).Replace("\\*", ".*").Replace("\\?", ".") + "$", RegexOptions.IgnoreCase | RegexOptions.Compiled)
            : null;

        var pending = new Stack<string>();
        pending.Push(rootPath);
        var results = new List<string>(Math.Min(maxResults, 200));

        while (pending.Count > 0 && results.Count < maxResults)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var current = pending.Pop();

            IEnumerable<string> files;
            try
            {
                files = Directory.EnumerateFiles(current);
            }
            catch
            {
                continue;
            }

            foreach (var file in files)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var name = Path.GetFileName(file);
                if (string.IsNullOrWhiteSpace(name))
                {
                    continue;
                }

                var matched = wildcard != null
                    ? wildcard.IsMatch(name)
                    : name.Contains(query, comparer) || file.Contains(query, comparer);
                if (!matched)
                {
                    continue;
                }

                results.Add(file);
                if (results.Count >= maxResults)
                {
                    break;
                }
            }

            if (results.Count >= maxResults)
            {
                break;
            }

            IEnumerable<string> directories;
            try
            {
                directories = Directory.EnumerateDirectories(current);
            }
            catch
            {
                continue;
            }

            foreach (var dir in directories)
            {
                pending.Push(dir);
            }
        }

        return results;
    }

    private async Task<StepResult> ExecuteNotifyAsync(PlanStep step, CancellationToken cancellationToken)
    {
        var message = (step.Text ?? step.Target ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(message))
        {
            return new StepResult { Success = false, Message = "Missing notification text" };
        }

        await _auditLog.WriteAsync(new AuditEvent
        {
            Timestamp = DateTimeOffset.UtcNow,
            EventType = "notify",
            Message = message
        }, cancellationToken);

        _logger.LogInformation("Notification: {Message}", message);
        return new StepResult { Success = true, Message = "Notification queued", Data = new { message } };
    }

    private async Task<StepResult> ExecuteMouseJiggleAsync(PlanStep step, CancellationToken cancellationToken)
    {
        var status = await _client.GetStatusAsync(cancellationToken);
        if (!status.Armed)
        {
            return new StepResult { Success = false, Message = "Adapter is disarmed" };
        }

        var duration = step.WaitFor ?? TimeSpan.FromSeconds(30);
        if (duration <= TimeSpan.Zero)
        {
            duration = TimeSpan.FromSeconds(1);
        }

        if (duration > TimeSpan.FromHours(8))
        {
            duration = TimeSpan.FromHours(8);
        }

        if (!TryGetCursorPosition(out var x, out var y, out var initError))
        {
            return new StepResult { Success = false, Message = $"Mouse jiggle not supported: {initError}" };
        }

        var jitterMinMs = 180;
        var jitterMaxMs = 720;
        var maxOffset = 55;
        var started = DateTimeOffset.UtcNow;
        var deadline = started + duration;
        var nextStatusCheckAt = started;
        var statusPollInterval = TimeSpan.FromMilliseconds(250);
        var moves = 0;

        while (DateTimeOffset.UtcNow < deadline)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (_killSwitch.IsTripped)
            {
                return new StepResult { Success = false, Message = $"Kill switch tripped: {_killSwitch.Reason}" };
            }

            if (DateTimeOffset.UtcNow >= nextStatusCheckAt)
            {
                var current = await _client.GetStatusAsync(cancellationToken);
                if (!current.Armed)
                {
                    return new StepResult { Success = false, Message = "Stopped: adapter disarmed during mouse jiggle" };
                }

                nextStatusCheckAt = DateTimeOffset.UtcNow + statusPollInterval;
            }

            var offsetX = Random.Shared.Next(-maxOffset, maxOffset + 1);
            var offsetY = Random.Shared.Next(-maxOffset, maxOffset + 1);
            if (offsetX == 0 && offsetY == 0)
            {
                offsetX = 1;
            }

            var targetX = x + offsetX;
            var targetY = y + offsetY;
            if (!TryMoveCursor(targetX, targetY, out var moveError))
            {
                return new StepResult { Success = false, Message = $"Mouse jiggle failed: {moveError}" };
            }

            x = targetX;
            y = targetY;
            moves++;

            var remaining = deadline - DateTimeOffset.UtcNow;
            if (remaining <= TimeSpan.Zero)
            {
                break;
            }

            var delayMs = Random.Shared.Next(jitterMinMs, jitterMaxMs + 1);
            var sleepMs = (int)Math.Min(delayMs, Math.Max(1, remaining.TotalMilliseconds));
            await Task.Delay(sleepMs, cancellationToken);
        }

        return new StepResult
        {
            Success = true,
            Message = $"Mouse jiggled for {duration.TotalSeconds:0.#} seconds",
            Data = new { moves, seconds = Math.Round(duration.TotalSeconds, 1) }
        };
    }

    private async Task<StepResult> ExecuteVolumeAsync(bool up, bool mute, PlanStep step, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (OperatingSystem.IsWindows())
        {
            if (mute)
            {
                SendWindowsVirtualKey(VkVolumeMute);
                return new StepResult { Success = true, Message = "Volume mute toggled" };
            }

            var times = ParsePositiveInt(step.Text, 1, 1, 50);
            var vk = up ? VkVolumeUp : VkVolumeDown;
            for (var i = 0; i < times; i++)
            {
                SendWindowsVirtualKey(vk);
                await Task.Delay(10, cancellationToken);
            }

            return new StepResult { Success = true, Message = up ? $"Volume up x{times}" : $"Volume down x{times}" };
        }

        if (OperatingSystem.IsLinux())
        {
            if (mute)
            {
                if (TryRunProcess("pactl", "set-sink-mute @DEFAULT_SINK@ toggle", out _, out _)
                    || TryRunProcess("amixer", "set Master toggle", out _, out _))
                {
                    return new StepResult { Success = true, Message = "Volume mute toggled" };
                }

                return new StepResult { Success = false, Message = "Volume control unavailable (install pactl or amixer)" };
            }

            var times = ParsePositiveInt(step.Text, 1, 1, 50);
            var delta = up ? "+5%" : "-5%";
            var ok = true;
            for (var i = 0; i < times; i++)
            {
                if (!(TryRunProcess("pactl", $"set-sink-volume @DEFAULT_SINK@ {delta}", out _, out _)
                      || TryRunProcess("amixer", $"set Master {delta}", out _, out _)))
                {
                    ok = false;
                    break;
                }
            }

            if (!ok)
            {
                return new StepResult { Success = false, Message = "Volume control unavailable (install pactl or amixer)" };
            }

            return new StepResult { Success = true, Message = up ? $"Volume up x{times}" : $"Volume down x{times}" };
        }

        if (OperatingSystem.IsMacOS())
        {
            string script;
            if (mute)
            {
                script = "set volume with output muted";
            }
            else
            {
                var amount = ParsePositiveInt(step.Text, 1, 1, 50) * 5;
                var op = up ? "+" : "-";
                script = $"set volume output volume (output volume of (get volume settings) {op} {amount})";
            }

            if (!TryRunProcess("osascript", $"-e \"{script}\"", out _, out var error))
            {
                return new StepResult { Success = false, Message = error ?? "Volume command failed" };
            }

            return new StepResult { Success = true, Message = mute ? "Volume mute toggled" : up ? "Volume increased" : "Volume decreased" };
        }

        return new StepResult { Success = false, Message = "Volume control not supported on this OS" };
    }

    private Task<StepResult> ExecuteBrightnessAsync(bool up, PlanStep step, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var amount = ParsePositiveInt(step.Text, 10, 1, 100);

        if (OperatingSystem.IsWindows())
        {
            var signed = up ? amount : -amount;
            var command =
                $"$delta={signed};" +
                "$current=(Get-CimInstance -Namespace root/WMI -ClassName WmiMonitorBrightness | Select-Object -First 1).CurrentBrightness;" +
                "$next=[Math]::Max(0,[Math]::Min(100,$current + $delta));" +
                "$method=Get-CimInstance -Namespace root/WMI -ClassName WmiMonitorBrightnessMethods | Select-Object -First 1;" +
                "$null=$method.WmiSetBrightness(1,$next);";
            if (!TryRunProcess("powershell", $"-NoProfile -Command \"{command}\"", out _, out var error))
            {
                return Task.FromResult(new StepResult { Success = false, Message = error ?? "Brightness command failed" });
            }

            return Task.FromResult(new StepResult { Success = true, Message = up ? $"Brightness up {amount}" : $"Brightness down {amount}" });
        }

        if (OperatingSystem.IsLinux())
        {
            var delta = up ? $"+{amount}%" : $"{amount}%-";
            if (!TryRunProcess("brightnessctl", $"set {delta}", out _, out var error))
            {
                return Task.FromResult(new StepResult { Success = false, Message = error ?? "Brightness control unavailable (install brightnessctl)" });
            }

            return Task.FromResult(new StepResult { Success = true, Message = up ? $"Brightness up {amount}" : $"Brightness down {amount}" });
        }

        return Task.FromResult(new StepResult { Success = false, Message = "Brightness control not supported on this OS" });
    }

    private Task<StepResult> ExecuteLockScreenAsync(CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (OperatingSystem.IsWindows())
        {
            if (!TryRunProcess("rundll32.exe", "user32.dll,LockWorkStation", out _, out var error))
            {
                return Task.FromResult(new StepResult { Success = false, Message = error ?? "Lock screen command failed" });
            }

            return Task.FromResult(new StepResult { Success = true, Message = "Screen locked" });
        }

        if (OperatingSystem.IsLinux())
        {
            if (TryRunProcess("loginctl", "lock-session", out _, out _)
                || TryRunProcess("xdg-screensaver", "lock", out _, out _))
            {
                return Task.FromResult(new StepResult { Success = true, Message = "Screen lock requested" });
            }

            return Task.FromResult(new StepResult { Success = false, Message = "Screen lock command unavailable" });
        }

        if (OperatingSystem.IsMacOS())
        {
            var cgSession = "/System/Library/CoreServices/Menu Extras/User.menu/Contents/Resources/CGSession";
            if (!TryRunProcess(cgSession, "-suspend", out _, out var error))
            {
                return Task.FromResult(new StepResult { Success = false, Message = error ?? "Lock screen command failed" });
            }

            return Task.FromResult(new StepResult { Success = true, Message = "Screen lock requested" });
        }

        return Task.FromResult(new StepResult { Success = false, Message = "Lock screen not supported on this OS" });
    }

    private async Task<StepResult> ExecuteActionWithPostCheckAsync(
        PlanStep step,
        Func<Task<ActionResult>> action,
        string successMessage,
        CancellationToken cancellationToken)
    {
        var preWindow = await _client.GetActiveWindowAsync(cancellationToken);
        var selector = BuildEffectiveSelector(step);
        int? preCount = null;
        string? targetRole = null;
        if (!IsSelectorEmpty(selector))
        {
            var preElements = await FindElementsWithRetryAsync(selector, cancellationToken);
            preCount = preElements.Count;
            var ranked = RankElements(preElements, selector, step.Type);
            targetRole = ranked.FirstOrDefault()?.Role;
        }
        var result = await action();
        if (!result.Success)
        {
            return new StepResult { Success = false, Message = result.Message };
        }

        if (IsSelectorEmpty(selector))
        {
            return new StepResult { Success = true, Message = successMessage };
        }

        var post = await VerifyPostConditionAsync(step, selector, preWindow, preCount, targetRole, cancellationToken);
        return BuildPostCheckedResult(result.Success, successMessage, null, post, _config.PostCheckStrict);
    }

    private async Task<ResolvePointResult> ResolveActionPointAsync(PlanStep step, CancellationToken cancellationToken)
    {
        var selector = BuildEffectiveSelector(step);
        var preWindow = await _client.GetActiveWindowAsync(cancellationToken);

        if (step.Point != null)
        {
            var point = step.Point;
            var x = point.Width > 0 ? point.X + point.Width / 2 : point.X;
            var y = point.Height > 0 ? point.Y + point.Height / 2 : point.Y;
            return new ResolvePointResult(true, x, y, selector, preWindow, null, null, null, string.Empty);
        }

        var elements = await FindElementsWithRetryAsync(selector, cancellationToken);
        if (elements.Count > 0)
        {
            var ranked = RankElements(elements, selector, step.Type);
            var target = ranked[0];
            var center = Center(target.Bounds);
            return new ResolvePointResult(
                true,
                center.x,
                center.y,
                selector,
                preWindow,
                elements.Count,
                target.Role,
                target,
                string.Empty);
        }

        if (_config.OcrEnabled && !string.IsNullOrWhiteSpace(selector.NameContains))
        {
            var find = await _contextProvider.FindByTextAsync(selector.NameContains, cancellationToken);
            if (find.OcrMatches.Count > 0)
            {
                var region = SelectBestOcrMatch(find.OcrMatches, selector.NameContains);
                var center = Center(region.Bounds);
                return new ResolvePointResult(
                    true,
                    center.x,
                    center.y,
                    selector,
                    preWindow,
                    null,
                    null,
                    region,
                    string.Empty);
            }
        }

        return new ResolvePointResult(false, 0, 0, selector, preWindow, null, null, null, "No target point found");
    }

    private static int ParsePositiveInt(string? text, int fallback, int min, int max)
    {
        if (!int.TryParse(text, out var value))
        {
            return fallback;
        }

        return Math.Clamp(value, min, max);
    }

    private static (int x, int y) Center(Rect bounds)
    {
        return (bounds.X + bounds.Width / 2, bounds.Y + bounds.Height / 2);
    }

    private static Selector BuildEffectiveSelector(PlanStep step)
    {
        var selector = step.Selector != null ? CloneSelector(step.Selector) : new Selector();
        if (string.IsNullOrWhiteSpace(selector.NameContains) && !string.IsNullOrWhiteSpace(step.Text))
        {
            selector.NameContains = step.Text;
        }
        if (string.IsNullOrWhiteSpace(selector.WindowId) && !string.IsNullOrWhiteSpace(step.ExpectedWindowId))
        {
            selector.WindowId = step.ExpectedWindowId;
        }
        return selector;
    }

    private static Selector CloneSelector(Selector selector)
    {
        var clone = new Selector
        {
            Role = selector.Role,
            NameContains = selector.NameContains,
            AutomationId = selector.AutomationId,
            ClassName = selector.ClassName,
            AncestorNameContains = selector.AncestorNameContains,
            Index = selector.Index,
            WindowId = selector.WindowId
        };

        if (selector.BoundsHint != null)
        {
            clone.BoundsHint = new Rect
            {
                X = selector.BoundsHint.X,
                Y = selector.BoundsHint.Y,
                Width = selector.BoundsHint.Width,
                Height = selector.BoundsHint.Height
            };
        }

        return clone;
    }

    private static bool IsSelectorEmpty(Selector selector)
    {
        return string.IsNullOrWhiteSpace(selector.Role)
               && string.IsNullOrWhiteSpace(selector.NameContains)
               && string.IsNullOrWhiteSpace(selector.AutomationId)
               && string.IsNullOrWhiteSpace(selector.ClassName)
               && string.IsNullOrWhiteSpace(selector.AncestorNameContains)
               && selector.Index <= 0
               && (selector.BoundsHint == null || selector.BoundsHint.Width <= 0 || selector.BoundsHint.Height <= 0);
    }

    private static List<ElementRef> RankElements(IReadOnlyList<ElementRef> elements, Selector selector, ActionType actionType)
    {
        if (elements.Count <= 1)
        {
            return elements.ToList();
        }

        return elements
            .OrderByDescending(element => ScoreElement(element, selector, actionType))
            .ThenByDescending(element => element.Bounds.Width * element.Bounds.Height)
            .ToList();
    }

    private static int ScoreElement(ElementRef element, Selector selector, ActionType actionType)
    {
        var score = 0;
        if (!string.IsNullOrWhiteSpace(selector.Role))
        {
            score += ContainsIgnoreCase(element.Role, selector.Role) ? 4 : -2;
        }

        if (!string.IsNullOrWhiteSpace(selector.NameContains))
        {
            var nameScore = NameMatchScore(element.Name, selector.NameContains);
            score += nameScore;
        }

        if (!string.IsNullOrWhiteSpace(selector.AutomationId))
        {
            score += string.Equals(element.AutomationId, selector.AutomationId, StringComparison.OrdinalIgnoreCase) ? 8 : -3;
        }

        if (!string.IsNullOrWhiteSpace(selector.ClassName))
        {
            score += string.Equals(element.ClassName, selector.ClassName, StringComparison.OrdinalIgnoreCase) ? 3 : -1;
        }

        if (!string.IsNullOrWhiteSpace(selector.AncestorNameContains))
        {
            score += ContainsIgnoreCase(element.PathHints, selector.AncestorNameContains) ? 2 : -1;
        }

        if (selector.BoundsHint != null && selector.BoundsHint.Width > 0 && selector.BoundsHint.Height > 0)
        {
            score += BoundsScore(element.Bounds, selector.BoundsHint);
        }

        score += RoleWeight(actionType, element.Role);

        return score;
    }

    private static List<OcrTextRegion> RankOcrMatches(IReadOnlyList<OcrTextRegion> regions, string query)
    {
        if (regions.Count <= 1)
        {
            return regions.ToList();
        }

        return regions
            .OrderByDescending(region => ScoreOcrRegion(region, query))
            .ToList();
    }

    private static OcrTextRegion SelectBestOcrMatch(IReadOnlyList<OcrTextRegion> regions, string query)
    {
        return RankOcrMatches(regions, query).First();
    }

    private static float ScoreOcrRegion(OcrTextRegion region, string query)
    {
        var score = region.Confidence;
        if (!string.IsNullOrWhiteSpace(query))
        {
            score += NameMatchScore(region.Text, query) / 3f;
        }
        return score;
    }

    private async Task<PostConditionResult> VerifyPostConditionAsync(
        PlanStep step,
        Selector selector,
        WindowRef? preWindow,
        int? preMatchCount,
        string? targetRole,
        CancellationToken cancellationToken)
    {
        if (IsSelectorEmpty(selector))
        {
            return PostConditionResult.CreateSatisfied("No post-check selector");
        }

        var timeoutMs = Math.Clamp(_config.PostCheckTimeoutMs, 100, 5000);
        var pollMs = Math.Clamp(_config.PostCheckPollMs, 20, 1000);
        var deadline = DateTime.UtcNow.AddMilliseconds(timeoutMs);
        PostConditionResult? lastResult = null;
        var lastPostCount = 0;
        var roleExpectation = GetRoleExpectation(targetRole);

        while (true)
        {
            var postWindow = await _client.GetActiveWindowAsync(cancellationToken);
            if (HasWindowChanged(preWindow, postWindow))
            {
                return PostConditionResult.CreateSatisfied("Post-check ok: window changed");
            }

            var elements = await FindElementsWithRetryAsync(selector, cancellationToken);
            lastPostCount = elements.Count;
            var current = EvaluatePostConditionSnapshot(step.Type, roleExpectation, preMatchCount, lastPostCount);
            if (current.Satisfied)
            {
                return current;
            }

            lastResult = current;
            if (DateTime.UtcNow >= deadline)
            {
                break;
            }

            var remainingMs = (int)Math.Ceiling((deadline - DateTime.UtcNow).TotalMilliseconds);
            if (remainingMs <= 0)
            {
                break;
            }

            await Task.Delay(Math.Min(pollMs, remainingMs), cancellationToken);
        }

        if (_config.OcrEnabled && !string.IsNullOrWhiteSpace(selector.NameContains))
        {
            var find = await _contextProvider.FindByTextAsync(selector.NameContains, cancellationToken);
            if (find.OcrMatches.Count > 0)
            {
                if (step.Type is ActionType.Click or ActionType.DoubleClick or ActionType.RightClick or ActionType.Drag or ActionType.Invoke)
                {
                    return PostConditionResult.CreateIndeterminate("Post-check indeterminate: OCR still sees target");
                }

                return PostConditionResult.CreateSatisfied("Post-check ok: OCR matched");
            }
        }

        if (roleExpectation == RoleExpectation.MustRemain && lastPostCount == 0)
        {
            return PostConditionResult.CreateFailed("Post-check failed: checkbox missing");
        }

        if (RequiresPresence(step.Type) && lastPostCount == 0)
        {
            return PostConditionResult.CreateFailed("Post-check failed: target not found");
        }

        return lastResult ?? PostConditionResult.CreateIndeterminate("Post-check indeterminate: target not found");
    }

    private PostConditionResult EvaluatePostConditionSnapshot(
        ActionType stepType,
        RoleExpectation expectation,
        int? preMatchCount,
        int postCount)
    {
        if (stepType is ActionType.Click or ActionType.DoubleClick or ActionType.RightClick or ActionType.Drag or ActionType.Invoke)
        {
            if (expectation == RoleExpectation.WindowChange)
            {
                return PostConditionResult.CreateIndeterminate("Post-check waiting: expected window change for menu item");
            }

            if (expectation == RoleExpectation.MustRemain)
            {
                return postCount > 0
                    ? PostConditionResult.CreateSatisfied("Post-check ok: checkbox still present")
                    : PostConditionResult.CreateIndeterminate("Post-check waiting: checkbox not visible yet");
            }

            if (expectation == RoleExpectation.DisappearOrWindow)
            {
                return postCount == 0
                    ? PostConditionResult.CreateSatisfied("Post-check ok: button disappeared")
                    : PostConditionResult.CreateIndeterminate("Post-check waiting: button still present");
            }

            if (preMatchCount.HasValue && preMatchCount.Value > 0)
            {
                if (postCount == 0)
                {
                    return PostConditionResult.CreateSatisfied("Post-check ok: target disappeared");
                }

                if (postCount != preMatchCount.Value)
                {
                    return PostConditionResult.CreateSatisfied("Post-check ok: target count changed");
                }
            }

            return postCount > 0
                ? PostConditionResult.CreateIndeterminate("Post-check waiting: target still present")
                : PostConditionResult.CreateIndeterminate("Post-check waiting: target not visible yet");
        }

        if (postCount > 0)
        {
            return PostConditionResult.CreateSatisfied("Post-check ok: target found");
        }

        if (RequiresPresence(stepType))
        {
            return PostConditionResult.CreateIndeterminate("Post-check waiting: target not visible yet");
        }

        return PostConditionResult.CreateIndeterminate("Post-check indeterminate: target not found");
    }

    private static bool RequiresPresence(ActionType type)
    {
        return type is ActionType.TypeText or ActionType.SetValue or ActionType.WaitForText;
    }

    private static StepResult BuildPostCheckedResult(bool success, string message, object? data, PostConditionResult post, bool strict)
    {
        if (!post.Satisfied)
        {
            if (post.Indeterminate && !strict)
            {
                return new StepResult { Success = success, Message = $"{message}. {post.Message} (lenient)", Data = data };
            }

            return new StepResult { Success = false, Message = $"{message}. {post.Message}", Data = data };
        }

        return new StepResult { Success = success, Message = $"{message}. {post.Message}", Data = data };
    }

    private static bool ContainsIgnoreCase(string? value, string fragment)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return false;
        }
        return value.Contains(fragment, StringComparison.OrdinalIgnoreCase);
    }

    private static int NameMatchScore(string candidate, string query)
    {
        if (string.IsNullOrWhiteSpace(candidate) || string.IsNullOrWhiteSpace(query))
        {
            return 0;
        }

        if (string.Equals(candidate, query, StringComparison.OrdinalIgnoreCase))
        {
            return 8;
        }

        if (candidate.StartsWith(query, StringComparison.OrdinalIgnoreCase))
        {
            return 6;
        }

        if (ContainsIgnoreCase(candidate, query))
        {
            return 4;
        }

        return -1;
    }

    private static int BoundsScore(Rect element, Rect hint)
    {
        var score = 0;
        if (Intersects(element, hint))
        {
            score += 3;
        }

        var elementCenter = Center(element);
        if (Contains(hint, elementCenter))
        {
            score += 3;
        }

        var hintCenter = Center(hint);
        var dx = elementCenter.x - hintCenter.x;
        var dy = elementCenter.y - hintCenter.y;
        var distance = Math.Sqrt(dx * dx + dy * dy);
        if (distance < 50) score += 3;
        else if (distance < 150) score += 2;
        else if (distance < 300) score += 1;

        return score;
    }

    private static int RoleWeight(ActionType actionType, string role)
    {
        if (string.IsNullOrWhiteSpace(role))
        {
            return 0;
        }

        var normalized = role.ToLowerInvariant();
        return actionType switch
        {
            ActionType.Click or ActionType.DoubleClick or ActionType.RightClick or ActionType.Drag or ActionType.Invoke => normalized.Contains("button") ? 4 :
                normalized.Contains("menuitem") ? 3 :
                normalized.Contains("link") ? 2 :
                normalized.Contains("checkbox") ? 2 :
                normalized.Contains("radio") ? 2 : 0,
            ActionType.TypeText or ActionType.SetValue => normalized.Contains("edit") ? 4 :
                normalized.Contains("textbox") ? 4 :
                normalized.Contains("text") ? 2 : 0,
            _ => 0
        };
    }

    private enum RoleExpectation
    {
        None,
        WindowChange,
        MustRemain,
        DisappearOrWindow
    }

    private RoleExpectation GetRoleExpectation(string? role)
    {
        if (string.IsNullOrWhiteSpace(role))
        {
            return RoleExpectation.None;
        }

        var normalized = role.ToLowerInvariant();
        if (normalized.Contains("menuitem") || normalized.Contains("menu item"))
        {
            return ParseRule(_config.PostCheckRules.MenuItem);
        }

        if (normalized.Contains("checkbox") || normalized.Contains("check box"))
        {
            return ParseRule(_config.PostCheckRules.Checkbox);
        }

        if (normalized.Contains("button") || normalized.Contains("pushbutton"))
        {
            return ParseRule(_config.PostCheckRules.Button);
        }

        return RoleExpectation.None;
    }

    private static RoleExpectation ParseRule(string? rule)
    {
        if (string.IsNullOrWhiteSpace(rule))
        {
            return RoleExpectation.None;
        }

        return rule.Trim().ToLowerInvariant() switch
        {
            "window-change" => RoleExpectation.WindowChange,
            "present" => RoleExpectation.MustRemain,
            "disappear-or-window" => RoleExpectation.DisappearOrWindow,
            "none" => RoleExpectation.None,
            _ => RoleExpectation.None
        };
    }

    private static bool HasWindowChanged(WindowRef? before, WindowRef? after)
    {
        if (before == null || after == null)
        {
            return false;
        }

        if (!string.Equals(before.Id, after.Id, StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        if (!string.Equals(before.Title, after.Title, StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        if (!string.Equals(before.AppId, after.AppId, StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        return false;
    }

    private void EnsureDefaultBinding(PlanStep step, WindowRef? activeWindow, RuntimeBinding binding)
    {
        if (!_config.ContextBindingEnabled || step.Type == ActionType.OpenApp || !RequiresUiContext(step.Type))
        {
            return;
        }

        if (activeWindow == null)
        {
            return;
        }

        if (string.IsNullOrWhiteSpace(binding.AppId))
        {
            binding.AppId = activeWindow.AppId;
        }

        if (string.IsNullOrWhiteSpace(binding.WindowId))
        {
            binding.WindowId = activeWindow.Id;
        }
    }

    private PolicyDecision EvaluateContextBinding(PlanStep step, WindowRef? activeWindow, RuntimeBinding binding)
    {
        if (!_config.ContextBindingEnabled || step.Type == ActionType.OpenApp || !RequiresUiContext(step.Type))
        {
            return new PolicyDecision { Allowed = true, RequiresConfirmation = false, Reason = "Context binding disabled or not applicable" };
        }

        var expectedApp = FirstNonEmpty(step.ExpectedAppId, binding.AppId);
        var expectedWindow = FirstNonEmpty(step.ExpectedWindowId, _config.ContextBindingRequireWindow ? binding.WindowId : null);
        if (string.IsNullOrWhiteSpace(expectedApp) && string.IsNullOrWhiteSpace(expectedWindow))
        {
            return new PolicyDecision { Allowed = true, RequiresConfirmation = false, Reason = "No binding constraints" };
        }

        if (activeWindow == null)
        {
            return new PolicyDecision
            {
                Allowed = false,
                RequiresConfirmation = false,
                Reason = "Blocked: no active window for context binding"
            };
        }

        if (!string.IsNullOrWhiteSpace(expectedApp) && !AppIdMatches(activeWindow.AppId, expectedApp))
        {
            return new PolicyDecision
            {
                Allowed = false,
                RequiresConfirmation = false,
                Reason = $"Blocked by context binding: expected app '{expectedApp}', got '{activeWindow.AppId}'."
            };
        }

        if (!string.IsNullOrWhiteSpace(expectedWindow)
            && !string.Equals(activeWindow.Id, expectedWindow, StringComparison.OrdinalIgnoreCase))
        {
            return new PolicyDecision
            {
                Allowed = false,
                RequiresConfirmation = false,
                Reason = $"Blocked by context binding: expected window '{expectedWindow}', got '{activeWindow.Id}'."
            };
        }

        return new PolicyDecision { Allowed = true, RequiresConfirmation = false, Reason = "Context binding matched" };
    }

    private async Task RefreshBindingAfterStepAsync(PlanStep step, RuntimeBinding binding, CancellationToken cancellationToken)
    {
        if (!_config.ContextBindingEnabled || (!RequiresUiContext(step.Type) && step.Type != ActionType.OpenApp))
        {
            return;
        }

        var activeWindow = await _client.GetActiveWindowAsync(cancellationToken);
        if (activeWindow == null)
        {
            return;
        }

        if (step.Type == ActionType.OpenApp)
        {
            binding.AppId = activeWindow.AppId;
            binding.WindowId = activeWindow.Id;
            return;
        }

        var expectedApp = FirstNonEmpty(step.ExpectedAppId, binding.AppId);
        if (!string.IsNullOrWhiteSpace(expectedApp) && !AppIdMatches(activeWindow.AppId, expectedApp))
        {
            return;
        }

        if (!string.IsNullOrWhiteSpace(activeWindow.AppId))
        {
            binding.AppId = activeWindow.AppId;
        }

        if (!string.IsNullOrWhiteSpace(activeWindow.Id))
        {
            binding.WindowId = activeWindow.Id;
        }
    }

    private static bool RequiresUiContext(ActionType type)
    {
        return type is ActionType.Find
            or ActionType.Click
            or ActionType.DoubleClick
            or ActionType.RightClick
            or ActionType.Drag
            or ActionType.TypeText
            or ActionType.KeyCombo
            or ActionType.Invoke
            or ActionType.SetValue
            or ActionType.ReadText
            or ActionType.CaptureScreen
            or ActionType.WaitForText;
    }

    private static bool AppIdMatches(string? actual, string? expected)
    {
        if (string.IsNullOrWhiteSpace(actual) || string.IsNullOrWhiteSpace(expected))
        {
            return false;
        }

        var actualNormalized = NormalizeAppToken(actual);
        var expectedNormalized = NormalizeAppToken(expected);
        if (string.Equals(actualNormalized, expectedNormalized, StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        if (actualNormalized.Contains(expectedNormalized, StringComparison.OrdinalIgnoreCase)
            || expectedNormalized.Contains(actualNormalized, StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        return false;
    }

    private static string NormalizeAppToken(string value)
    {
        var trimmed = value.Trim().Trim('"', '\'');
        if (string.IsNullOrWhiteSpace(trimmed))
        {
            return string.Empty;
        }

        if (trimmed.Contains(Path.DirectorySeparatorChar)
            || trimmed.Contains(Path.AltDirectorySeparatorChar))
        {
            var fileName = Path.GetFileNameWithoutExtension(trimmed);
            if (!string.IsNullOrWhiteSpace(fileName))
            {
                trimmed = fileName;
            }
        }

        return trimmed.Replace(" ", string.Empty).ToLowerInvariant();
    }

    private static string? FirstNonEmpty(params string?[] values)
    {
        foreach (var value in values)
        {
            if (!string.IsNullOrWhiteSpace(value))
            {
                return value;
            }
        }

        return null;
    }

    private async Task<IReadOnlyList<ElementRef>> FindElementsWithRetryAsync(Selector selector, CancellationToken cancellationToken)
    {
        var attempts = Math.Clamp(_config.FindRetryCount, 0, 20) + 1;
        var delayMs = Math.Clamp(_config.FindRetryDelayMs, 0, 2000);
        var selectorVariants = BuildSelectorVariants(selector);
        IReadOnlyList<ElementRef> elements = Array.Empty<ElementRef>();

        for (var attempt = 0; attempt < attempts; attempt++)
        {
            foreach (var variant in selectorVariants)
            {
                elements = await _client.FindElementsAsync(variant, cancellationToken);
                if (elements.Count > 0)
                {
                    return elements;
                }
            }

            if (attempt < attempts - 1 && delayMs > 0)
            {
                await Task.Delay(delayMs, cancellationToken);
            }
        }

        return elements;
    }

    private static IReadOnlyList<Selector> BuildSelectorVariants(Selector selector)
    {
        var variants = new List<Selector> { CloneSelector(selector) };

        // Self-healing path: if automation id changed, degrade to stable fields.
        if (!string.IsNullOrWhiteSpace(selector.AutomationId))
        {
            var withoutAutomationId = CloneSelector(selector);
            withoutAutomationId.AutomationId = string.Empty;
            variants.Add(withoutAutomationId);

            if (!string.IsNullOrWhiteSpace(withoutAutomationId.ClassName))
            {
                var withoutClass = CloneSelector(withoutAutomationId);
                withoutClass.ClassName = string.Empty;
                variants.Add(withoutClass);
            }

            if (!string.IsNullOrWhiteSpace(withoutAutomationId.AncestorNameContains))
            {
                var withoutAncestor = CloneSelector(withoutAutomationId);
                withoutAncestor.AncestorNameContains = string.Empty;
                variants.Add(withoutAncestor);
            }
        }

        if (!string.IsNullOrWhiteSpace(selector.NameContains) && selector.NameContains.Contains(' '))
        {
            var tokens = selector.NameContains.Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            foreach (var token in tokens)
            {
                if (token.Length < 3)
                {
                    continue;
                }

                var byToken = CloneSelector(selector);
                byToken.NameContains = token;
                variants.Add(byToken);
                break;
            }
        }

        return variants
            .GroupBy(SelectorCacheKey, StringComparer.OrdinalIgnoreCase)
            .Select(group => group.First())
            .ToList();
    }

    private static string SelectorCacheKey(Selector selector)
    {
        var bounds = selector.BoundsHint;
        return string.Join("|",
            selector.Role?.Trim() ?? string.Empty,
            selector.NameContains?.Trim() ?? string.Empty,
            selector.AutomationId?.Trim() ?? string.Empty,
            selector.ClassName?.Trim() ?? string.Empty,
            selector.AncestorNameContains?.Trim() ?? string.Empty,
            selector.Index.ToString(),
            selector.WindowId?.Trim() ?? string.Empty,
            bounds?.X.ToString() ?? string.Empty,
            bounds?.Y.ToString() ?? string.Empty,
            bounds?.Width.ToString() ?? string.Empty,
            bounds?.Height.ToString() ?? string.Empty);
    }

    private static bool Intersects(Rect a, Rect b)
    {
        return !(a.X + a.Width < b.X || b.X + b.Width < a.X || a.Y + a.Height < b.Y || b.Y + b.Height < a.Y);
    }

    private static bool Contains(Rect rect, (int x, int y) point)
    {
        return point.x >= rect.X && point.x <= rect.X + rect.Width && point.y >= rect.Y && point.y <= rect.Y + rect.Height;
    }

    private bool TryResolveAllowedPath(string rawPath, out string fullPath, out string reason)
    {
        fullPath = string.Empty;
        reason = string.Empty;
        try
        {
            fullPath = Path.GetFullPath(rawPath.Trim().Trim('"', '\''), AppContext.BaseDirectory);
        }
        catch (Exception ex)
        {
            reason = $"Invalid path: {ex.Message}";
            return false;
        }

        var roots = _config.FilesystemAllowedRoots;
        if (roots.Count == 0)
        {
            return true;
        }

        foreach (var root in roots)
        {
            if (string.IsNullOrWhiteSpace(root))
            {
                continue;
            }

            string resolvedRoot;
            try
            {
                resolvedRoot = Path.GetFullPath(root.Trim().Trim('"', '\''), AppContext.BaseDirectory);
            }
            catch
            {
                continue;
            }

            if (fullPath.StartsWith(resolvedRoot, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        reason = $"Path not allowed by FilesystemAllowedRoots: {fullPath}";
        return false;
    }

    private static bool TryNormalizeUrl(string raw, out string normalized)
    {
        normalized = raw.Trim();
        if (Uri.TryCreate(normalized, UriKind.Absolute, out var absolute) && (absolute.Scheme is "http" or "https"))
        {
            normalized = absolute.ToString();
            return true;
        }

        var withScheme = $"https://{normalized}";
        if (Uri.TryCreate(withScheme, UriKind.Absolute, out var implicitHttps) && !string.IsNullOrWhiteSpace(implicitHttps.Host))
        {
            normalized = implicitHttps.ToString();
            return true;
        }

        return false;
    }

    private static bool TryGetCursorPosition(out int x, out int y, out string? error)
    {
        x = 0;
        y = 0;
        error = null;

        if (OperatingSystem.IsWindows())
        {
            if (GetCursorPos(out var pt))
            {
                x = pt.X;
                y = pt.Y;
                return true;
            }

            error = "GetCursorPos failed";
            return false;
        }

        if (OperatingSystem.IsLinux())
        {
            if (!TryRunProcess("xdotool", "getmouselocation --shell", out var output, out error))
            {
                return false;
            }

            var lines = output.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            var xLine = lines.FirstOrDefault(line => line.StartsWith("X=", StringComparison.OrdinalIgnoreCase));
            var yLine = lines.FirstOrDefault(line => line.StartsWith("Y=", StringComparison.OrdinalIgnoreCase));
            if (xLine == null || yLine == null)
            {
                error = "xdotool output parse failed";
                return false;
            }

            if (!int.TryParse(xLine[2..], out x) || !int.TryParse(yLine[2..], out y))
            {
                error = "xdotool coordinates parse failed";
                return false;
            }

            return true;
        }

        error = "Unsupported OS for mouse jiggle";
        return false;
    }

    private static bool TryMoveCursor(int x, int y, out string? error)
    {
        error = null;
        if (OperatingSystem.IsWindows())
        {
            if (!SetCursorPos(x, y))
            {
                error = "SetCursorPos failed";
                return false;
            }

            return true;
        }

        if (OperatingSystem.IsLinux())
        {
            return TryRunProcess("xdotool", $"mousemove {x} {y}", out _, out error);
        }

        error = "Unsupported OS for mouse jiggle";
        return false;
    }

    private static void SendWindowsVirtualKey(byte virtualKey)
    {
        if (!OperatingSystem.IsWindows())
        {
            return;
        }

        KeybdEvent(virtualKey, 0, 0, UIntPtr.Zero);
        KeybdEvent(virtualKey, 0, KeyEventKeyUp, UIntPtr.Zero);
    }

    private static bool TryRunProcess(string fileName, string arguments, out string output, out string? error)
    {
        output = string.Empty;
        error = null;
        try
        {
            using var process = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = fileName,
                    Arguments = arguments,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                }
            };

            process.Start();
            output = process.StandardOutput.ReadToEnd();
            var stdErr = process.StandardError.ReadToEnd();
            process.WaitForExit();
            if (process.ExitCode != 0)
            {
                error = string.IsNullOrWhiteSpace(stdErr) ? $"{fileName} exited with code {process.ExitCode}" : stdErr.Trim();
                return false;
            }

            return true;
        }
        catch (Exception ex)
        {
            error = ex.Message;
            return false;
        }
    }

    private static async Task<(bool Success, string? Error)> RunProcessAsync(string fileName, string arguments, TimeSpan timeout, CancellationToken cancellationToken)
    {
        try
        {
            using var process = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = fileName,
                    Arguments = arguments,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                }
            };

            process.Start();
            var stdOutTask = process.StandardOutput.ReadToEndAsync();
            var stdErrTask = process.StandardError.ReadToEndAsync();

            using var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            timeoutCts.CancelAfter(timeout);

            try
            {
                await process.WaitForExitAsync(timeoutCts.Token);
            }
            catch (OperationCanceledException)
            {
                try
                {
                    if (!process.HasExited)
                    {
                        process.Kill(entireProcessTree: true);
                    }
                }
                catch
                {
                    // best-effort kill
                }

                return (false, cancellationToken.IsCancellationRequested
                    ? "Recording canceled"
                    : "Recording timed out");
            }

            var stdOut = await stdOutTask;
            var stdErr = await stdErrTask;
            if (process.ExitCode == 0)
            {
                return (true, null);
            }

            var error = string.IsNullOrWhiteSpace(stdErr) ? stdOut : stdErr;
            if (string.IsNullOrWhiteSpace(error))
            {
                error = $"{fileName} exited with code {process.ExitCode}";
            }
            return (false, error.Trim());
        }
        catch (Exception ex)
        {
            return (false, ex.Message);
        }
    }

    private string BuildMediaOutputPath(string? requested, string prefix, string extension)
    {
        var root = ResolveMediaRoot();
        Directory.CreateDirectory(root);

        var fileName = Path.GetFileName((requested ?? string.Empty).Trim().Trim('"', '\''));
        if (string.IsNullOrWhiteSpace(fileName))
        {
            fileName = $"{prefix}-{DateTime.UtcNow:yyyyMMdd-HHmmss}-{Guid.NewGuid():N}{extension}";
        }

        if (!Path.HasExtension(fileName))
        {
            fileName += extension;
        }

        var safeChars = fileName
            .Select(ch => Path.GetInvalidFileNameChars().Contains(ch) ? '_' : ch)
            .ToArray();
        var safeName = new string(safeChars);
        if (!safeName.EndsWith(extension, StringComparison.OrdinalIgnoreCase))
        {
            safeName = Path.GetFileNameWithoutExtension(safeName) + extension;
        }

        return Path.Combine(root, safeName);
    }

    private string ResolveMediaRoot()
    {
        var storageRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "DesktopAgent");
        Directory.CreateDirectory(storageRoot);

        var configured = (_config.MediaOutputDirectory ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(configured))
        {
            configured = "media";
        }

        var root = Path.IsPathRooted(configured)
            ? configured
            : Path.Combine(storageRoot, configured);

        return Path.GetFullPath(root);
    }

    private static string ResolveFfmpegPath()
    {
        var ffmpegPath = Environment.GetEnvironmentVariable("DESKTOP_AGENT_FFMPEG_PATH");
        return string.IsNullOrWhiteSpace(ffmpegPath) ? "ffmpeg" : ffmpegPath;
    }

    private string? BuildRecordingArguments(string ffmpegPath, string outputPath, int seconds, FfmpegAudioInput? audioInput, out string? error)
    {
        error = null;
        var fps = 12;
        var inputArguments = BuildInputArguments(ffmpegPath, fps, audioInput, out error);
        if (string.IsNullOrWhiteSpace(inputArguments))
        {
            return null;
        }

        if (audioInput.HasValue)
        {
            return $"-y {inputArguments} -t {seconds} -c:v libx264 -preset veryfast -pix_fmt yuv420p -c:a aac \"{outputPath}\"";
        }

        return $"-y {inputArguments} -t {seconds} -c:v libx264 -preset veryfast -pix_fmt yuv420p \"{outputPath}\"";
    }

    private string? BuildLiveRecordingArguments(string ffmpegPath, string outputPath, FfmpegAudioInput? audioInput, out string? error)
    {
        error = null;
        var fps = 12;
        var inputArguments = BuildInputArguments(ffmpegPath, fps, audioInput, out error);
        if (string.IsNullOrWhiteSpace(inputArguments))
        {
            return null;
        }

        if (audioInput.HasValue)
        {
            return $"-y {inputArguments} -c:v libx264 -preset veryfast -pix_fmt yuv420p -c:a aac \"{outputPath}\"";
        }

        return $"-y {inputArguments} -c:v libx264 -preset veryfast -pix_fmt yuv420p \"{outputPath}\"";
    }

    private string? BuildInputArguments(string ffmpegPath, int fps, FfmpegAudioInput? audioInput, out string? error)
    {
        if (!TryBuildVideoInputArguments(ffmpegPath, fps, out var videoInput, out error))
        {
            return null;
        }

        if (audioInput is { CombinedInputArguments: { } combined } && !string.IsNullOrWhiteSpace(combined))
        {
            return combined;
        }

        if (audioInput is { InputArguments: { } input } && !string.IsNullOrWhiteSpace(input))
        {
            return $"{videoInput} {input}";
        }

        return videoInput;
    }

    private bool TryBuildVideoInputArguments(string ffmpegPath, int fps, out string videoInput, out string? error)
    {
        videoInput = string.Empty;
        error = null;

        if (OperatingSystem.IsWindows())
        {
            if (!_config.ScreenRecordingPrimaryDisplayOnly)
            {
                videoInput = $"-f gdigrab -framerate {fps} -i desktop";
                return true;
            }

            if (TryGetPrimaryDisplaySize(out var width, out var height))
            {
                videoInput = $"-f gdigrab -framerate {fps} -offset_x 0 -offset_y 0 -video_size {width}x{height} -i desktop";
                return true;
            }

            _logger.LogInformation("Primary display capture requested but size detection failed. Falling back to full desktop capture.");
            videoInput = $"-f gdigrab -framerate {fps} -i desktop";
            return true;
        }

        if (OperatingSystem.IsMacOS())
        {
            if (!FfmpegSupportsInput(ffmpegPath, "avfoundation"))
            {
                error = "ffmpeg build has no avfoundation input support.";
                return false;
            }

            var videoIndex = ResolveAvFoundationVideoDeviceIndex(ffmpegPath);
            videoInput = $"-f avfoundation -framerate {fps} -capture_cursor 1 -capture_mouse_clicks 1 -i \"{videoIndex}:none\"";
            return true;
        }

        if (OperatingSystem.IsLinux())
        {
            var display = (Environment.GetEnvironmentVariable("DISPLAY") ?? string.Empty).Trim();
            if (!string.IsNullOrWhiteSpace(display))
            {
                videoInput = $"-f x11grab -framerate {fps} -i {display}+0,0";
                return true;
            }

            if (IsWaylandSession() && FfmpegSupportsInput(ffmpegPath, "pipewire"))
            {
                var node = (Environment.GetEnvironmentVariable("DESKTOP_AGENT_PIPEWIRE_NODE") ?? "0").Trim();
                if (string.IsNullOrWhiteSpace(node))
                {
                    node = "0";
                }

                videoInput = $"-f pipewire -framerate {fps} -i {node}";
                return true;
            }

            error = IsWaylandSession()
                ? "No X11 display detected. For Wayland recording set DESKTOP_AGENT_PIPEWIRE_NODE and use an ffmpeg build with pipewire input."
                : "No X11 display detected (DISPLAY is empty).";
            return false;
        }

        error = "Unsupported OS for screen recording.";
        return false;
    }

    private static bool IsWaylandSession()
    {
        var wayland = (Environment.GetEnvironmentVariable("WAYLAND_DISPLAY") ?? string.Empty).Trim();
        return !string.IsNullOrWhiteSpace(wayland);
    }

    private static string ResolveAvFoundationVideoDeviceIndex(string ffmpegPath)
    {
        if (TryListAvFoundationDevices(ffmpegPath, out var videoDevices, out _))
        {
            var captureScreen = videoDevices
                .FirstOrDefault(device => device.Name.Contains("capture screen", StringComparison.OrdinalIgnoreCase));
            if (captureScreen.Index.Length > 0)
            {
                return captureScreen.Index;
            }

            if (videoDevices.Count > 0 && videoDevices[0].Index.Length > 0)
            {
                return videoDevices[0].Index;
            }
        }

        return "1";
    }

    private static bool TryGetPrimaryDisplaySize(out int width, out int height)
    {
        width = 0;
        height = 0;
        try
        {
            var w = GetSystemMetrics(SystemMetricCxScreen);
            var h = GetSystemMetrics(SystemMetricCyScreen);
            if (w > 0 && h > 0)
            {
                width = w;
                height = h;
                return true;
            }
        }
        catch
        {
            // ignored
        }

        return false;
    }

    private bool TryStartRecordingProcess(
        string ffmpegPath,
        string outputPath,
        bool includeAudioRequested,
        out bool audioUsed,
        out string? audioBackend,
        out string? error)
    {
        audioUsed = false;
        audioBackend = null;
        error = null;

        if (!includeAudioRequested)
        {
            if (TryStartRecordingProcessCore(ffmpegPath, outputPath, audioInput: null, out error))
            {
                return true;
            }

            return false;
        }

        var audioInputs = ResolveAvailableAudioInputs(ffmpegPath);
        if (audioInputs.Count == 0)
        {
            error = GetNoSupportedAudioBackendMessage();
        }
        else
        {
            var failures = new List<string>();
            foreach (var input in audioInputs)
            {
                if (TryStartRecordingProcessCore(ffmpegPath, outputPath, input, out error))
                {
                    audioUsed = true;
                    audioBackend = input.DisplayName;
                    return true;
                }

                failures.Add($"{input.DisplayName}: {TrimForMessage(error)}");
                _logger.LogWarning("Live recording with audio backend {Backend} failed. Error: {Error}", input.DisplayName, error);
            }

            error = string.Join(" | ", failures);
        }

        _logger.LogWarning("Live recording audio failed, retrying video-only. Error: {Error}", error);
        if (TryStartRecordingProcessCore(ffmpegPath, outputPath, audioInput: null, out var videoOnlyError))
        {
            audioUsed = false;
            audioBackend = null;
            error = string.IsNullOrWhiteSpace(error) ? null : error;
            return true;
        }

        error = string.IsNullOrWhiteSpace(error)
            ? videoOnlyError
            : $"{error} | video-only: {videoOnlyError}";
        return false;
    }

    private bool TryStartRecordingProcessCore(string ffmpegPath, string outputPath, FfmpegAudioInput? audioInput, out string? error)
    {
        error = null;
        try
        {
            var arguments = BuildLiveRecordingArguments(ffmpegPath, outputPath, audioInput, out error);
            if (string.IsNullOrWhiteSpace(arguments))
            {
                return false;
            }

            var process = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = ffmpegPath,
                    Arguments = arguments,
                    UseShellExecute = false,
                    CreateNoWindow = true
                }
            };

            process.Start();
            Thread.Sleep(800);
            if (process.HasExited)
            {
                var code = process.ExitCode;
                process.Dispose();
                error = $"ffmpeg exited immediately with code {code}";
                return false;
            }

            lock (_recordingSync)
            {
                _recordingProcess = process;
                _recordingOutputPath = outputPath;
                _recordingStartedAtUtc = DateTimeOffset.UtcNow;
                _recordingAudio = audioInput.HasValue;
            }

            return true;
        }
        catch (Exception ex)
        {
            error = ex.Message;
            return false;
        }
    }

    private List<FfmpegAudioInput> ResolveAvailableAudioInputs(string ffmpegPath)
    {
        if (OperatingSystem.IsWindows())
        {
            return ResolveWindowsAudioInputs(ffmpegPath);
        }

        if (OperatingSystem.IsMacOS())
        {
            return ResolveMacAudioInputs(ffmpegPath);
        }

        if (OperatingSystem.IsLinux())
        {
            return ResolveLinuxAudioInputs(ffmpegPath);
        }

        return new List<FfmpegAudioInput>();
    }

    private List<FfmpegAudioInput> ResolveWindowsAudioInputs(string ffmpegPath)
    {
        var inputs = new List<FfmpegAudioInput>();
        var pref = (_config.ScreenRecordingAudioBackendPreference ?? "auto").Trim().ToLowerInvariant();
        var preferredDevice = (_config.ScreenRecordingAudioDevice ?? string.Empty).Trim();
        var wasapiSupported = FfmpegSupportsInput(ffmpegPath, "wasapi");
        var dshowSupported = FfmpegSupportsInput(ffmpegPath, "dshow");

        FfmpegAudioInput? dshowInput = null;
        FfmpegAudioInput? dshowAutoInput = null;
        if (dshowSupported)
        {
            if (!string.IsNullOrWhiteSpace(preferredDevice))
            {
                var escapedPref = EscapeFfmpegDshowDevice(preferredDevice);
                dshowInput = new FfmpegAudioInput($"DirectShow ({preferredDevice})", $"-f dshow -i audio=\"{escapedPref}\"");
                if (TryGetFirstDshowAudioDevice(ffmpegPath, out var fallbackDetected)
                    && !string.Equals(fallbackDetected, preferredDevice, StringComparison.OrdinalIgnoreCase))
                {
                    var escapedFallback = EscapeFfmpegDshowDevice(fallbackDetected);
                    dshowAutoInput = new FfmpegAudioInput($"DirectShow ({fallbackDetected})", $"-f dshow -i audio=\"{escapedFallback}\"");
                }
            }
            else if (TryGetFirstDshowAudioDevice(ffmpegPath, out var detected))
            {
                var escaped = EscapeFfmpegDshowDevice(detected);
                dshowInput = new FfmpegAudioInput($"DirectShow ({detected})", $"-f dshow -i audio=\"{escaped}\"");
            }
        }

        AddInPreferredOrder(inputs, pref,
            ("wasapi", wasapiSupported ? new FfmpegAudioInput("WASAPI", "-f wasapi -i default") : null),
            ("dshow", dshowInput),
            ("dshow:auto", dshowAutoInput));

        if (!wasapiSupported)
        {
            _logger.LogInformation("ffmpeg input format wasapi not available.");
        }
        if (dshowSupported && dshowInput == null)
        {
            _logger.LogInformation("ffmpeg dshow detected but no audio device found.");
        }

        return inputs
            .Where(static input => !string.IsNullOrWhiteSpace(input.InputArguments) || !string.IsNullOrWhiteSpace(input.CombinedInputArguments))
            .DistinctBy(static input => $"{input.InputArguments}|{input.CombinedInputArguments}")
            .ToList();
    }

    private List<FfmpegAudioInput> ResolveMacAudioInputs(string ffmpegPath)
    {
        var inputs = new List<FfmpegAudioInput>();
        if (!FfmpegSupportsInput(ffmpegPath, "avfoundation"))
        {
            return inputs;
        }

        var pref = (_config.ScreenRecordingAudioBackendPreference ?? "auto").Trim().ToLowerInvariant();
        if (pref != "auto" && pref != "avfoundation")
        {
            return inputs;
        }

        var preferredDevice = (_config.ScreenRecordingAudioDevice ?? string.Empty).Trim();
        var videoIndex = ResolveAvFoundationVideoDeviceIndex(ffmpegPath);
        var audioDevice = ResolveAvFoundationAudioDevice(ffmpegPath, preferredDevice);
        if (!audioDevice.HasValue)
        {
            _logger.LogInformation("ffmpeg avfoundation detected but no audio device found.");
            return inputs;
        }

        var audioToken = audioDevice.Value.Index;
        var displayName = $"AVFoundation ({audioDevice.Value.Name})";
        var combinedInput = $"-f avfoundation -framerate 12 -capture_cursor 1 -capture_mouse_clicks 1 -i \"{videoIndex}:{audioToken}\"";
        inputs.Add(new FfmpegAudioInput(displayName, InputArguments: null, CombinedInputArguments: combinedInput));
        return inputs;
    }

    private List<FfmpegAudioInput> ResolveLinuxAudioInputs(string ffmpegPath)
    {
        var inputs = new List<FfmpegAudioInput>();
        var pref = (_config.ScreenRecordingAudioBackendPreference ?? "auto").Trim().ToLowerInvariant();
        var preferredDevice = (_config.ScreenRecordingAudioDevice ?? string.Empty).Trim();

        var hasPipewire = FfmpegSupportsInput(ffmpegPath, "pipewire");
        var hasPulse = FfmpegSupportsInput(ffmpegPath, "pulse");
        var hasAlsa = FfmpegSupportsInput(ffmpegPath, "alsa");

        FfmpegAudioInput? pipewireInput = null;
        if (hasPipewire)
        {
            var node = string.IsNullOrWhiteSpace(preferredDevice)
                ? (Environment.GetEnvironmentVariable("DESKTOP_AGENT_PIPEWIRE_AUDIO_NODE") ?? "0").Trim()
                : preferredDevice;
            if (string.IsNullOrWhiteSpace(node))
            {
                node = "0";
            }

            pipewireInput = new FfmpegAudioInput("PipeWire", $"-f pipewire -i {node}");
        }

        FfmpegAudioInput? pulseInput = null;
        if (hasPulse)
        {
            var pulseDevice = string.IsNullOrWhiteSpace(preferredDevice) ? "default" : preferredDevice;
            pulseInput = new FfmpegAudioInput($"PulseAudio ({pulseDevice})", $"-f pulse -i {EscapeFfmpegSimpleToken(pulseDevice)}");
        }

        FfmpegAudioInput? alsaInput = null;
        if (hasAlsa)
        {
            var alsaDevice = string.IsNullOrWhiteSpace(preferredDevice) ? "default" : preferredDevice;
            alsaInput = new FfmpegAudioInput($"ALSA ({alsaDevice})", $"-f alsa -i {EscapeFfmpegSimpleToken(alsaDevice)}");
        }

        AddInPreferredOrder(inputs, pref,
            ("pipewire", pipewireInput),
            ("pulse", pulseInput),
            ("alsa", alsaInput));

        return inputs
            .Where(static input => !string.IsNullOrWhiteSpace(input.InputArguments) || !string.IsNullOrWhiteSpace(input.CombinedInputArguments))
            .DistinctBy(static input => $"{input.InputArguments}|{input.CombinedInputArguments}")
            .ToList();
    }

    private static void AddInPreferredOrder(List<FfmpegAudioInput> target, string preference, params (string Key, FfmpegAudioInput? Input)[] candidates)
    {
        var ordered = new List<(string Key, FfmpegAudioInput? Input)>();
        if (!string.IsNullOrWhiteSpace(preference) && !string.Equals(preference, "auto", StringComparison.OrdinalIgnoreCase))
        {
            var prioritized = candidates.FirstOrDefault(candidate => string.Equals(candidate.Key, preference, StringComparison.OrdinalIgnoreCase));
            if (prioritized.Input.HasValue)
            {
                ordered.Add(prioritized);
            }
        }

        foreach (var candidate in candidates)
        {
            if (!candidate.Input.HasValue)
            {
                continue;
            }

            if (ordered.Any(existing => string.Equals(existing.Key, candidate.Key, StringComparison.OrdinalIgnoreCase)))
            {
                continue;
            }

            ordered.Add(candidate);
        }

        foreach (var item in ordered)
        {
            if (item.Input.HasValue)
            {
                target.Add(item.Input.Value);
            }
        }
    }

    private string GetNoSupportedAudioBackendMessage()
    {
        if (OperatingSystem.IsWindows())
        {
            return "No supported audio backend in ffmpeg (WASAPI/DirectShow missing).";
        }

        if (OperatingSystem.IsMacOS())
        {
            return "No supported audio backend in ffmpeg (AVFoundation missing or no audio device).";
        }

        if (OperatingSystem.IsLinux())
        {
            return "No supported audio backend in ffmpeg (PipeWire/PulseAudio/ALSA missing).";
        }

        return "No supported audio backend in ffmpeg.";
    }

    private static bool FfmpegSupportsInput(string ffmpegPath, string inputName)
    {
        if (!TryRunProcess(ffmpegPath, "-hide_banner -devices", out var output, out var error))
        {
            output = string.Join(Environment.NewLine, output, error);
        }

        return Regex.IsMatch(output ?? string.Empty, $@"\b{Regex.Escape(inputName)}\b", RegexOptions.IgnoreCase);
    }

    private static bool TryGetFirstDshowAudioDevice(string ffmpegPath, out string deviceName)
    {
        deviceName = string.Empty;
        var audioDump = string.Empty;
        TryRunProcess(ffmpegPath, "-hide_banner -f dshow -list_devices true -i dummy", out _, out var error);
        if (!string.IsNullOrWhiteSpace(error))
        {
            audioDump = error!;
        }

        if (string.IsNullOrWhiteSpace(audioDump))
        {
            return false;
        }

        var directAudioCandidates = new List<string>();
        var sectionAudioCandidates = new List<string>();
        var inAudioSection = false;
        foreach (var rawLine in audioDump.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
        {
            var line = rawLine.Trim();
            if (line.Contains("Alternative name", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (line.Contains("DirectShow audio devices", StringComparison.OrdinalIgnoreCase))
            {
                inAudioSection = true;
                continue;
            }

            if (line.Contains("DirectShow video devices", StringComparison.OrdinalIgnoreCase))
            {
                inAudioSection = false;
            }

            var match = Regex.Match(line, "\"([^\"]+)\"");
            if (!match.Success)
            {
                continue;
            }

            var name = match.Groups[1].Value.Trim();
            if (string.IsNullOrWhiteSpace(name))
            {
                continue;
            }

            if (line.Contains("(audio)", StringComparison.OrdinalIgnoreCase))
            {
                directAudioCandidates.Add(name);
            }
            else if (inAudioSection)
            {
                sectionAudioCandidates.Add(name);
            }
        }

        var candidate = directAudioCandidates
            .Concat(sectionAudioCandidates)
            .FirstOrDefault(static x => !string.IsNullOrWhiteSpace(x));

        if (!string.IsNullOrWhiteSpace(candidate))
        {
            deviceName = candidate;
            return true;
        }

        return false;
    }

    private static (string Index, string Name)? ResolveAvFoundationAudioDevice(string ffmpegPath, string preferredDevice)
    {
        if (!TryListAvFoundationDevices(ffmpegPath, out _, out var audioDevices))
        {
            return null;
        }

        if (!string.IsNullOrWhiteSpace(preferredDevice))
        {
            var byIndex = audioDevices.FirstOrDefault(device => string.Equals(device.Index, preferredDevice, StringComparison.OrdinalIgnoreCase));
            if (byIndex.Index.Length > 0)
            {
                return byIndex;
            }

            var byName = audioDevices.FirstOrDefault(device => string.Equals(device.Name, preferredDevice, StringComparison.OrdinalIgnoreCase));
            if (byName.Index.Length > 0)
            {
                return byName;
            }
        }

        if (audioDevices.Count == 0)
        {
            return null;
        }

        return audioDevices[0];
    }

    private static bool TryListAvFoundationDevices(
        string ffmpegPath,
        out List<(string Index, string Name)> videoDevices,
        out List<(string Index, string Name)> audioDevices)
    {
        videoDevices = new List<(string Index, string Name)>();
        audioDevices = new List<(string Index, string Name)>();

        TryRunProcess(ffmpegPath, "-hide_banner -f avfoundation -list_devices true -i \"\"", out _, out var error);
        var dump = error ?? string.Empty;
        if (string.IsNullOrWhiteSpace(dump))
        {
            return false;
        }

        var inVideoSection = false;
        var inAudioSection = false;
        foreach (var raw in dump.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
        {
            var line = raw.Trim();
            if (line.Contains("AVFoundation video devices", StringComparison.OrdinalIgnoreCase))
            {
                inVideoSection = true;
                inAudioSection = false;
                continue;
            }

            if (line.Contains("AVFoundation audio devices", StringComparison.OrdinalIgnoreCase))
            {
                inVideoSection = false;
                inAudioSection = true;
                continue;
            }

            var match = Regex.Match(line, @"\[(\d+)\]\s+(.+)$");
            if (!match.Success)
            {
                continue;
            }

            var index = match.Groups[1].Value.Trim();
            var name = match.Groups[2].Value.Trim();
            if (string.IsNullOrWhiteSpace(index) || string.IsNullOrWhiteSpace(name))
            {
                continue;
            }

            if (inVideoSection)
            {
                videoDevices.Add((index, name));
            }
            else if (inAudioSection)
            {
                audioDevices.Add((index, name));
            }
        }

        return videoDevices.Count > 0 || audioDevices.Count > 0;
    }

    private static string EscapeFfmpegDshowDevice(string value)
        => value.Replace("\"", "\\\"", StringComparison.Ordinal);

    private static string EscapeFfmpegSimpleToken(string value)
    {
        var token = value.Trim();
        if (token.Length == 0)
        {
            return "default";
        }

        if (!token.Contains(' ') && !token.Contains('"'))
        {
            return token;
        }

        return $"\"{token.Replace("\"", "\\\"", StringComparison.Ordinal)}\"";
    }

    private static string TrimForMessage(string? value)
    {
        var text = (value ?? string.Empty)
            .Replace("\r", " ", StringComparison.Ordinal)
            .Replace("\n", " ", StringComparison.Ordinal)
            .Trim();

        if (text.Length <= 180)
        {
            return text;
        }

        return text[..180] + "...";
    }

    private readonly record struct FfmpegAudioInput(string DisplayName, string? InputArguments, string? CombinedInputArguments = null);

    private void AddClipboardHistory(string source, string? text)
    {
        var value = text ?? string.Empty;
        lock (ClipboardHistoryLock)
        {
            ClipboardHistory.Add(new ClipboardHistoryEntry
            {
                Timestamp = DateTimeOffset.UtcNow,
                Source = source,
                Value = value
            });

            var maxItems = Math.Clamp(_config.ClipboardHistoryMaxItems, 1, 500);
            if (ClipboardHistory.Count <= maxItems)
            {
                return;
            }

            var excess = ClipboardHistory.Count - maxItems;
            ClipboardHistory.RemoveRange(0, excess);
        }
    }

    private IReadOnlyList<object> GetClipboardHistory()
    {
        lock (ClipboardHistoryLock)
        {
            return ClipboardHistory
                .OrderByDescending(item => item.Timestamp)
                .Select(item => (object)new
                {
                    item.Timestamp,
                    item.Source,
                    item.Value
                })
                .ToList();
        }
    }

    private static async Task<StepResult> ExecuteSimpleResultAsync(Func<Task<ActionResult>> action, string successMessage)
    {
        var result = await action();
        return new StepResult { Success = result.Success, Message = result.Success ? successMessage : result.Message };
    }

    private async Task WriteAuditAsync(string eventType, string message, PlanStep step, CancellationToken cancellationToken)
    {
        await _auditLog.WriteAsync(new AuditEvent
        {
            Timestamp = DateTimeOffset.UtcNow,
            EventType = eventType,
            Message = message,
            Data = new
            {
                step.Type,
                step.Text,
                step.Target,
                step.AppIdOrPath,
                step.ExpectedAppId,
                step.ExpectedWindowId
            }
        }, cancellationToken);
    }

    private sealed class RuntimeBinding
    {
        public string? AppId { get; set; }
        public string? WindowId { get; set; }
    }

    private sealed record ResolvePointResult(
        bool Success,
        int X,
        int Y,
        Selector Selector,
        WindowRef? PreWindow,
        int? PreCount,
        string? TargetRole,
        object? Data,
        string Error);

    private sealed class ClipboardHistoryEntry
    {
        public DateTimeOffset Timestamp { get; set; }
        public string Source { get; set; } = string.Empty;
        public string Value { get; set; } = string.Empty;
    }

    private sealed class PostConditionResult
    {
        public bool Satisfied { get; init; }
        public bool Indeterminate { get; init; }
        public string Message { get; init; } = string.Empty;

        public static PostConditionResult CreateSatisfied(string message) => new() { Satisfied = true, Message = message };
        public static PostConditionResult CreateFailed(string message) => new() { Satisfied = false, Indeterminate = false, Message = message };
        public static PostConditionResult CreateIndeterminate(string message) => new() { Satisfied = false, Indeterminate = true, Message = message };
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct NativePoint
    {
        public int X;
        public int Y;
    }

    private const uint MouseEventLeftDown = 0x0002;
    private const uint MouseEventLeftUp = 0x0004;
    private const uint MouseEventRightDown = 0x0008;
    private const uint MouseEventRightUp = 0x0010;
    private const uint KeyEventKeyUp = 0x0002;
    private const byte VkVolumeMute = 0xAD;
    private const byte VkVolumeDown = 0xAE;
    private const byte VkVolumeUp = 0xAF;
    private const int SystemMetricCxScreen = 0;
    private const int SystemMetricCyScreen = 1;

    [DllImport("user32.dll")]
    private static extern bool GetCursorPos(out NativePoint lpPoint);

    [DllImport("user32.dll")]
    private static extern bool SetCursorPos(int x, int y);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern void MouseEvent(uint dwFlags, uint dx, uint dy, uint dwData, UIntPtr dwExtraInfo);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern void KeybdEvent(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

    [DllImport("user32.dll")]
    private static extern int GetSystemMetrics(int nIndex);
}
