using System.Diagnostics;
using System.Net;
using System.Net.Http;
using System.Net.Http.Json;
using System.Reflection;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using DesktopAgent.Core.Abstractions;
using DesktopAgent.Core.Config;
using DesktopAgent.Core.Models;
using DesktopAgent.Core.Services;
using DesktopAgent.Proto;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace DesktopAgent.Tray;

internal sealed class TrayLocalAgent : IDisposable
{
    private const string SupportedCommandsHelp = "Available commands: status, kill, reset kill, lock status, lock on <current window|app>, unlock, profile <safe|balanced|power>, arm, disarm, simulate presence, require presence, list apps [query] [allowed], goals, goal add <text>, goal run <id>, goal done <id>, goal remove <id>, goal priority <id> <low|normal|high>, goal auto <id> <on|off>, goal scheduler on|off|every <sec>, continue goal, memory, run <intent>, dry-run <intent>, translate <text> to <language> (or 'translate to <language>: <text>'), order intake [<email text>], order preview, order clear, order fill <url>. Plugin intents: file write/read/list/append/search, take screenshot [for each screen|single screen], record screen [and audio] for <duration>, start recording [screen] [with/without audio], stop recording, jiggle mouse for <duration>.";
    private const int GoalPriorityLow = 0;
    private const int GoalPriorityNormal = 1;
    private const int GoalPriorityHigh = 2;

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true,
        WriteIndented = true
    };

    private readonly object _sync = new();
    private readonly AgentConfig _config;
    private readonly string _configPath;
    private readonly string _storageRoot;
    private readonly string _goalLibraryPath;
    private readonly string _memoryLibraryPath;
    private readonly string _version;
    private readonly TimeSpan _httpTimeout;
    private readonly ILoggerFactory _loggerFactory;
    private readonly DesktopGrpcClient _desktopClient;
    private readonly IAuditLog _auditLog;
    private readonly IContextProvider _contextProvider;
    private readonly IAppResolver _appResolver;
    private readonly IPlanner _planner;
    private readonly IPolicyEngine _policyEngine;
    private readonly IRateLimiter _rateLimiter;
    private readonly IKillSwitch _killSwitch;
    private IOcrEngine _ocrEngine;
    private bool _ocrRestartRequired;
    private readonly Dictionary<string, PendingAction> _pendingActions = new(StringComparer.OrdinalIgnoreCase);
    private ClarificationRequest? _pendingClarification;
    private readonly List<WebTaskItem> _tasks = new();
    private readonly List<ScheduleState> _schedules = new();
    private readonly List<GoalState> _goals = new();
    private readonly List<IntentMemoryEntry> _intentMemory = new();
    private readonly CancellationTokenSource _schedulerCts = new();
    private readonly Task _schedulerTask;

    private ContextLockState _contextLock = ContextLockState.None;
    private DateTimeOffset _lastGoalSchedulerSweepUtc = DateTimeOffset.MinValue;
    private bool _awaitingOrderEmailInput;
    private OrderDraft? _latestOrderDraft;

    public TrayLocalAgent(string adapterEndpoint, TimeSpan timeout, string? configPath)
    {
        AppContext.SetSwitch("System.Net.Http.SocketsHttpHandler.Http2UnencryptedSupport", true);

        _storageRoot = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "DesktopAgent");
        Directory.CreateDirectory(_storageRoot);
        _goalLibraryPath = Path.Combine(_storageRoot, "goals.library.json");
        _memoryLibraryPath = Path.Combine(_storageRoot, "memory.library.json");

        _configPath = ResolveConfigPath(configPath, _storageRoot);
        _config = LoadConfig(_configPath, adapterEndpoint, _storageRoot);
        _version = ResolveAppVersion();
        _httpTimeout = TimeSpan.FromSeconds(Math.Clamp((int)Math.Ceiling(timeout.TotalSeconds), 2, 15));

        _loggerFactory = LoggerFactory.Create(builder => builder.SetMinimumLevel(LogLevel.Warning));
        _desktopClient = new DesktopGrpcClient(_config.AdapterEndpoint, _loggerFactory.CreateLogger<DesktopGrpcClient>());
        _auditLog = new JsonlAuditLog(_config);
        _killSwitch = new KillSwitch();
        _ocrEngine = OcrEngineFactory.Create(_config, _loggerFactory);
        _contextProvider = new ContextProvider(_desktopClient, _ocrEngine, _config, _loggerFactory.CreateLogger<ContextProvider>());
        _appResolver = new AppResolver(_config, new LocalAppCatalog(_config));
        _planner = new SimplePlanner(new FallbackIntentInterpreter(
            new RuleBasedIntentInterpreter(_appResolver, _config),
            new LocalLlmIntentRewriter(_config, _auditLog),
            _auditLog,
            _config));
        _policyEngine = new PolicyEngine(_config);
        _rateLimiter = new SlidingWindowRateLimiter(() => _config.MaxActionsPerSecond);

        LoadTasksFromDisk();
        LoadSchedulesFromDisk();
        LoadGoalsFromDisk();
        LoadMemoryFromDisk();
        _schedulerTask = Task.Run(() => ScheduleLoopAsync(_schedulerCts.Token));
    }

    private static string ResolveAppVersion()
    {
        var assembly = typeof(TrayLocalAgent).Assembly;
        var informational = assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion;
        if (!string.IsNullOrWhiteSpace(informational))
        {
            return NormalizeVersionForDisplay(informational);
        }

        var fileVersion = assembly.GetCustomAttribute<AssemblyFileVersionAttribute>()?.Version;
        if (!string.IsNullOrWhiteSpace(fileVersion))
        {
            return NormalizeVersionForDisplay(fileVersion);
        }

        var fallback = assembly.GetName().Version?.ToString() ?? "unknown";
        return NormalizeVersionForDisplay(fallback);
    }

    private static string NormalizeVersionForDisplay(string raw)
    {
        var value = (raw ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(value))
        {
            return "unknown";
        }

        var plusIndex = value.IndexOf('+');
        if (plusIndex > 0)
        {
            value = value[..plusIndex];
        }

        var semverCore = Regex.Match(value, @"^\d+\.\d+\.\d+");
        if (semverCore.Success)
        {
            return semverCore.Value;
        }

        if (Version.TryParse(value, out var parsed))
        {
            return $"{parsed.Major}.{parsed.Minor}.{Math.Max(0, parsed.Build)}";
        }

        return value;
    }

    public async Task<WebChatResponse> SendChatAsync(string message, CancellationToken cancellationToken)
    {
        var text = message?.Trim() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(text))
        {
            return WebChatResponse.Simple("Enter a valid command.");
        }

        var normalized = text.ToLowerInvariant();

        if (TryHandlePendingClarificationReply(text, out var clarificationResponse, out var clarificationMessage))
        {
            if (clarificationResponse == null)
            {
                return WebChatResponse.Simple(clarificationMessage ?? "Please answer with a number (e.g. 1) or 'cancel'.");
            }

            return await BuildPlanConfirmationAsync(clarificationResponse.Intent, dryRun: false, cancellationToken);
        }
        _pendingClarification = null;

        var orderResponse = await TryHandleOrderIntakeAsync(text, normalized, cancellationToken);
        if (orderResponse != null)
        {
            return orderResponse;
        }

        if (normalized is "kill" or "kill switch" or "panic" or "panic stop" or "stop now" or "abort")
        {
            _killSwitch.Trip("Kill requested from chat");
            await WriteAuditAsync("kill", "Kill switch tripped by user", cancellationToken);
            return WebChatResponse.Simple("Kill switch enabled. Running actions will stop immediately.");
        }

        if (normalized is "reset kill" or "clear kill" or "unkill")
        {
            _killSwitch.Reset();
            await WriteAuditAsync("kill_reset", "Kill switch reset by user", cancellationToken);
            return WebChatResponse.Simple("Kill switch reset.");
        }

        if (normalized is "lock status" or "context lock status")
        {
            return WebChatResponse.Simple(FormatContextLock());
        }

        if (normalized is "goals" or "goal list" or "list goals")
        {
            return WebChatResponse.Simple(FormatGoals());
        }

        if (normalized.StartsWith("goal add ", StringComparison.Ordinal))
        {
            var goalText = text["goal add ".Length..].Trim();
            return await AddGoalAsync(goalText, cancellationToken);
        }

        if (normalized.StartsWith("set goal ", StringComparison.Ordinal))
        {
            var goalText = text["set goal ".Length..].Trim();
            return await AddGoalAsync(goalText, cancellationToken);
        }

        if (normalized.StartsWith("goal done ", StringComparison.Ordinal))
        {
            var key = text["goal done ".Length..].Trim();
            return await MarkGoalDoneAsync(key, cancellationToken);
        }

        if (normalized.StartsWith("goal remove ", StringComparison.Ordinal))
        {
            var key = text["goal remove ".Length..].Trim();
            return await RemoveGoalAsync(key, cancellationToken);
        }

        if (normalized.StartsWith("goal run ", StringComparison.Ordinal))
        {
            var key = text["goal run ".Length..].Trim();
            return await BuildGoalPlanConfirmationAsync(key, cancellationToken);
        }

        if (normalized.StartsWith("goal priority ", StringComparison.Ordinal))
        {
            var payload = text["goal priority ".Length..].Trim();
            return await SetGoalPriorityAsync(payload, cancellationToken);
        }

        if (normalized.StartsWith("goal auto ", StringComparison.Ordinal))
        {
            var payload = text["goal auto ".Length..].Trim();
            return await SetGoalAutoModeAsync(payload, cancellationToken);
        }

        if (normalized is "goal scheduler on" or "goals scheduler on")
        {
            _config.GoalSchedulerEnabled = true;
            await SaveConfigToDiskAsync(cancellationToken);
            return WebChatResponse.Simple("Goal scheduler enabled.");
        }

        if (normalized is "goal scheduler off" or "goals scheduler off")
        {
            _config.GoalSchedulerEnabled = false;
            await SaveConfigToDiskAsync(cancellationToken);
            return WebChatResponse.Simple("Goal scheduler disabled.");
        }

        if (normalized.StartsWith("goal scheduler every ", StringComparison.Ordinal))
        {
            var raw = text["goal scheduler every ".Length..].Trim();
            if (int.TryParse(raw.Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).FirstOrDefault(), out var seconds))
            {
                _config.GoalSchedulerIntervalSeconds = Math.Clamp(seconds, 10, 3600);
                await SaveConfigToDiskAsync(cancellationToken);
                return WebChatResponse.Simple($"Goal scheduler interval set to {_config.GoalSchedulerIntervalSeconds}s.");
            }

            return WebChatResponse.Simple("Use: goal scheduler every <seconds>");
        }

        if (normalized is "continue goal" or "continue goals" or "next goal")
        {
            return await ContinueLatestGoalAsync(cancellationToken);
        }

        if (normalized is "memory" or "agent memory")
        {
            return WebChatResponse.Simple(FormatMemory());
        }

        if (normalized is "memory clear" or "clear memory")
        {
            lock (_sync)
            {
                _intentMemory.Clear();
            }

            await SaveMemoryToDiskAsync(cancellationToken);
            return WebChatResponse.Simple("Memory cleared.");
        }

        if (normalized is "unlock" or "unlock context" or "unlock context lock" or "release lock")
        {
            lock (_sync)
            {
                _contextLock = ContextLockState.None;
            }

            await WriteAuditAsync("context_unlock", "Context lock disabled", cancellationToken);
            return WebChatResponse.Simple("Context lock disabled.");
        }

        if (normalized is "lock current" or "lock current window" or "lock on current" or "lock on current window")
        {
            return await LockCurrentWindowAsync(cancellationToken);
        }

        if (normalized.StartsWith("lock on ", StringComparison.Ordinal))
        {
            var target = text["lock on ".Length..].Trim();
            return await LockTargetAsync(target, cancellationToken);
        }

        if (normalized.Contains("status", StringComparison.Ordinal))
        {
            var status = await _desktopClient.GetStatusAsync(cancellationToken);
            var killStatus = _killSwitch.IsTripped
                ? $"Kill switch: ON ({_killSwitch.Reason ?? "manual"})."
                : "Kill switch: OFF.";
            return WebChatResponse.Simple($"Adapter armed: {status.Armed}, require presence: {status.RequireUserPresence}. {killStatus} {FormatContextLock()}");
        }

        if (normalized.StartsWith("profile ", StringComparison.Ordinal))
        {
            var profile = AgentProfileService.NormalizeProfile(text["profile ".Length..].Trim());
            _config.ActiveProfile = profile;
            _config.ProfileModeEnabled = true;
            AgentProfileService.ApplyActiveProfile(_config);
            await SaveConfigToDiskAsync(cancellationToken);
            return WebChatResponse.Simple($"Profile set to {profile}. Limits/policy updated.");
        }

        if (normalized.StartsWith("arm", StringComparison.Ordinal))
        {
            _killSwitch.Reset();
            var status = await _desktopClient.ArmAsync(requireUserPresence: true, cancellationToken);
            await WriteAuditAsync("arm", "Adapter armed", cancellationToken);
            return WebChatResponse.Simple($"Armed: {status.Armed}, require presence: {status.RequireUserPresence}. Kill switch reset.");
        }

        if (normalized.StartsWith("disarm", StringComparison.Ordinal))
        {
            _killSwitch.Trip("Disarm requested from chat");
            var status = await _desktopClient.DisarmAsync(cancellationToken);
            await WriteAuditAsync("disarm", "Adapter disarmed", cancellationToken);
            return WebChatResponse.Simple($"Armed: {status.Armed}. Running actions stopped.");
        }

        if (normalized.Contains("simulate presence", StringComparison.Ordinal))
        {
            var token = CreatePendingAction(PendingActionType.SimulatePresence, text, null, dryRun: false);
            return WebChatResponse.Confirm("Simulate presence? This will arm the adapter without user presence.", token);
        }

        if (normalized.Contains("require presence", StringComparison.Ordinal))
        {
            _killSwitch.Reset();
            var status = await _desktopClient.ArmAsync(requireUserPresence: true, cancellationToken);
            await WriteAuditAsync("presence_required", "Require presence enabled", cancellationToken);
            return WebChatResponse.Simple($"Presence required. Armed: {status.Armed}. Kill switch reset.");
        }
        if (normalized.StartsWith("list apps", StringComparison.Ordinal) || normalized.StartsWith("apps", StringComparison.Ordinal))
        {
            return ListApps(text, normalized);
        }

        if (normalized.StartsWith("run ", StringComparison.Ordinal))
        {
            var intent = text["run ".Length..].Trim();
            return await BuildPlanConfirmationAsync(intent, dryRun: false, cancellationToken);
        }

        if (normalized.StartsWith("dry-run ", StringComparison.Ordinal) || normalized.StartsWith("dryrun ", StringComparison.Ordinal))
        {
            var intent = text.Contains(' ') ? text[(text.IndexOf(' ') + 1)..].Trim() : string.Empty;
            var result = await ExecuteIntentAsync(intent, dryRun: true, cancellationToken);
            return result == null ? WebChatResponse.Error("Execution failed.") : ToChatResponse(result);
        }

        if (normalized.StartsWith("intent ", StringComparison.Ordinal))
        {
            var intent = text["intent ".Length..].Trim();
            return await BuildPlanConfirmationAsync(intent, dryRun: false, cancellationToken);
        }

        if (TryParseTranslationIntent(text, out var translationIntent))
        {
            return await TranslateWithLlmAsync(translationIntent, cancellationToken);
        }

        if (IsDirectIntent(normalized))
        {
            return await BuildPlanConfirmationAsync(text, dryRun: false, cancellationToken);
        }

        var plan = _planner.PlanFromIntent(text);
        if (IsUnrecognizedPlan(plan))
        {
            return await BuildUnknownCommandResponseAsync(text, cancellationToken);
        }

        if (TryCreateOpenAppClarification(plan, out var clarification))
        {
            return clarification;
        }

        var pendingToken = CreatePendingAction(PendingActionType.ExecutePlan, text, plan, dryRun: false);
        var notice = GetRewriteNotice(plan);
        var prompt = string.IsNullOrWhiteSpace(notice)
            ? "I interpreted your request. Confirm execution?"
            : $"I interpreted your request. {notice}. Confirm execution?";
        return WebChatResponse.ConfirmWithSteps(prompt, pendingToken, PlanToLines(plan), PlanToJson(plan), GetModeLabel(plan));
    }

    public async Task<WebChatResponse> ConfirmAsync(string token, bool approve, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(token))
        {
            return WebChatResponse.Error("Token missing.");
        }

        var pending = TakePendingAction(token);
        if (pending == null)
        {
            return WebChatResponse.Error("Token not found.");
        }

        if (!approve)
        {
            return WebChatResponse.Simple("Cancelled.");
        }

        switch (pending.Type)
        {
            case PendingActionType.SimulatePresence:
                _killSwitch.Reset();
                var status = await _desktopClient.ArmAsync(requireUserPresence: false, cancellationToken);
                await WriteAuditAsync("simulate_presence", "Presence simulation enabled", cancellationToken);
                return WebChatResponse.Simple($"Presence simulated. Armed: {status.Armed}, require presence: {status.RequireUserPresence}. Kill switch reset.");

            case PendingActionType.ExecutePlan:
                if (pending.Plan == null)
                {
                    return WebChatResponse.Error("Plan missing.");
                }

                var execution = await ExecutePlanInternalAsync(pending.Source, pending.Plan, pending.DryRun, approvedByUser: true, cancellationToken);
                return ToChatResponse(execution);

            default:
                return WebChatResponse.Error("Action not supported.");
        }
    }

    public async Task<WebStatusResponse?> GetStatusAsync(CancellationToken cancellationToken)
    {
        var status = await _desktopClient.GetStatusAsync(cancellationToken);
        var llm = await GetLlmStatusAsync(cancellationToken);
        return new WebStatusResponse(
            Version: _version,
            Adapter: new WebAdapterStatus(status.Armed, status.RequireUserPresence, status.Message),
            Llm: llm,
            KillSwitch: new WebKillSwitchStatus(_killSwitch.IsTripped, _killSwitch.Reason));
    }

    public Task<WebConfigResponse?> GetConfigAsync(CancellationToken cancellationToken)
    {
        var response = new WebConfigResponse(
            ProfileModeEnabled: _config.ProfileModeEnabled,
            ActiveProfile: _config.ActiveProfile,
            LlmInterpretationMode: _config.LlmInterpretationMode,
            RequireConfirmation: _config.RequireConfirmation,
            MaxActionsPerSecond: _config.MaxActionsPerSecond,
            QuizSafeModeEnabled: _config.QuizSafeModeEnabled,
            OcrEnabled: _config.OcrEnabled,
            OcrRestartRequired: _ocrRestartRequired,
            MediaOutputDirectory: _config.MediaOutputDirectory,
            ScreenRecordingAudioBackendPreference: _config.ScreenRecordingAudioBackendPreference,
            ScreenRecordingAudioDevice: _config.ScreenRecordingAudioDevice,
            ScreenRecordingPrimaryDisplayOnly: _config.ScreenRecordingPrimaryDisplayOnly,
            AdapterRestartCommand: _config.AdapterRestartCommand,
            AdapterRestartWorkingDir: _config.AdapterRestartWorkingDir,
            FindRetryCount: _config.FindRetryCount,
            FindRetryDelayMs: _config.FindRetryDelayMs,
            PostCheckTimeoutMs: _config.PostCheckTimeoutMs,
            PostCheckPollMs: _config.PostCheckPollMs,
            ClipboardHistoryMaxItems: _config.ClipboardHistoryMaxItems,
            FilesystemAllowedRoots: _config.FilesystemAllowedRoots,
            AllowedApps: _config.AllowedApps,
            AppAliases: _config.AppAliases,
            Llm: new WebConfigLlm(
                Enabled: _config.LlmFallbackEnabled,
                AllowNonLoopbackEndpoint: _config.AllowNonLoopbackLlmEndpoint,
                Provider: _config.LlmFallback.Provider,
                Endpoint: _config.LlmFallback.Endpoint,
                Model: _config.LlmFallback.Model,
                TimeoutSeconds: _config.LlmFallback.TimeoutSeconds,
                MaxTokens: _config.LlmFallback.MaxTokens),
            AuditLlmInteractions: _config.AuditLlmInteractions,
            AuditLlmIncludeRawText: _config.AuditLlmIncludeRawText);
        return Task.FromResult<WebConfigResponse?>(response);
    }

    public async Task<WebConfigResponse?> SaveConfigAsync(WebConfigUpdate payload, CancellationToken cancellationToken)
    {
        if (payload.ProfileModeEnabled.HasValue)
        {
            _config.ProfileModeEnabled = payload.ProfileModeEnabled.Value;
        }

        if (!string.IsNullOrWhiteSpace(payload.ActiveProfile))
        {
            _config.ActiveProfile = AgentProfileService.NormalizeProfile(payload.ActiveProfile);
        }

        if (payload.RequireConfirmation.HasValue)
        {
            _config.RequireConfirmation = payload.RequireConfirmation.Value;
        }

        if (payload.MaxActionsPerSecond.HasValue)
        {
            _config.MaxActionsPerSecond = Math.Max(1, payload.MaxActionsPerSecond.Value);
        }

        if (payload.QuizSafeModeEnabled.HasValue)
        {
            _config.QuizSafeModeEnabled = payload.QuizSafeModeEnabled.Value;
        }

        if (payload.OcrEnabled.HasValue && payload.OcrEnabled.Value != _config.OcrEnabled)
        {
            _config.OcrEnabled = payload.OcrEnabled.Value;
            _ocrRestartRequired = true;
        }

        if (payload.MediaOutputDirectory != null)
        {
            _config.MediaOutputDirectory = payload.MediaOutputDirectory.Trim();
        }

        if (payload.ScreenRecordingAudioBackendPreference != null)
        {
            _config.ScreenRecordingAudioBackendPreference = payload.ScreenRecordingAudioBackendPreference.Trim();
        }

        if (payload.ScreenRecordingAudioDevice != null)
        {
            _config.ScreenRecordingAudioDevice = payload.ScreenRecordingAudioDevice.Trim();
        }

        if (payload.ScreenRecordingPrimaryDisplayOnly.HasValue)
        {
            _config.ScreenRecordingPrimaryDisplayOnly = payload.ScreenRecordingPrimaryDisplayOnly.Value;
        }

        if (payload.AdapterRestartCommand != null)
        {
            _config.AdapterRestartCommand = payload.AdapterRestartCommand.Trim();
        }

        if (payload.AdapterRestartWorkingDir != null)
        {
            _config.AdapterRestartWorkingDir = payload.AdapterRestartWorkingDir.Trim();
        }

        if (payload.FindRetryCount.HasValue)
        {
            _config.FindRetryCount = Math.Max(0, payload.FindRetryCount.Value);
        }

        if (payload.FindRetryDelayMs.HasValue)
        {
            _config.FindRetryDelayMs = Math.Max(0, payload.FindRetryDelayMs.Value);
        }

        if (payload.PostCheckTimeoutMs.HasValue)
        {
            _config.PostCheckTimeoutMs = Math.Max(0, payload.PostCheckTimeoutMs.Value);
        }

        if (payload.PostCheckPollMs.HasValue)
        {
            _config.PostCheckPollMs = Math.Max(1, payload.PostCheckPollMs.Value);
        }
        if (payload.ClipboardHistoryMaxItems.HasValue)
        {
            _config.ClipboardHistoryMaxItems = Math.Max(1, payload.ClipboardHistoryMaxItems.Value);
        }

        if (payload.FilesystemAllowedRoots != null)
        {
            _config.FilesystemAllowedRoots = payload.FilesystemAllowedRoots
                .Select(root => root?.Trim())
                .Where(root => !string.IsNullOrWhiteSpace(root))
                .Cast<string>()
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        if (payload.AllowedApps != null)
        {
            _config.AllowedApps = payload.AllowedApps
                .Select(app => app?.Trim())
                .Where(app => !string.IsNullOrWhiteSpace(app))
                .Cast<string>()
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        if (payload.AppAliases != null)
        {
            _config.AppAliases = new Dictionary<string, string>(
                payload.AppAliases
                    .Where(kv => !string.IsNullOrWhiteSpace(kv.Key) && !string.IsNullOrWhiteSpace(kv.Value))
                    .ToDictionary(kv => kv.Key.Trim(), kv => kv.Value.Trim(), StringComparer.OrdinalIgnoreCase),
                StringComparer.OrdinalIgnoreCase);
        }

        if (payload.AuditLlmInteractions.HasValue)
        {
            _config.AuditLlmInteractions = payload.AuditLlmInteractions.Value;
        }

        if (payload.AuditLlmIncludeRawText.HasValue)
        {
            _config.AuditLlmIncludeRawText = payload.AuditLlmIncludeRawText.Value;
        }

        if (!string.IsNullOrWhiteSpace(payload.LlmInterpretationMode))
        {
            _config.LlmInterpretationMode = payload.LlmInterpretationMode.Trim();
        }

        if (payload.Llm != null)
        {
            if (payload.Llm.Enabled.HasValue)
            {
                _config.LlmFallbackEnabled = payload.Llm.Enabled.Value;
            }

            if (payload.Llm.AllowNonLoopbackEndpoint.HasValue)
            {
                _config.AllowNonLoopbackLlmEndpoint = payload.Llm.AllowNonLoopbackEndpoint.Value;
            }

            if (!string.IsNullOrWhiteSpace(payload.Llm.Provider))
            {
                _config.LlmFallback.Provider = payload.Llm.Provider.Trim();
            }

            if (payload.Llm.Endpoint != null)
            {
                _config.LlmFallback.Endpoint = payload.Llm.Endpoint.Trim();
            }

            if (payload.Llm.Model != null)
            {
                _config.LlmFallback.Model = payload.Llm.Model.Trim();
            }

            if (payload.Llm.TimeoutSeconds.HasValue)
            {
                _config.LlmFallback.TimeoutSeconds = Math.Max(1, payload.Llm.TimeoutSeconds.Value);
            }

            if (payload.Llm.MaxTokens.HasValue)
            {
                _config.LlmFallback.MaxTokens = Math.Max(1, payload.Llm.MaxTokens.Value);
            }
        }

        AgentConfigSanitizer.Normalize(_config);
        if (_config.ProfileModeEnabled)
        {
            AgentProfileService.ApplyActiveProfile(_config);
        }

        await SaveConfigToDiskAsync(cancellationToken);
        return await GetConfigAsync(cancellationToken);
    }

    public Task<WebTasksResponse?> GetTasksAsync(CancellationToken cancellationToken)
    {
        lock (_sync)
        {
            var ordered = _tasks.OrderByDescending(t => t.UpdatedAt).ToList();
            return Task.FromResult<WebTasksResponse?>(new WebTasksResponse(ordered));
        }
    }

    public async Task<WebApiSimpleResponse?> SaveTaskAsync(WebTaskUpsertRequest payload, CancellationToken cancellationToken)
    {
        var name = payload.Name?.Trim() ?? string.Empty;
        var intent = payload.Intent?.Trim() ?? string.Empty;
        var planJson = payload.PlanJson?.Trim();
        if (string.IsNullOrWhiteSpace(name) || (string.IsNullOrWhiteSpace(intent) && string.IsNullOrWhiteSpace(planJson)))
        {
            return new WebApiSimpleResponse("Task name and intent/plan are required.", false);
        }

        if (!string.IsNullOrWhiteSpace(planJson) && !TryParseActionPlanJson(planJson, out _, out var parseError))
        {
            return new WebApiSimpleResponse($"Invalid plan JSON: {parseError}", false);
        }

        var updated = new WebTaskItem(name, intent, payload.Description?.Trim(), planJson, DateTimeOffset.UtcNow);

        lock (_sync)
        {
            _tasks.RemoveAll(task => string.Equals(task.Name, name, StringComparison.OrdinalIgnoreCase));
            _tasks.Add(updated);
        }

        await SaveTasksToDiskAsync(cancellationToken);
        return new WebApiSimpleResponse("Task saved.", true);
    }

    public async Task<WebApiSimpleResponse?> DeleteTaskAsync(string name, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            return new WebApiSimpleResponse("Task name is required.", false);
        }

        var removed = false;
        lock (_sync)
        {
            removed = _tasks.RemoveAll(task => string.Equals(task.Name, name, StringComparison.OrdinalIgnoreCase)) > 0;
        }

        if (!removed)
        {
            return new WebApiSimpleResponse("Task not found.", false);
        }

        await SaveTasksToDiskAsync(cancellationToken);
        return new WebApiSimpleResponse("Task deleted.", true);
    }

    public async Task<WebIntentResponse?> RunTaskAsync(string name, bool dryRun, CancellationToken cancellationToken)
    {
        var task = FindTask(name);
        if (task == null)
        {
            return new WebIntentResponse("Task not found.", false, null, null, null, null, null);
        }

        ActionPlan plan;
        if (!string.IsNullOrWhiteSpace(task.PlanJson))
        {
            if (!TryParseActionPlanJson(task.PlanJson, out var parsed, out var parseError))
            {
                return new WebIntentResponse($"Task plan invalid: {parseError}", false, null, null, null, task.PlanJson, null);
            }

            plan = parsed!;
        }
        else
        {
            plan = _planner.PlanFromIntent(task.Intent);
        }

        return await ExecutePlanInternalAsync($"task:{task.Name}", plan, dryRun, approvedByUser: false, cancellationToken);
    }

    public Task<WebSchedulesResponse?> GetSchedulesAsync(CancellationToken cancellationToken)
    {
        lock (_sync)
        {
            var items = _schedules
                .OrderByDescending(schedule => schedule.UpdatedAt)
                .Select(ToScheduleItem)
                .ToList();
            return Task.FromResult<WebSchedulesResponse?>(new WebSchedulesResponse(items));
        }
    }
    public async Task<WebScheduleSaveResponse?> SaveScheduleAsync(WebScheduleUpsertRequest payload, CancellationToken cancellationToken)
    {
        var taskName = payload.TaskName?.Trim() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(taskName))
        {
            return new WebScheduleSaveResponse("Task name is required.", null);
        }

        var task = FindTask(taskName);
        if (task == null)
        {
            return new WebScheduleSaveResponse("Task not found.", null);
        }

        ScheduleState schedule;
        lock (_sync)
        {
            var id = string.IsNullOrWhiteSpace(payload.Id) ? $"sch-{Guid.NewGuid():N}"[..12] : payload.Id.Trim();
            schedule = _schedules.FirstOrDefault(item => string.Equals(item.Id, id, StringComparison.OrdinalIgnoreCase))
                ?? new ScheduleState(id, task.Name, payload.StartAtUtc, payload.IntervalSeconds, payload.Enabled ?? true, DateTimeOffset.UtcNow, null);

            schedule = schedule with
            {
                TaskName = task.Name,
                StartAtUtc = payload.StartAtUtc,
                IntervalSeconds = payload.IntervalSeconds,
                Enabled = payload.Enabled ?? schedule.Enabled,
                UpdatedAt = DateTimeOffset.UtcNow
            };

            _schedules.RemoveAll(item => string.Equals(item.Id, schedule.Id, StringComparison.OrdinalIgnoreCase));
            _schedules.Add(schedule);
        }

        await SaveSchedulesToDiskAsync(cancellationToken);
        return new WebScheduleSaveResponse("Schedule saved.", ToScheduleItem(schedule));
    }

    public async Task<WebApiSimpleResponse?> DeleteScheduleAsync(string id, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(id))
        {
            return new WebApiSimpleResponse("Schedule id is required.", false);
        }

        var removed = false;
        lock (_sync)
        {
            removed = _schedules.RemoveAll(item => string.Equals(item.Id, id, StringComparison.OrdinalIgnoreCase)) > 0;
        }

        if (!removed)
        {
            return new WebApiSimpleResponse("Schedule not found.", false);
        }

        await SaveSchedulesToDiskAsync(cancellationToken);
        return new WebApiSimpleResponse("Schedule deleted.", true);
    }

    public async Task<WebApiSimpleResponse?> RunScheduleNowAsync(string id, CancellationToken cancellationToken)
    {
        ScheduleState? schedule;
        lock (_sync)
        {
            schedule = _schedules.FirstOrDefault(item => string.Equals(item.Id, id, StringComparison.OrdinalIgnoreCase));
        }

        if (schedule == null)
        {
            return new WebApiSimpleResponse("Schedule not found.", false);
        }

        var result = await RunTaskAsync(schedule.TaskName, dryRun: false, cancellationToken);
        if (result == null)
        {
            return new WebApiSimpleResponse("Schedule run failed.", false);
        }

        lock (_sync)
        {
            var index = _schedules.FindIndex(item => string.Equals(item.Id, id, StringComparison.OrdinalIgnoreCase));
            if (index >= 0)
            {
                _schedules[index] = _schedules[index] with { LastRunAtUtc = DateTimeOffset.UtcNow, UpdatedAt = DateTimeOffset.UtcNow };
            }
        }

        await SaveSchedulesToDiskAsync(cancellationToken);
        return new WebApiSimpleResponse(result.Reply, result.Reply.Contains("Success: True", StringComparison.OrdinalIgnoreCase));
    }

    public Task<WebGoalsResponse?> GetGoalsAsync(CancellationToken cancellationToken)
    {
        lock (_sync)
        {
            var items = _goals
                .OrderBy(goal => goal.Completed)
                .ThenByDescending(goal => goal.Priority)
                .ThenByDescending(goal => goal.UpdatedAtUtc)
                .Select(goal => new WebGoalItem(
                    goal.Id,
                    goal.Text,
                    goal.Completed,
                    goal.Priority,
                    goal.AutoRunEnabled,
                    goal.Attempts,
                    goal.UpdatedAtUtc,
                    goal.LastRunAtUtc,
                    goal.LastResult))
                .ToList();

            return Task.FromResult<WebGoalsResponse?>(new WebGoalsResponse(
                _config.GoalSchedulerEnabled,
                _config.GoalSchedulerIntervalSeconds,
                items));
        }
    }

    public async Task<WebApiSimpleResponse?> AddGoalFromUiAsync(string text, CancellationToken cancellationToken)
    {
        var response = await AddGoalAsync(text, cancellationToken);
        var success = response.Reply.StartsWith("Goal added", StringComparison.OrdinalIgnoreCase);
        return new WebApiSimpleResponse(response.Reply, success);
    }

    public async Task<WebApiSimpleResponse?> SetGoalAutoFromUiAsync(string id, bool enabled, CancellationToken cancellationToken)
    {
        var command = $"{id} {(enabled ? "on" : "off")}";
        var response = await SetGoalAutoModeAsync(command, cancellationToken);
        var success = response.Reply.Contains("auto mode", StringComparison.OrdinalIgnoreCase);
        return new WebApiSimpleResponse(response.Reply, success);
    }

    public async Task<WebApiSimpleResponse?> MarkGoalDoneFromUiAsync(string id, CancellationToken cancellationToken)
    {
        var response = await MarkGoalDoneAsync(id, cancellationToken);
        var success = response.Reply.Contains("marked as done", StringComparison.OrdinalIgnoreCase);
        return new WebApiSimpleResponse(response.Reply, success);
    }

    public async Task<WebApiSimpleResponse?> RemoveGoalFromUiAsync(string id, CancellationToken cancellationToken)
    {
        var response = await RemoveGoalAsync(id, cancellationToken);
        var success = response.Reply.Contains("removed", StringComparison.OrdinalIgnoreCase);
        return new WebApiSimpleResponse(response.Reply, success);
    }

    public Task<WebAuditResponse?> GetAuditAsync(int take, CancellationToken cancellationToken)
    {
        var max = Math.Clamp(take, 1, 500);
        var path = Path.GetFullPath(_config.AuditLogPath);
        if (!File.Exists(path))
        {
            return Task.FromResult<WebAuditResponse?>(new WebAuditResponse(Array.Empty<string>()));
        }

        var queue = new Queue<string>();
        foreach (var line in File.ReadLines(path))
        {
            if (queue.Count == max)
            {
                queue.Dequeue();
            }

            queue.Enqueue(line);
        }

        return Task.FromResult<WebAuditResponse?>(new WebAuditResponse(queue.ToArray()));
    }

    public async Task WriteSystemAuditAsync(string eventType, string message, object? data, CancellationToken cancellationToken)
    {
        try
        {
            await _auditLog.WriteAsync(new AuditEvent
            {
                Timestamp = DateTimeOffset.UtcNow,
                EventType = string.IsNullOrWhiteSpace(eventType) ? "system" : eventType.Trim(),
                Message = string.IsNullOrWhiteSpace(message) ? "System event" : message.Trim(),
                Data = data
            }, cancellationToken);
        }
        catch
        {
            // Best effort: tray logging must never break app flow.
        }
    }

    public Task<WebApiSimpleResponse?> RestartServerAsync(CancellationToken cancellationToken)
    {
        return Task.FromResult<WebApiSimpleResponse?>(new WebApiSimpleResponse("No separate web server in tray-only mode.", true));
    }

    public Task<WebApiSimpleResponse?> RestartAdapterAsync(CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(_config.AdapterRestartCommand))
        {
            return Task.FromResult<WebApiSimpleResponse?>(new WebApiSimpleResponse("Adapter restart command not configured.", false));
        }

        if (!TryParseCommand(_config.AdapterRestartCommand, out var fileName, out var args))
        {
            return Task.FromResult<WebApiSimpleResponse?>(new WebApiSimpleResponse("Invalid adapter restart command.", false));
        }

        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = fileName,
                Arguments = args,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            if (!string.IsNullOrWhiteSpace(_config.AdapterRestartWorkingDir))
            {
                psi.WorkingDirectory = _config.AdapterRestartWorkingDir;
            }

            Process.Start(psi);
            return Task.FromResult<WebApiSimpleResponse?>(new WebApiSimpleResponse("Adapter restart command launched.", true));
        }
        catch (Exception ex)
        {
            return Task.FromResult<WebApiSimpleResponse?>(new WebApiSimpleResponse($"Failed to launch adapter restart command: {ex.Message}", false));
        }
    }

    public async Task<WebIntentResponse?> ExecuteIntentAsync(string intent, bool dryRun, CancellationToken cancellationToken)
    {
        var plan = _planner.PlanFromIntent(intent ?? string.Empty);
        return await ExecutePlanInternalAsync(intent ?? string.Empty, plan, dryRun, approvedByUser: false, cancellationToken);
    }

    public async Task<WebIntentResponse?> ExecutePlanJsonAsync(string planJson, bool dryRun, CancellationToken cancellationToken)
    {
        if (!TryParseActionPlanJson(planJson, out var parsedPlan, out var parseError) || parsedPlan == null)
        {
            return new WebIntentResponse($"Invalid plan JSON: {parseError}", false, null, null, null, planJson, "Mode: Plan editor");
        }

        if (TryCreateOpenAppClarification(parsedPlan, out var clarification))
        {
            return new WebIntentResponse(
                clarification.Reply,
                clarification.NeedsConfirmation,
                clarification.Token,
                clarification.ActionLabel,
                clarification.Steps,
                clarification.PlanJson,
                clarification.ModeLabel ?? "Mode: Plan editor");
        }

        if (dryRun)
        {
            return await ExecutePlanInternalAsync("plan:edited", parsedPlan, dryRun: true, approvedByUser: false, cancellationToken);
        }

        var safety = await EvaluatePlanSafetyAsync(parsedPlan, cancellationToken);
        if (!safety.Allowed)
        {
            return new WebIntentResponse($"Blocked: {safety.Reason}", false, null, null, null, PlanToJson(parsedPlan), "Mode: Plan editor");
        }

        if (!safety.AutoExecute)
        {
            var token = CreatePendingAction(PendingActionType.ExecutePlan, "plan:edited", parsedPlan, dryRun: false);
            return new WebIntentResponse(
                "Edited plan requires confirmation.",
                true,
                token,
                "Confirm",
                PlanToLines(parsedPlan),
                PlanToJson(parsedPlan),
                "Mode: Plan editor");
        }

        return await ExecutePlanInternalAsync("plan:edited", parsedPlan, dryRun: false, approvedByUser: false, cancellationToken);
    }

    public void Dispose()
    {
        _schedulerCts.Cancel();
        try
        {
            _schedulerTask.Wait(TimeSpan.FromSeconds(1));
        }
        catch
        {
            // ignored on shutdown
        }

        _schedulerCts.Dispose();
        _desktopClient.Dispose();
        _loggerFactory.Dispose();
    }

    private async Task<WebChatResponse> BuildPlanConfirmationAsync(string intent, bool dryRun, CancellationToken cancellationToken)
    {
        var plan = _planner.PlanFromIntent(intent ?? string.Empty);
        if (IsUnrecognizedPlan(plan))
        {
            return await BuildUnknownCommandResponseAsync(intent ?? string.Empty, cancellationToken);
        }

        if (TryCreateOpenAppClarification(plan, out var clarification))
        {
            return clarification;
        }

        if (dryRun)
        {
            var result = await ExecutePlanInternalAsync(intent ?? string.Empty, plan, dryRun: true, approvedByUser: false, CancellationToken.None);
            return ToChatResponse(result);
        }

        var safety = await EvaluatePlanSafetyAsync(plan, cancellationToken);
        if (!safety.Allowed)
        {
            return WebChatResponse.Simple($"Blocked: {safety.Reason}");
        }

        if (safety.AutoExecute)
        {
            var result = await ExecutePlanInternalAsync(intent ?? string.Empty, plan, dryRun: false, approvedByUser: false, cancellationToken);
            return ToChatResponse(result);
        }

        var token = CreatePendingAction(PendingActionType.ExecutePlan, intent ?? string.Empty, plan, dryRun: false);
        var notice = GetRewriteNotice(plan);
        var prompt = string.IsNullOrWhiteSpace(notice)
            ? "I interpreted your request. Confirm execution?"
            : $"I interpreted your request. {notice}. Confirm execution?";
        return WebChatResponse.ConfirmWithSteps(prompt, token, PlanToLines(plan), PlanToJson(plan), GetModeLabel(plan));
    }

    private async Task<PlanSafetyResult> EvaluatePlanSafetyAsync(ActionPlan plan, CancellationToken cancellationToken)
    {
        if (RequiresLlmConfidenceConfirmation(plan, out var llmReason))
        {
            return new PlanSafetyResult(true, false, llmReason);
        }

        var activeWindow = await _desktopClient.GetActiveWindowAsync(cancellationToken);
        foreach (var step in plan.Steps)
        {
            var decision = _policyEngine.Evaluate(step, activeWindow);
            if (!decision.Allowed)
            {
                return new PlanSafetyResult(false, false, string.IsNullOrWhiteSpace(decision.Reason) ? "Blocked by policy" : decision.Reason);
            }

            if (decision.RequiresConfirmation)
            {
                return new PlanSafetyResult(true, false, decision.Reason);
            }

            if (!IsSafeAutoExecuteAction(step.Type))
            {
                return new PlanSafetyResult(true, false, "Manual confirmation required for interactive actions");
            }
        }

        return new PlanSafetyResult(true, true, string.Empty);
    }

    private static bool RequiresLlmConfidenceConfirmation(ActionPlan plan, out string reason)
    {
        reason = string.Empty;
        var note = plan.Steps
            .Select(step => step.Note)
            .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value) && value.StartsWith("Rewritten intent:", StringComparison.OrdinalIgnoreCase));

        if (string.IsNullOrWhiteSpace(note))
        {
            return false;
        }

        if (note.Contains("llm-needs-clarification", StringComparison.OrdinalIgnoreCase))
        {
            reason = "LLM requested clarification before auto-execution.";
            return true;
        }

        if (note.Contains("llm-low-confidence", StringComparison.OrdinalIgnoreCase))
        {
            reason = "LLM confidence is low; explicit confirmation required.";
            return true;
        }

        return false;
    }

    private async Task<WebChatResponse> BuildUnknownCommandResponseAsync(string input, CancellationToken cancellationToken)
    {
        if (TryBuildClarificationFromInput(input, out var clarification))
        {
            return clarification;
        }

        var suggestion = await SuggestSupportedCommandAsync(input, cancellationToken);
        if (!string.IsNullOrWhiteSpace(suggestion))
        {
            var suggestedPlan = _planner.PlanFromIntent(suggestion);
            if (!IsUnrecognizedPlan(suggestedPlan))
            {
                var token = CreatePendingAction(PendingActionType.ExecutePlan, input, suggestedPlan, dryRun: false);
                var prompt = $"I interpreted your request as: {suggestion}. Confirm execution?";
                await WriteAuditAsync(
                    "llm_suggestion_parsed",
                    "Unknown input mapped via AI suggestion",
                    cancellationToken,
                    new
                    {
                        input = input,
                        suggestion = suggestion,
                        steps = suggestedPlan.Steps.Count
                    });
                return WebChatResponse.ConfirmWithSteps(
                    prompt,
                    token,
                    PlanToLines(suggestedPlan),
                    PlanToJson(suggestedPlan),
                    "Mode: LLM interpreter");
            }

            var orderSuggestion = await TryHandleOrderIntakeAsync(suggestion, suggestion.Trim().ToLowerInvariant(), cancellationToken);
            if (orderSuggestion != null)
            {
                return orderSuggestion;
            }

            return WebChatResponse.Simple($"I couldn't map that safely. AI suggestion: {suggestion}\nIf it's correct, send that command.");
        }

        return WebChatResponse.Simple(SupportedCommandsHelp);
    }

    private async Task<WebChatResponse?> TryHandleOrderIntakeAsync(string text, string normalized, CancellationToken cancellationToken)
    {
        if (_awaitingOrderEmailInput && normalized is "cancel" or "annulla" or "stop")
        {
            _awaitingOrderEmailInput = false;
            return WebChatResponse.Simple("Order intake cancelled.");
        }

        if (normalized is "order preview" or "preview order" or "mostra ordine")
        {
            if (_latestOrderDraft == null)
            {
                return WebChatResponse.Simple("No order draft available. Use `order intake` first.");
            }

            return WebChatResponse.WithSteps(
                $"Order draft [{_latestOrderDraft.Id}]",
                _latestOrderDraft.SummaryLines,
                _latestOrderDraft.PayloadJson,
                "Mode: Order intake");
        }

        if (normalized is "order clear" or "clear order")
        {
            _latestOrderDraft = null;
            _awaitingOrderEmailInput = false;
            return WebChatResponse.Simple("Order draft cleared.");
        }

        if (normalized.StartsWith("order fill", StringComparison.Ordinal)
            || normalized.StartsWith("fill order", StringComparison.Ordinal)
            || normalized.StartsWith("compila ordine", StringComparison.Ordinal)
            || normalized.StartsWith("compila il form", StringComparison.Ordinal))
        {
            var url = ExtractFirstUrl(text);
            if (string.IsNullOrWhiteSpace(url))
            {
                return WebChatResponse.Simple("Missing target URL. Use `order fill <url>`.");
            }

            return await BuildOrderFillConfirmationAsync(url, cancellationToken);
        }

        if (normalized.StartsWith("order intake", StringComparison.Ordinal))
        {
            var payload = text.Length > "order intake".Length
                ? text["order intake".Length..].Trim().TrimStart(':')
                : string.Empty;

            if (string.IsNullOrWhiteSpace(payload))
            {
                _awaitingOrderEmailInput = true;
                return WebChatResponse.Simple("Paste the order email text/message, and I will extract structured fields.");
            }

            return await ExtractOrderDraftAsync(payload, cancellationToken);
        }

        var hasNaturalOrderFillRequest = TryParseNaturalOrderFillRequest(text, out var naturalUrl, out var inlineOrderText);

        if (_awaitingOrderEmailInput
            && !IsReservedControlCommand(normalized)
            && !hasNaturalOrderFillRequest)
        {
            return await ExtractOrderDraftAsync(text, cancellationToken);
        }

        if (hasNaturalOrderFillRequest && !string.IsNullOrWhiteSpace(naturalUrl))
        {
            if (!string.IsNullOrWhiteSpace(inlineOrderText))
            {
                var extracted = await ExtractOrderDraftAsync(inlineOrderText, cancellationToken);
                if (_latestOrderDraft == null)
                {
                    return extracted;
                }
            }

            return await BuildOrderFillConfirmationAsync(naturalUrl, cancellationToken);
        }

        if (!IsDirectIntent(normalized)
            && !TryParseTranslationIntent(text, out _)
            && await TryRouteToOrderIntakeAsync(text, cancellationToken))
        {
            _awaitingOrderEmailInput = true;
            return WebChatResponse.Simple("Got it. Paste the order email text/message and I will extract the fields for form filling.");
        }

        return null;
    }

    private static bool IsReservedControlCommand(string normalized)
    {
        if (string.IsNullOrWhiteSpace(normalized))
        {
            return false;
        }

        return normalized is "status" or "arm" or "disarm" or "kill" or "reset kill"
            || normalized.StartsWith("run ", StringComparison.Ordinal)
            || normalized.StartsWith("dry-run ", StringComparison.Ordinal)
            || normalized.StartsWith("goal ", StringComparison.Ordinal)
            || normalized.StartsWith("profile ", StringComparison.Ordinal)
            || normalized.StartsWith("list apps", StringComparison.Ordinal)
            || normalized.StartsWith("translate ", StringComparison.Ordinal);
    }

    private async Task<WebChatResponse> ExtractOrderDraftAsync(string rawText, CancellationToken cancellationToken)
    {
        if (!_config.LlmFallbackEnabled)
        {
            return WebChatResponse.Simple("LLM is disabled. Enable it to extract order fields from free text.");
        }

        var endpoint = _config.LlmFallback.Endpoint?.Trim();
        if (string.IsNullOrWhiteSpace(endpoint) || !Uri.TryCreate(endpoint, UriKind.Absolute, out var uri))
        {
            return WebChatResponse.Simple("LLM endpoint is not configured or invalid.");
        }

        if (!uri.IsLoopback && !_config.AllowNonLoopbackLlmEndpoint)
        {
            return WebChatResponse.Simple("LLM endpoint must be local unless remote LLM is enabled.");
        }

        var provider = (_config.LlmFallback.Provider ?? "ollama").Trim().ToLowerInvariant();
        uri = NormalizeLlmEndpoint(uri, provider);
        var prompt = BuildOrderExtractionPrompt(rawText);
        var maxTokens = Math.Clamp(Math.Max(512, _config.LlmFallback.MaxTokens * 4), 512, 4096);

        try
        {
            using var client = new HttpClient { Timeout = TimeSpan.FromSeconds(Math.Max(10, _config.LlmFallback.TimeoutSeconds)) };
            var raw = provider switch
            {
                "openai" => await CallOpenAiPromptAsync(client, uri, "You extract structured order data from email text. Return strict JSON only.", prompt, maxTokens, cancellationToken),
                "llama.cpp" => await CallLlamaCppPromptAsync(client, uri, prompt, maxTokens, cancellationToken),
                _ => await CallOllamaPromptAsync(client, uri, prompt, maxTokens, cancellationToken)
            };

            if (!TryParseOrderDraftJson(raw, out var payloadJson, out var summaryLines))
            {
                if (!TryBuildOrderDraftFromTextHeuristics(rawText, out var fallbackPayloadJson, out var fallbackSummaryLines))
                {
                    return WebChatResponse.Simple("I couldn't extract a valid order draft. Paste a fuller email text with customer/items/address.");
                }

                payloadJson = fallbackPayloadJson;
                summaryLines = fallbackSummaryLines;
            }
            else if (TryParseOrderPayload(payloadJson, out var parsedPayload))
            {
                // Merge deterministic regex extraction to fill gaps from LLM output.
                EnrichOrderPayloadFromText(parsedPayload, rawText);
                payloadJson = JsonSerializer.Serialize(parsedPayload, JsonOptions);
                summaryLines = BuildOrderSummary(parsedPayload);
            }

            var draft = new OrderDraft(
                Id: Guid.NewGuid().ToString("N")[..8],
                CreatedAtUtc: DateTimeOffset.UtcNow,
                RawText: rawText,
                PayloadJson: payloadJson,
                SummaryLines: summaryLines);

            _latestOrderDraft = draft;
            _awaitingOrderEmailInput = false;

            await WriteAuditAsync("order_draft_extracted", "Order draft extracted from free text", cancellationToken, new
            {
                draftId = draft.Id,
                preview = summaryLines.Take(6).ToArray()
            });

            return WebChatResponse.WithSteps(
                $"Order draft extracted [{draft.Id}]. Review fields, then ask to fill your target form.",
                summaryLines,
                payloadJson,
                "Mode: LLM order intake");
        }
        catch (Exception ex)
        {
            await WriteAuditAsync("order_draft_error", "Order draft extraction failed", cancellationToken, new { error = ex.Message });
            return WebChatResponse.Simple($"Order extraction failed: {Compact(ex.Message, 140)}");
        }
    }

    private async Task<WebChatResponse> BuildOrderFillConfirmationAsync(string targetUrl, CancellationToken cancellationToken)
    {
        if (_latestOrderDraft == null)
        {
            _awaitingOrderEmailInput = true;
            return WebChatResponse.Simple("No order draft available. Paste the order email text first (or use `order intake <text>`), then run `order fill <url>`.");
        }

        if (!TryParseOrderPayload(_latestOrderDraft.PayloadJson, out var payload))
        {
            return WebChatResponse.Simple("Order draft is invalid. Run `order intake` again.");
        }

        var isGoogleForms = IsGoogleFormsUrl(targetUrl);
        var googleOrderedValues = isGoogleForms ? BuildGoogleFormsOrderedValues(payload) : new List<(string Key, string Label, string Value)>();
        SmartOrderFillPlan? smart = null;
        var usingGoogleTabFallback = isGoogleForms;
        ActionPlan plan;
        if (isGoogleForms)
        {
            // Google Forms UI tree is usually not reliable via desktop accessibility.
            // Always prefer deterministic keyboard tab-fill for MVP reliability.
            plan = BuildGoogleFormsInteractionPlan(payload, targetUrl, googleOrderedValues);
        }
        else
        {
            smart = await TryBuildSmartOrderFillPlanAsync(payload, targetUrl, cancellationToken);
            plan = smart?.Plan ?? BuildOrderFillPlan(payload, targetUrl);
        }
        if (plan.Steps.Count <= 2)
        {
            return WebChatResponse.Simple("No fillable fields found in current order draft.");
        }

        var token = CreatePendingAction(PendingActionType.ExecutePlan, $"order-fill:{targetUrl}", plan, dryRun: false);
        await WriteAuditAsync(
            "order_fill_plan",
            "Order fill plan generated",
            cancellationToken,
            new
            {
                draftId = _latestOrderDraft.Id,
                url = targetUrl,
                steps = plan.Steps.Count,
                smart = smart != null,
                mapped = smart?.MappedFields ?? 0,
                discovered = smart?.DiscoveredFields ?? 0,
                googleForms = isGoogleForms,
                googleTabFallback = usingGoogleTabFallback,
                googleValues = googleOrderedValues.Count
            });

        var lines = PlanToLines(plan)
            .Concat(new[]
            {
                $"Draft: {_latestOrderDraft.Id}",
                usingGoogleTabFallback
                    ? "Mapping: google forms tab-fill mode"
                    : smart == null
                    ? "Mapping: fallback heuristics"
                    : $"Mapping: smart ({smart.MappedFields}/{smart.DiscoveredFields} fields)"
            })
            .ToList();
        return WebChatResponse.ConfirmWithSteps(
            $"Order draft ready. Fill form at {targetUrl}? (Submit action is not included.)",
            token,
            lines,
            PlanToJson(plan),
            usingGoogleTabFallback
                ? "Mode: Order autofill (google forms)"
                : smart == null
                    ? "Mode: Order autofill"
                    : "Mode: Order autofill (smart)");
    }

    private static bool IsGoogleFormsUrl(string targetUrl)
    {
        if (!Uri.TryCreate(targetUrl, UriKind.Absolute, out var uri))
        {
            return false;
        }

        var host = uri.Host.ToLowerInvariant();
        return host.Contains("forms.gle", StringComparison.Ordinal)
               || (host.Contains("docs.google.com", StringComparison.Ordinal)
                   && uri.AbsolutePath.Contains("/forms", StringComparison.OrdinalIgnoreCase));
    }

    private static ActionPlan BuildGoogleFormsInteractionPlan(
        OrderDraftPayload payload,
        string targetUrl,
        List<(string Key, string Label, string Value)>? orderedValues = null)
    {
        var plan = new ActionPlan { Intent = $"order fill google-forms {targetUrl}" };
        plan.Steps.Add(new PlanStep
        {
            Type = ActionType.OpenUrl,
            Target = targetUrl
        });
        plan.Steps.Add(new PlanStep
        {
            Type = ActionType.WaitFor,
            WaitFor = TimeSpan.FromMilliseconds(2200)
        });

        orderedValues ??= BuildGoogleFormsOrderedValues(payload);

        if (orderedValues.Count == 0)
        {
            return BuildOrderFillPlan(payload, targetUrl);
        }

        plan.Steps.Add(new PlanStep
        {
            Type = ActionType.Click,
            Selector = new Selector { NameContains = "Name" }
        });
        plan.Steps.Add(new PlanStep
        {
            Type = ActionType.WaitFor,
            WaitFor = TimeSpan.FromMilliseconds(180)
        });

        var selectAllKeys = GetSelectAllKeys();
        for (var i = 0; i < orderedValues.Count; i++)
        {
            var value = orderedValues[i];

            if (!string.IsNullOrWhiteSpace(value.Value))
            {
                plan.Steps.Add(new PlanStep
                {
                    Type = ActionType.KeyCombo,
                    Keys = selectAllKeys
                });
                plan.Steps.Add(new PlanStep
                {
                    Type = ActionType.TypeText,
                    Text = value.Value,
                    Note = $"fill-field:{value.Key}"
                });
            }

            if (i >= orderedValues.Count - 1)
            {
                continue;
            }

            plan.Steps.Add(new PlanStep
            {
                Type = ActionType.WaitFor,
                WaitFor = TimeSpan.FromMilliseconds(80)
            });
            plan.Steps.Add(new PlanStep
            {
                Type = ActionType.KeyCombo,
                Keys = new List<string> { "tab" }
            });
            plan.Steps.Add(new PlanStep
            {
                Type = ActionType.WaitFor,
                WaitFor = TimeSpan.FromMilliseconds(130)
            });
        }

        return plan;
    }

    private static List<(string Key, string Label, string Value)> BuildGoogleFormsOrderedValues(OrderDraftPayload payload)
    {
        var map = BuildCanonicalGoogleFormsValueMap(payload);
        return new List<(string Key, string Label, string Value)>
        {
            ("customer_name", "Name", map.GetValueOrDefault("customer_name", string.Empty)),
            ("customer_email", "Email", map.GetValueOrDefault("customer_email", string.Empty)),
            ("shipping_address", "Address", map.GetValueOrDefault("shipping_address", string.Empty)),
            ("customer_phone", "Phone number", map.GetValueOrDefault("customer_phone", string.Empty)),
            ("order_notes", "Comments", map.GetValueOrDefault("order_notes", string.Empty))
        };
    }

    private static Dictionary<string, string> BuildCanonicalGoogleFormsValueMap(OrderDraftPayload payload)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var candidates = new List<string>();

        AddCandidate(candidates, payload.Customer?.Name);
        AddCandidate(candidates, payload.Customer?.Email);
        AddCandidate(candidates, payload.Customer?.Phone);
        AddCandidate(candidates, payload.ShippingAddress);
        AddCandidate(candidates, payload.BillingAddress);
        AddCandidate(candidates, payload.Notes);
        AddCandidate(candidates, BuildOrderItemsSummary(payload.Items, payload.Notes));

        var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        var email = PickCandidate(candidates, IsLikelyEmail, used, payload.Customer?.Email);
        var phone = PickCandidate(candidates, IsLikelyPhone, used, payload.Customer?.Phone);
        var address = PickCandidate(candidates, IsLikelyAddress, used, payload.ShippingAddress, payload.BillingAddress);
        var name = PickCandidate(candidates, IsLikelyPersonName, used, payload.Customer?.Name);
        var comments = PickCandidate(candidates, IsLikelyCommentText, used, payload.Notes, BuildOrderItemsSummary(payload.Items, payload.Notes));

        AddMap(result, "customer_name", name);
        AddMap(result, "customer_email", email);
        AddMap(result, "shipping_address", address);
        AddMap(result, "customer_phone", phone);
        AddMap(result, "order_notes", comments);

        return result;
    }

    private static void AddCandidate(List<string> candidates, string? value)
    {
        var text = value?.Trim();
        if (string.IsNullOrWhiteSpace(text))
        {
            return;
        }

        if (!candidates.Any(existing => existing.Equals(text, StringComparison.OrdinalIgnoreCase)))
        {
            candidates.Add(text);
        }
    }

    private static string? PickCandidate(
        IReadOnlyList<string> candidates,
        Func<string, bool> predicate,
        HashSet<string> used,
        params string?[] preferred)
    {
        foreach (var pref in preferred)
        {
            var text = pref?.Trim();
            if (string.IsNullOrWhiteSpace(text))
            {
                continue;
            }

            if (used.Contains(text) || !predicate(text))
            {
                continue;
            }

            used.Add(text);
            return text;
        }

        foreach (var candidate in candidates)
        {
            if (used.Contains(candidate) || !predicate(candidate))
            {
                continue;
            }

            used.Add(candidate);
            return candidate;
        }

        return null;
    }

    private static bool IsLikelyEmail(string value)
        => Regex.IsMatch(value, @"(?i)^[A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,}$", RegexOptions.CultureInvariant);

    private static bool IsLikelyPhone(string value)
        => Regex.IsMatch(value, @"^\+?\d[\d\s().-]{6,}\d$", RegexOptions.CultureInvariant);

    private static bool IsLikelyAddress(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return false;
        }

        var text = value.Trim();
        if (Regex.IsMatch(text, @"\b(via|viale|piazza|corso|street|st\.|road|rd\.|avenue|ave\.)\b", RegexOptions.IgnoreCase))
        {
            return true;
        }

        return Regex.IsMatch(text, @"\d{1,5}.*\b[A-Za-z]{2,}", RegexOptions.CultureInvariant)
               && text.Contains(',', StringComparison.Ordinal);
    }

    private static bool IsLikelyPersonName(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return false;
        }

        if (IsLikelyEmail(value) || IsLikelyPhone(value) || IsLikelyAddress(value))
        {
            return false;
        }

        return Regex.IsMatch(value.Trim(), @"^[\p{L}'-]+(?:\s+[\p{L}'-]+){1,3}$", RegexOptions.CultureInvariant);
    }

    private static bool IsLikelyCommentText(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return false;
        }

        if (IsLikelyEmail(value) || IsLikelyPhone(value))
        {
            return false;
        }

        return value.Contains("comment", StringComparison.OrdinalIgnoreCase)
               || value.Contains("note", StringComparison.OrdinalIgnoreCase)
               || value.Length >= 16;
    }

    private static List<string> GetSelectAllKeys()
    {
        return OperatingSystem.IsMacOS()
            ? new List<string> { "cmd", "a" }
            : new List<string> { "ctrl", "a" };
    }

    private async Task<SmartOrderFillPlan?> TryBuildSmartOrderFillPlanAsync(
        OrderDraftPayload payload,
        string targetUrl,
        CancellationToken cancellationToken)
    {
        var discovered = await TryDiscoverFormFieldsFromUrlAsync(targetUrl, cancellationToken);
        if (discovered.Count == 0)
        {
            return null;
        }

        var values = BuildOrderValueMap(payload);
        if (values.Count == 0)
        {
            return null;
        }

        var mapped = await MapOrderValuesToDiscoveredFieldsAsync(values, discovered, cancellationToken);
        if (mapped.Count == 0)
        {
            return null;
        }

        var plan = new ActionPlan { Intent = $"order fill smart {targetUrl}" };
        plan.Steps.Add(new PlanStep
        {
            Type = ActionType.OpenUrl,
            Target = targetUrl
        });
        plan.Steps.Add(new PlanStep
        {
            Type = ActionType.WaitFor,
            WaitFor = TimeSpan.FromMilliseconds(1200)
        });

        foreach (var pair in values)
        {
            if (!mapped.TryGetValue(pair.Key, out var field) || string.IsNullOrWhiteSpace(pair.Value))
            {
                continue;
            }

            AddDiscoveredFieldSetValueCandidates(plan, pair.Key, pair.Value, field);
            AddOptionalSetValueCandidates(plan, pair.Value, pair.Key, GetDefaultOrderFieldHints(pair.Key));
        }

        if (plan.Steps.Count <= 2)
        {
            return null;
        }

        return new SmartOrderFillPlan(plan, mapped.Count, discovered.Count);
    }

    private static void AddDiscoveredFieldSetValueCandidates(ActionPlan plan, string group, string value, DiscoveredFormField field)
    {
        if (!string.IsNullOrWhiteSpace(field.AutomationId))
        {
            plan.Steps.Add(new PlanStep
            {
                Type = ActionType.SetValue,
                Selector = new Selector
                {
                    AutomationId = field.AutomationId,
                    NameContains = field.BestTextHint
                },
                Text = value,
                Note = $"optional-group:{group};optional;fill-field:{group}"
            });
        }

        if (!string.IsNullOrWhiteSpace(field.BestTextHint))
        {
            plan.Steps.Add(new PlanStep
            {
                Type = ActionType.SetValue,
                Selector = new Selector { NameContains = field.BestTextHint },
                Text = value,
                Note = $"optional-group:{group};optional;fill-field:{group}"
            });
        }
    }

    private static ActionPlan BuildOrderFillPlan(OrderDraftPayload payload, string targetUrl)
    {
        var plan = new ActionPlan { Intent = $"order fill {targetUrl}" };
        plan.Steps.Add(new PlanStep
        {
            Type = ActionType.OpenUrl,
            Target = targetUrl
        });
        plan.Steps.Add(new PlanStep
        {
            Type = ActionType.WaitFor,
            WaitFor = TimeSpan.FromMilliseconds(1200)
        });

        AddOptionalSetValueCandidates(
            plan,
            payload.Customer?.Name,
            "customer_name",
            GetDefaultOrderFieldHints("customer_name"));

        AddOptionalSetValueCandidates(
            plan,
            payload.Customer?.Email,
            "customer_email",
            GetDefaultOrderFieldHints("customer_email"));

        AddOptionalSetValueCandidates(
            plan,
            payload.Customer?.Phone,
            "customer_phone",
            GetDefaultOrderFieldHints("customer_phone"));

        AddOptionalSetValueCandidates(
            plan,
            payload.OrderNumber,
            "order_number",
            GetDefaultOrderFieldHints("order_number"));

        AddOptionalSetValueCandidates(
            plan,
            payload.ShippingAddress,
            "shipping_address",
            GetDefaultOrderFieldHints("shipping_address"));

        AddOptionalSetValueCandidates(
            plan,
            payload.BillingAddress,
            "billing_address",
            GetDefaultOrderFieldHints("billing_address"));

        var notes = BuildOrderItemsSummary(payload.Items, payload.Notes);
        AddOptionalSetValueCandidates(
            plan,
            notes,
            "order_notes",
            GetDefaultOrderFieldHints("order_notes"));

        return plan;
    }

    private static string[] GetDefaultOrderFieldHints(string key)
    {
        return key switch
        {
            "customer_name" => new[] { "customer name", "full name", "name", "nome" },
            "customer_email" => new[] { "customer email", "email", "e-mail", "mail" },
            "customer_phone" => new[] { "phone", "telephone", "telefono", "mobile" },
            "order_number" => new[] { "order number", "order id", "order no", "numero ordine" },
            "shipping_address" => new[] { "shipping address", "delivery address", "ship to", "indirizzo spedizione" },
            "billing_address" => new[] { "billing address", "invoice address", "bill to", "indirizzo fatturazione" },
            "order_notes" => new[] { "order notes", "notes", "message", "comment", "additional info", "note" },
            _ => Array.Empty<string>()
        };
    }

    private static void AddOptionalSetValueCandidates(ActionPlan plan, string? value, string group, params string[] fieldHints)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return;
        }

        foreach (var hint in fieldHints.Where(h => !string.IsNullOrWhiteSpace(h)).Distinct(StringComparer.OrdinalIgnoreCase))
        {
            plan.Steps.Add(new PlanStep
            {
                Type = ActionType.SetValue,
                Selector = new Selector { NameContains = hint },
                Text = value,
                Note = $"optional-group:{group};optional;fill-field:{group}"
            });
        }
    }

    private static string? BuildOrderItemsSummary(List<OrderItemPayload>? items, string? notes)
    {
        var lines = new List<string>();
        if (items != null && items.Count > 0)
        {
            foreach (var item in items.Take(6))
            {
                var qty = item.Qty.HasValue ? item.Qty.Value.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture) : "1";
                var name = string.IsNullOrWhiteSpace(item.Name) ? item.Sku ?? "item" : item.Name;
                if (!string.IsNullOrWhiteSpace(name))
                {
                    lines.Add($"{qty} x {name}");
                }
            }
        }

        if (!string.IsNullOrWhiteSpace(notes))
        {
            lines.Add(notes.Trim());
        }

        return lines.Count == 0 ? null : string.Join("; ", lines);
    }

    private static Dictionary<string, string> BuildOrderValueMap(OrderDraftPayload payload)
    {
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        AddMap(map, "customer_name", payload.Customer?.Name);
        AddMap(map, "customer_email", payload.Customer?.Email);
        AddMap(map, "customer_phone", payload.Customer?.Phone);
        AddMap(map, "order_number", payload.OrderNumber);
        AddMap(map, "shipping_address", payload.ShippingAddress);
        AddMap(map, "billing_address", payload.BillingAddress);
        AddMap(map, "order_notes", BuildOrderItemsSummary(payload.Items, payload.Notes));
        return map;
    }

    private static void AddMap(Dictionary<string, string> map, string key, string? value)
    {
        var text = value?.Trim();
        if (!string.IsNullOrWhiteSpace(text))
        {
            map[key] = text;
        }
    }

    private async Task<IReadOnlyDictionary<string, DiscoveredFormField>> MapOrderValuesToDiscoveredFieldsAsync(
        IReadOnlyDictionary<string, string> values,
        IReadOnlyList<DiscoveredFormField> fields,
        CancellationToken cancellationToken)
    {
        var llmMapped = await TryMapFieldsWithLlmAsync(values.Keys.ToList(), fields, cancellationToken);
        if (llmMapped.Count > 0)
        {
            return llmMapped;
        }

        return MapFieldsDeterministically(values.Keys.ToList(), fields);
    }

    private IReadOnlyDictionary<string, DiscoveredFormField> MapFieldsDeterministically(
        IReadOnlyList<string> orderKeys,
        IReadOnlyList<DiscoveredFormField> fields)
    {
        var result = new Dictionary<string, DiscoveredFormField>(StringComparer.OrdinalIgnoreCase);
        var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var key in orderKeys)
        {
            var keywords = GetDefaultOrderFieldHints(key);
            if (keywords.Length == 0)
            {
                continue;
            }

            var best = fields
                .Where(field => !used.Contains(field.Key))
                .Select(field => new { Field = field, Score = ScoreField(field, keywords) })
                .Where(item => item.Score > 0)
                .OrderByDescending(item => item.Score)
                .ThenByDescending(item => item.Field.BestTextHint.Length)
                .FirstOrDefault();
            if (best == null)
            {
                continue;
            }

            result[key] = best.Field;
            used.Add(best.Field.Key);
        }

        return result;
    }

    private static int ScoreField(DiscoveredFormField field, IReadOnlyList<string> keywords)
    {
        if (keywords.Count == 0)
        {
            return 0;
        }

        var text = $"{field.BestTextHint} {field.AutomationId} {field.Name} {field.Placeholder}".ToLowerInvariant();
        var score = 0;
        foreach (var keyword in keywords.Where(k => !string.IsNullOrWhiteSpace(k)))
        {
            var token = keyword.ToLowerInvariant();
            if (text.Equals(token, StringComparison.Ordinal))
            {
                score += 8;
                continue;
            }

            if (text.Contains(token, StringComparison.Ordinal))
            {
                score += 4;
                continue;
            }

            var compactToken = token.Replace(" ", string.Empty, StringComparison.Ordinal);
            if (!string.IsNullOrWhiteSpace(compactToken)
                && text.Replace(" ", string.Empty, StringComparison.Ordinal).Contains(compactToken, StringComparison.Ordinal))
            {
                score += 2;
            }
        }

        return score;
    }

    private async Task<IReadOnlyDictionary<string, DiscoveredFormField>> TryMapFieldsWithLlmAsync(
        IReadOnlyList<string> orderKeys,
        IReadOnlyList<DiscoveredFormField> fields,
        CancellationToken cancellationToken)
    {
        if (!_config.LlmFallbackEnabled || fields.Count == 0 || orderKeys.Count == 0)
        {
            return new Dictionary<string, DiscoveredFormField>(StringComparer.OrdinalIgnoreCase);
        }

        var endpoint = _config.LlmFallback.Endpoint?.Trim();
        if (string.IsNullOrWhiteSpace(endpoint) || !Uri.TryCreate(endpoint, UriKind.Absolute, out var uri))
        {
            return new Dictionary<string, DiscoveredFormField>(StringComparer.OrdinalIgnoreCase);
        }

        if (!uri.IsLoopback && !_config.AllowNonLoopbackLlmEndpoint)
        {
            return new Dictionary<string, DiscoveredFormField>(StringComparer.OrdinalIgnoreCase);
        }

        var provider = (_config.LlmFallback.Provider ?? "ollama").Trim().ToLowerInvariant();
        uri = NormalizeLlmEndpoint(uri, provider);
        var prompt = BuildOrderFieldMappingPrompt(orderKeys, fields);

        try
        {
            using var client = new HttpClient { Timeout = _httpTimeout };
            var raw = provider switch
            {
                "openai" => await CallOpenAiPromptAsync(client, uri, "Map form fields to order keys. Return strict JSON only.", prompt, 256, cancellationToken),
                "llama.cpp" => await CallLlamaCppPromptAsync(client, uri, prompt, 256, cancellationToken),
                _ => await CallOllamaPromptAsync(client, uri, prompt, 256, cancellationToken)
            };

            var json = TryExtractJsonObject(raw ?? string.Empty);
            if (string.IsNullOrWhiteSpace(json))
            {
                return new Dictionary<string, DiscoveredFormField>(StringComparer.OrdinalIgnoreCase);
            }

            using var doc = JsonDocument.Parse(json);
            if (doc.RootElement.ValueKind != JsonValueKind.Object)
            {
                return new Dictionary<string, DiscoveredFormField>(StringComparer.OrdinalIgnoreCase);
            }

            var byKey = fields.ToDictionary(field => field.Key, StringComparer.OrdinalIgnoreCase);
            var mapped = new Dictionary<string, DiscoveredFormField>(StringComparer.OrdinalIgnoreCase);
            var usedFields = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var property in doc.RootElement.EnumerateObject())
            {
                var orderKey = property.Name;
                var fieldKey = property.Value.ValueKind == JsonValueKind.String
                    ? property.Value.GetString()
                    : null;
                if (string.IsNullOrWhiteSpace(fieldKey))
                {
                    continue;
                }

                if (!orderKeys.Contains(orderKey, StringComparer.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (byKey.TryGetValue(fieldKey, out var field) && usedFields.Add(field.Key))
                {
                    mapped[orderKey] = field;
                }
            }

            return mapped;
        }
        catch
        {
            return new Dictionary<string, DiscoveredFormField>(StringComparer.OrdinalIgnoreCase);
        }
    }

    private static string BuildOrderFieldMappingPrompt(IReadOnlyList<string> orderKeys, IReadOnlyList<DiscoveredFormField> fields)
    {
        var fieldLines = fields
            .Take(80)
            .Select(field => $"- key={field.Key} | hint={field.BestTextHint} | id={field.AutomationId} | name={field.Name} | type={field.Type} | placeholder={field.Placeholder}");

        return
            "Map order keys to discovered form fields.\n" +
            "Return STRICT JSON object only where each property key is an order key and value is one field key.\n" +
            "Use only keys from the provided field list.\n" +
            "Do not invent keys.\n" +
            "Order keys:\n" +
            string.Join(", ", orderKeys) + "\n" +
            "Fields:\n" +
            string.Join("\n", fieldLines) + "\n" +
            "Output example:\n" +
            "{\"customer_name\":\"field_1\",\"customer_email\":\"field_2\"}";
    }

    private async Task<IReadOnlyList<DiscoveredFormField>> TryDiscoverFormFieldsFromUrlAsync(string targetUrl, CancellationToken cancellationToken)
    {
        try
        {
            using var client = new HttpClient { Timeout = _httpTimeout };
            using var response = await client.GetAsync(targetUrl, cancellationToken);
            if (!response.IsSuccessStatusCode)
            {
                return Array.Empty<DiscoveredFormField>();
            }

            var contentType = response.Content.Headers.ContentType?.MediaType ?? string.Empty;
            if (!contentType.Contains("html", StringComparison.OrdinalIgnoreCase))
            {
                return Array.Empty<DiscoveredFormField>();
            }

            var html = await response.Content.ReadAsStringAsync(cancellationToken);
            return ParseDiscoveredFormFields(html);
        }
        catch
        {
            return Array.Empty<DiscoveredFormField>();
        }
    }

    private static IReadOnlyList<DiscoveredFormField> ParseDiscoveredFormFields(string html)
    {
        if (string.IsNullOrWhiteSpace(html))
        {
            return Array.Empty<DiscoveredFormField>();
        }

        var labels = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (Match labelMatch in Regex.Matches(
                     html,
                     "<label[^>]*for\\s*=\\s*['\"](?<for>[^'\"]+)['\"][^>]*>(?<text>[\\s\\S]*?)</label>",
                     RegexOptions.IgnoreCase))
        {
            var forId = labelMatch.Groups["for"].Value.Trim();
            var text = CleanHtmlText(labelMatch.Groups["text"].Value);
            if (!string.IsNullOrWhiteSpace(forId) && !string.IsNullOrWhiteSpace(text))
            {
                labels[forId] = text;
            }
        }

        var result = new List<DiscoveredFormField>();
        var index = 0;
        foreach (Match nodeMatch in Regex.Matches(html, "<(?<tag>input|textarea|select)\\b(?<attrs>[^>]*?)>", RegexOptions.IgnoreCase))
        {
            var tag = nodeMatch.Groups["tag"].Value.Trim().ToLowerInvariant();
            var attrs = ParseHtmlAttributes(nodeMatch.Groups["attrs"].Value);
            var type = attrs.TryGetValue("type", out var typeValue) ? typeValue.ToLowerInvariant() : tag;
            if (type is "hidden" or "submit" or "button" or "checkbox" or "radio" or "file")
            {
                continue;
            }

            var id = attrs.TryGetValue("id", out var idValue) ? idValue : string.Empty;
            var name = attrs.TryGetValue("name", out var nameValue) ? nameValue : string.Empty;
            var placeholder = attrs.TryGetValue("placeholder", out var placeholderValue) ? placeholderValue : string.Empty;
            var aria = attrs.TryGetValue("aria-label", out var ariaLabel) ? ariaLabel : string.Empty;
            var autocomplete = attrs.TryGetValue("autocomplete", out var auto) ? auto : string.Empty;
            labels.TryGetValue(id, out var labelText);
            var bestText = FirstNonEmpty(labelText, aria, placeholder, name, id);
            if (string.IsNullOrWhiteSpace(bestText))
            {
                continue;
            }

            index++;
            result.Add(new DiscoveredFormField(
                Key: $"field_{index}",
                BestTextHint: bestText,
                AutomationId: id,
                Name: name,
                Placeholder: placeholder,
                Type: type,
                AutoComplete: autocomplete));
        }

        return result;
    }

    private static Dictionary<string, string> ParseHtmlAttributes(string raw)
    {
        var attrs = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (string.IsNullOrWhiteSpace(raw))
        {
            return attrs;
        }

        foreach (Match match in Regex.Matches(
                     raw,
                     "(?<name>[a-zA-Z_:][-a-zA-Z0-9_:.]*)\\s*=\\s*(?:\"(?<v1>[^\"]*)\"|'(?<v2>[^']*)'|(?<v3>[^\\s\"'>]+))",
                     RegexOptions.IgnoreCase))
        {
            var name = match.Groups["name"].Value.Trim();
            if (string.IsNullOrWhiteSpace(name))
            {
                continue;
            }

            var value = match.Groups["v1"].Success
                ? match.Groups["v1"].Value
                : match.Groups["v2"].Success
                    ? match.Groups["v2"].Value
                    : match.Groups["v3"].Value;
            attrs[name] = WebUtility.HtmlDecode(value ?? string.Empty).Trim();
        }

        return attrs;
    }

    private static string CleanHtmlText(string input)
    {
        if (string.IsNullOrWhiteSpace(input))
        {
            return string.Empty;
        }

        var noTags = Regex.Replace(input, "<[^>]+>", " ");
        var decoded = WebUtility.HtmlDecode(noTags);
        return Regex.Replace(decoded, "\\s+", " ").Trim();
    }

    private static string FirstNonEmpty(params string?[] values)
    {
        foreach (var value in values)
        {
            if (!string.IsNullOrWhiteSpace(value))
            {
                return value.Trim();
            }
        }

        return string.Empty;
    }

    private static bool TryParseOrderPayload(string payloadJson, out OrderDraftPayload payload)
    {
        payload = new OrderDraftPayload();
        if (string.IsNullOrWhiteSpace(payloadJson))
        {
            return false;
        }

        try
        {
            var parsed = JsonSerializer.Deserialize<OrderDraftPayload>(payloadJson, JsonOptions);
            if (parsed == null)
            {
                return false;
            }

            payload = parsed;
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static bool TryParseNaturalOrderFillRequest(string text, out string? url, out string? inlineOrderText)
    {
        url = null;
        inlineOrderText = null;
        if (string.IsNullOrWhiteSpace(text))
        {
            return false;
        }

        url = ExtractFirstUrl(text);
        if (string.IsNullOrWhiteSpace(url))
        {
            return false;
        }

        var lower = text.ToLowerInvariant();
        var looksLikeFillRequest =
            lower.Contains("fill", StringComparison.Ordinal)
            || lower.Contains("compile", StringComparison.Ordinal)
            || lower.Contains("compila", StringComparison.Ordinal)
            || lower.Contains("form", StringComparison.Ordinal)
            || lower.Contains("modulo", StringComparison.Ordinal)
            || lower.Contains("ordine", StringComparison.Ordinal)
            || lower.Contains("order", StringComparison.Ordinal);
        if (!looksLikeFillRequest)
        {
            return false;
        }

        inlineOrderText = TryExtractInlineOrderText(text);
        return true;
    }

    private static string? TryExtractInlineOrderText(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return null;
        }

        var quoteMatch = Regex.Match(text, "\"(?<body>[\\s\\S]{24,})\"", RegexOptions.Singleline);
        if (quoteMatch.Success)
        {
            var body = quoteMatch.Groups["body"].Value.Trim();
            if (!string.IsNullOrWhiteSpace(body))
            {
                return body;
            }
        }

        var marker = text.IndexOf("mail:", StringComparison.OrdinalIgnoreCase);
        var markerLength = 5;
        if (marker < 0)
        {
            marker = text.IndexOf("email:", StringComparison.OrdinalIgnoreCase);
            markerLength = 6;
        }

        if (marker < 0)
        {
            return null;
        }

        var bodyPart = text[(marker + markerLength)..].Trim();
        var url = ExtractFirstUrl(bodyPart);
        if (!string.IsNullOrWhiteSpace(url))
        {
            var idx = bodyPart.IndexOf(url, StringComparison.OrdinalIgnoreCase);
            if (idx >= 0)
            {
                bodyPart = bodyPart[..idx].Trim();
            }
        }

        return string.IsNullOrWhiteSpace(bodyPart) ? null : bodyPart;
    }

    private static string? ExtractFirstUrl(string input)
    {
        if (string.IsNullOrWhiteSpace(input))
        {
            return null;
        }

        var match = Regex.Match(input, "(?<url>(?:https?://|www\\.)[^\\s\"'<>]+)", RegexOptions.IgnoreCase);
        if (!match.Success)
        {
            return null;
        }

        var raw = match.Groups["url"].Value.Trim().TrimEnd('.', ',', ';', ')', ']', '}');
        if (string.IsNullOrWhiteSpace(raw))
        {
            return null;
        }

        if (!raw.StartsWith("http://", StringComparison.OrdinalIgnoreCase)
            && !raw.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
        {
            raw = $"https://{raw}";
        }

        return Uri.TryCreate(raw, UriKind.Absolute, out var uri) ? uri.ToString() : null;
    }

    private static string BuildOrderExtractionPrompt(string rawText)
    {
        return
            "Extract structured order data from the following text.\n" +
            "Return STRICT JSON only (no markdown), using this schema:\n" +
            "{\n" +
            "  \"orderNumber\": string|null,\n" +
            "  \"orderDate\": string|null,\n" +
            "  \"customer\": {\"name\": string|null, \"email\": string|null, \"phone\": string|null},\n" +
            "  \"shippingAddress\": string|null,\n" +
            "  \"billingAddress\": string|null,\n" +
            "  \"items\": [{\"sku\": string|null, \"name\": string|null, \"qty\": number|null, \"unitPrice\": number|null, \"currency\": string|null}],\n" +
            "  \"totals\": {\"subtotal\": number|null, \"shipping\": number|null, \"tax\": number|null, \"total\": number|null, \"currency\": string|null},\n" +
            "  \"notes\": string|null\n" +
            "}\n" +
            "Use null when unknown. Keep items as empty array if none are found.\n" +
            "INPUT:\n" +
            rawText;
    }

    private static bool TryParseOrderDraftJson(string? raw, out string payloadJson, out IReadOnlyList<string> summaryLines)
    {
        payloadJson = string.Empty;
        summaryLines = Array.Empty<string>();
        if (string.IsNullOrWhiteSpace(raw))
        {
            return false;
        }

        var jsonCandidate = TryExtractJsonObject(raw);
        if (string.IsNullOrWhiteSpace(jsonCandidate))
        {
            return false;
        }

        try
        {
            using var doc = JsonDocument.Parse(jsonCandidate);
            if (doc.RootElement.ValueKind != JsonValueKind.Object)
            {
                return false;
            }

            payloadJson = JsonSerializer.Serialize(doc.RootElement, JsonOptions);
            summaryLines = BuildOrderSummary(doc.RootElement);
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static string? TryExtractJsonObject(string raw)
    {
        var text = raw.Trim();
        if (text.StartsWith('{') && text.EndsWith('}'))
        {
            return text;
        }

        var start = text.IndexOf('{');
        var end = text.LastIndexOf('}');
        if (start < 0 || end <= start)
        {
            return null;
        }

        return text.Substring(start, end - start + 1);
    }

    private static IReadOnlyList<string> BuildOrderSummary(JsonElement root)
    {
        var lines = new List<string>();
        AddSummaryLine(lines, "Order #", GetJsonString(root, "orderNumber"));
        AddSummaryLine(lines, "Order date", GetJsonString(root, "orderDate"));

        if (root.TryGetProperty("customer", out var customer) && customer.ValueKind == JsonValueKind.Object)
        {
            AddSummaryLine(lines, "Customer", GetJsonString(customer, "name"));
            AddSummaryLine(lines, "Email", GetJsonString(customer, "email"));
            AddSummaryLine(lines, "Phone", GetJsonString(customer, "phone"));
        }

        AddSummaryLine(lines, "Shipping", GetJsonString(root, "shippingAddress"));
        AddSummaryLine(lines, "Billing", GetJsonString(root, "billingAddress"));

        if (root.TryGetProperty("items", out var items) && items.ValueKind == JsonValueKind.Array)
        {
            var index = 0;
            foreach (var item in items.EnumerateArray())
            {
                index++;
                var name = item.TryGetProperty("name", out var nameEl) ? nameEl.GetString() : null;
                var sku = item.TryGetProperty("sku", out var skuEl) ? skuEl.GetString() : null;
                var qty = item.TryGetProperty("qty", out var qtyEl) ? qtyEl.ToString() : null;
                var label = !string.IsNullOrWhiteSpace(name) ? name : (!string.IsNullOrWhiteSpace(sku) ? sku : $"item {index}");
                lines.Add($"Item {index}: {label} | qty={qty ?? "?"}");
            }
        }

        if (root.TryGetProperty("totals", out var totals) && totals.ValueKind == JsonValueKind.Object)
        {
            AddSummaryLine(lines, "Total", totals.TryGetProperty("total", out var totalEl) ? totalEl.ToString() : null);
            AddSummaryLine(lines, "Currency", GetJsonString(totals, "currency"));
        }

        AddSummaryLine(lines, "Notes", GetJsonString(root, "notes"));
        if (lines.Count == 0)
        {
            lines.Add("No fields extracted.");
        }

        return lines;
    }

    private static IReadOnlyList<string> BuildOrderSummary(OrderDraftPayload payload)
    {
        var lines = new List<string>();
        AddSummaryLine(lines, "Order #", payload.OrderNumber);
        AddSummaryLine(lines, "Order date", payload.OrderDate);
        AddSummaryLine(lines, "Customer", payload.Customer?.Name);
        AddSummaryLine(lines, "Email", payload.Customer?.Email);
        AddSummaryLine(lines, "Phone", payload.Customer?.Phone);
        AddSummaryLine(lines, "Shipping", payload.ShippingAddress);
        AddSummaryLine(lines, "Billing", payload.BillingAddress);

        if (payload.Items is { Count: > 0 })
        {
            var index = 0;
            foreach (var item in payload.Items)
            {
                index++;
                var name = !string.IsNullOrWhiteSpace(item.Name) ? item.Name : item.Sku;
                var qty = item.Qty?.ToString() ?? "?";
                if (!string.IsNullOrWhiteSpace(name))
                {
                    lines.Add($"Item {index}: {name} | qty={qty}");
                }
            }
        }

        AddSummaryLine(lines, "Notes", payload.Notes);
        if (lines.Count == 0)
        {
            lines.Add("No fields extracted.");
        }

        return lines;
    }

    private static bool TryBuildOrderDraftFromTextHeuristics(string rawText, out string payloadJson, out IReadOnlyList<string> summaryLines)
    {
        var payload = new OrderDraftPayload
        {
            Customer = new OrderCustomerPayload(),
            Items = new List<OrderItemPayload>()
        };

        EnrichOrderPayloadFromText(payload, rawText);

        var hasAny = !string.IsNullOrWhiteSpace(payload.OrderNumber)
                     || !string.IsNullOrWhiteSpace(payload.Customer?.Name)
                     || !string.IsNullOrWhiteSpace(payload.Customer?.Email)
                     || !string.IsNullOrWhiteSpace(payload.Customer?.Phone)
                     || !string.IsNullOrWhiteSpace(payload.ShippingAddress)
                     || !string.IsNullOrWhiteSpace(payload.BillingAddress)
                     || !string.IsNullOrWhiteSpace(payload.Notes);

        if (!hasAny)
        {
            payloadJson = string.Empty;
            summaryLines = Array.Empty<string>();
            return false;
        }

        payloadJson = JsonSerializer.Serialize(payload, JsonOptions);
        summaryLines = BuildOrderSummary(payload);
        return true;
    }

    private static void EnrichOrderPayloadFromText(OrderDraftPayload payload, string rawText)
    {
        if (payload.Customer == null)
        {
            payload.Customer = new OrderCustomerPayload();
        }

        var text = rawText ?? string.Empty;

        if (string.IsNullOrWhiteSpace(payload.OrderNumber)
            && TryExtractOrderNumber(text, out var orderNumber))
        {
            payload.OrderNumber = orderNumber;
        }

        if (string.IsNullOrWhiteSpace(payload.Customer.Email)
            && TryExtractEmail(text, out var email))
        {
            payload.Customer.Email = email;
        }

        if (string.IsNullOrWhiteSpace(payload.Customer.Phone)
            && TryExtractPhone(text, out var phone))
        {
            payload.Customer.Phone = phone;
        }

        if (string.IsNullOrWhiteSpace(payload.Customer.Name))
        {
            var name = TryExtractLabeledValue(text, "cliente", "customer", "name", "nome");
            if (!string.IsNullOrWhiteSpace(name))
            {
                payload.Customer.Name = name;
            }
        }

        if (string.IsNullOrWhiteSpace(payload.ShippingAddress)
            && string.IsNullOrWhiteSpace(payload.BillingAddress))
        {
            var address = TryExtractLabeledValue(text, "indirizzo", "address", "shipping address", "delivery address");
            if (!string.IsNullOrWhiteSpace(address))
            {
                payload.ShippingAddress = address;
            }
        }

        if (string.IsNullOrWhiteSpace(payload.Notes))
        {
            var notes = TryExtractLabeledValue(text, "note", "notes", "commenti", "comments", "comment");
            if (!string.IsNullOrWhiteSpace(notes))
            {
                payload.Notes = notes;
            }
        }
    }

    private static bool TryExtractOrderNumber(string text, out string value)
    {
        value = string.Empty;
        if (string.IsNullOrWhiteSpace(text))
        {
            return false;
        }

        var patterns = new[]
        {
            @"(?i)\b(?:ordine|order)\s*(?:n\.?|no\.?|number|numero|id|#)?\s*[:#-]?\s*(?<v>[A-Z0-9][A-Z0-9._/-]{2,})",
            @"(?i)\b(?<v>TEST-\d{4}-\d{3,})\b"
        };

        foreach (var pattern in patterns)
        {
            var match = Regex.Match(text, pattern, RegexOptions.CultureInvariant);
            if (!match.Success)
            {
                continue;
            }

            value = CleanExtractedValue(match.Groups["v"].Value);
            if (!string.IsNullOrWhiteSpace(value))
            {
                return true;
            }
        }

        return false;
    }

    private static bool TryExtractEmail(string text, out string value)
    {
        value = string.Empty;
        if (string.IsNullOrWhiteSpace(text))
        {
            return false;
        }

        var match = Regex.Match(text, @"(?i)(?<v>[A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,})", RegexOptions.CultureInvariant);
        if (!match.Success)
        {
            return false;
        }

        value = CleanExtractedValue(match.Groups["v"].Value);
        return !string.IsNullOrWhiteSpace(value);
    }

    private static bool TryExtractPhone(string text, out string value)
    {
        value = string.Empty;
        if (string.IsNullOrWhiteSpace(text))
        {
            return false;
        }

        var labeled = TryExtractLabeledValue(text, "telefono", "phone", "mobile", "cell");
        if (!string.IsNullOrWhiteSpace(labeled))
        {
            value = labeled;
            return true;
        }

        var match = Regex.Match(text, @"(?<v>\+?\d[\d\s().-]{6,}\d)", RegexOptions.CultureInvariant);
        if (!match.Success)
        {
            return false;
        }

        value = CleanExtractedValue(match.Groups["v"].Value);
        return !string.IsNullOrWhiteSpace(value);
    }

    private static string? TryExtractLabeledValue(string text, params string[] labels)
    {
        if (string.IsNullOrWhiteSpace(text) || labels.Length == 0)
        {
            return null;
        }

        var escapedLabels = labels
            .Where(static label => !string.IsNullOrWhiteSpace(label))
            .Select(Regex.Escape)
            .ToArray();
        if (escapedLabels.Length == 0)
        {
            return null;
        }

        var blocker = "(?:cliente|customer|name|nome|email|e-mail|telefono|phone|mobile|indirizzo|address|shipping address|delivery address|note|notes|commenti|comments?)";
        var pattern = $@"(?is)\b(?:{string.Join("|", escapedLabels)})\b\s*[:\-]?\s*(?<v>.+?)(?=(?:\b{blocker}\b\s*[:\-])|[\r\n]|$)";
        var match = Regex.Match(text, pattern, RegexOptions.CultureInvariant);
        if (!match.Success)
        {
            return null;
        }

        var value = CleanExtractedValue(match.Groups["v"].Value);
        return string.IsNullOrWhiteSpace(value) ? null : value;
    }

    private static string CleanExtractedValue(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        var text = value.Trim();
        text = text.Trim('"', '\'', ' ', '\t', '\r', '\n', '.', ';', ',', ':');
        text = Regex.Replace(text, @"\s{2,}", " ");
        return text.Trim();
    }

    private static void AddSummaryLine(List<string> lines, string label, string? value)
    {
        if (!string.IsNullOrWhiteSpace(value))
        {
            lines.Add($"{label}: {value}");
        }
    }

    private static string? GetJsonString(JsonElement element, string propertyName)
    {
        if (!element.TryGetProperty(propertyName, out var prop))
        {
            return null;
        }

        return prop.ValueKind switch
        {
            JsonValueKind.String => prop.GetString(),
            JsonValueKind.Number => prop.ToString(),
            JsonValueKind.True => "true",
            JsonValueKind.False => "false",
            _ => null
        };
    }

    private async Task<bool> TryRouteToOrderIntakeAsync(string userInput, CancellationToken cancellationToken)
    {
        if (!_config.LlmFallbackEnabled || string.IsNullOrWhiteSpace(userInput))
        {
            return false;
        }

        var endpoint = _config.LlmFallback.Endpoint?.Trim();
        if (string.IsNullOrWhiteSpace(endpoint) || !Uri.TryCreate(endpoint, UriKind.Absolute, out var uri))
        {
            return false;
        }

        if (!uri.IsLoopback && !_config.AllowNonLoopbackLlmEndpoint)
        {
            return false;
        }

        var provider = (_config.LlmFallback.Provider ?? "ollama").Trim().ToLowerInvariant();
        uri = NormalizeLlmEndpoint(uri, provider);
        var prompt = BuildOrderRoutingPrompt(userInput);

        try
        {
            using var client = new HttpClient { Timeout = _httpTimeout };
            var raw = provider switch
            {
                "openai" => await CallOpenAiPromptAsync(client, uri, "You classify if a request is about processing order data from an email/message.", prompt, 96, cancellationToken),
                "llama.cpp" => await CallLlamaCppPromptAsync(client, uri, prompt, 96, cancellationToken),
                _ => await CallOllamaPromptAsync(client, uri, prompt, 96, cancellationToken)
            };

            var answer = (raw ?? string.Empty).Trim();
            if (answer.StartsWith('{'))
            {
                var json = TryExtractJsonObject(answer);
                if (!string.IsNullOrWhiteSpace(json))
                {
                    using var doc = JsonDocument.Parse(json);
                    if (doc.RootElement.TryGetProperty("intent", out var intent))
                    {
                        var value = intent.GetString() ?? string.Empty;
                        return value.Equals("order_intake", StringComparison.OrdinalIgnoreCase);
                    }
                }
            }

            return answer.Contains("order_intake", StringComparison.OrdinalIgnoreCase)
                   || answer.Equals("YES", StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
    }

    private static string BuildOrderRoutingPrompt(string userInput)
    {
        return
            "Classify this user request. Return JSON only: {\"intent\":\"order_intake|order_fill|other\"}.\n" +
            "Use order_intake when user asks to process incoming order details (often from an email/message).\n" +
            "Use order_fill when user explicitly asks to compile/fill a form at a URL with order data.\n" +
            "User request:\n" +
            userInput;
    }

    private bool TryHandlePendingClarificationReply(string input, out ClarificationChoice? choice, out string? message)
    {
        choice = null;
        message = null;
        var pending = _pendingClarification;
        if (pending == null)
        {
            return false;
        }

        var text = input.Trim();
        if (string.IsNullOrWhiteSpace(text))
        {
            message = "Please answer with a number (e.g. 1) or 'cancel'.";
            return true;
        }

        if (text.Equals("cancel", StringComparison.OrdinalIgnoreCase)
            || text.Equals("no", StringComparison.OrdinalIgnoreCase)
            || text.Equals("skip", StringComparison.OrdinalIgnoreCase))
        {
            _pendingClarification = null;
            message = "Cancelled.";
            return true;
        }

        if (int.TryParse(text, out var index)
            && index >= 1
            && index <= pending.Options.Count)
        {
            choice = pending.Options[index - 1];
            _pendingClarification = null;
            return true;
        }

        message = $"Please choose 1-{pending.Options.Count} or type 'cancel'.";
        return true;
    }

    private bool TryCreateOpenAppClarification(ActionPlan plan, out WebChatResponse clarification)
    {
        clarification = WebChatResponse.Simple(string.Empty);
        if (plan.Steps.Count == 0)
        {
            return false;
        }

        var openStep = plan.Steps.FirstOrDefault(step => step.Type == ActionType.OpenApp && !string.IsNullOrWhiteSpace(step.AppIdOrPath));
        if (openStep == null || string.IsNullOrWhiteSpace(openStep.AppIdOrPath))
        {
            return false;
        }

        var query = openStep.AppIdOrPath.Trim();
        if (_appResolver.TryResolveApp(query, out _))
        {
            return false;
        }

        var options = _appResolver.Suggest(query, 5)
            .Where(match => match.Score >= 0.35)
            .Take(4)
            .Select(match => new ClarificationChoice(
                Label: match.Entry.Name,
                Intent: $"open {match.Entry.Path}",
                Description: $"{match.Entry.Name} ({match.Entry.Path})"))
            .ToList();

        if (options.Count == 0)
        {
            return false;
        }

        _pendingClarification = new ClarificationRequest("open-app", query, options);
        var lines = options.Select((option, idx) => $"{idx + 1}. {option.Label}").ToList();
        var prompt = $"I found multiple matches for '{query}'. Reply with 1-{options.Count} to choose, or 'cancel'.";
        clarification = WebChatResponse.WithSteps(prompt, lines, null, "Mode: Clarification");
        return true;
    }

    private bool TryBuildClarificationFromInput(string input, out WebChatResponse clarification)
    {
        clarification = WebChatResponse.Simple(string.Empty);
        if (!TryExtractOpenQuery(input, out var query))
        {
            return false;
        }

        var options = _appResolver.Suggest(query, 5)
            .Where(match => match.Score >= 0.35)
            .Take(4)
            .Select(match => new ClarificationChoice(
                Label: match.Entry.Name,
                Intent: $"open {match.Entry.Path}",
                Description: $"{match.Entry.Name} ({match.Entry.Path})"))
            .ToList();

        if (options.Count == 0)
        {
            return false;
        }

        _pendingClarification = new ClarificationRequest("open-app", query, options);
        var lines = options.Select((option, idx) => $"{idx + 1}. {option.Label}").ToList();
        clarification = WebChatResponse.WithSteps(
            $"Do you want me to open one of these apps for '{query}'? Reply with 1-{options.Count} or 'cancel'.",
            lines,
            null,
            "Mode: Clarification");
        return true;
    }

    private static bool TryExtractOpenQuery(string input, out string query)
    {
        query = string.Empty;
        if (string.IsNullOrWhiteSpace(input))
        {
            return false;
        }

        var trimmed = input.Trim();
        var patterns = new[]
        {
            @"^(?:open|launch|start|run)\s+(?<app>.+)$",
            @"^(?:apri|avvia|lancia|esegui)\s+(?<app>.+)$",
            @"^(?:can you|could you|please)\s+(?:open|launch|start)\s+(?:for me\s+)?(?<app>.+)$",
            @"^(?:potresti|puoi|per favore)\s+(?:aprire|avviare|lanciare)\s+(?<app>.+)$"
        };

        foreach (var pattern in patterns)
        {
            var match = Regex.Match(trimmed, pattern, RegexOptions.IgnoreCase);
            if (!match.Success)
            {
                continue;
            }

            query = match.Groups["app"].Value.Trim(' ', '.', '!', '?', '"', '\'');
            if (!string.IsNullOrWhiteSpace(query))
            {
                return true;
            }
        }

        return false;
    }

    private async Task<WebIntentResponse> ExecutePlanInternalAsync(string source, ActionPlan plan, bool dryRun, bool approvedByUser, CancellationToken cancellationToken)
    {
        if (plan.Steps.Count == 0)
        {
            return new WebIntentResponse("Plan is empty.", false, null, null, null, PlanToJson(plan), GetModeLabel(plan));
        }

        var beforeObservation = await CaptureObservationAsync(cancellationToken);
        var activeWindow = await _desktopClient.GetActiveWindowAsync(cancellationToken);
        var lockReason = ValidateAndApplyContextLock(plan, activeWindow);
        if (!string.IsNullOrWhiteSpace(lockReason))
        {
            return new WebIntentResponse($"Blocked: {lockReason}", false, null, null, null, PlanToJson(plan), GetModeLabel(plan));
        }

        IUserConfirmation confirmation = approvedByUser
            ? new AutoApproveConfirmation()
            : new RejectConfirmation();

        var executor = new Executor(
            _desktopClient,
            _contextProvider,
            _appResolver,
            _policyEngine,
            _rateLimiter,
            _auditLog,
            confirmation,
            _killSwitch,
            _config,
            _ocrEngine,
            _loggerFactory.CreateLogger<Executor>());

        var result = await executor.ExecutePlanAsync(plan, dryRun, cancellationToken);
        string recoveryNote = string.Empty;
        if (!dryRun && _config.AutoRecoveryEnabled && _config.AutoRecoveryMaxAttempts > 0)
        {
            var recovered = await TryAutoRecoverAsync(plan, result, executor, cancellationToken);
            if (recovered.RecoveredResult != null)
            {
                result = MergeExecutionResults(result, recovered.RecoveredResult);
                recoveryNote = recovered.Note;
            }
        }

        var afterObservation = await CaptureObservationAsync(cancellationToken);
        await RecordIntentMemoryAsync(source, plan, result, beforeObservation, afterObservation, cancellationToken);
        await UpdateGoalProgressAsync(source, result, cancellationToken);

        var perceptionNote = BuildPerceptionNote(beforeObservation, afterObservation);
        var reply = AppendNotice(FormatExecution(result), GetRewriteNotice(plan));
        if (!string.IsNullOrWhiteSpace(perceptionNote))
        {
            reply = $"{reply} {perceptionNote}";
        }
        if (!string.IsNullOrWhiteSpace(recoveryNote))
        {
            reply = $"{reply} {recoveryNote}";
        }

        var lines = ExecutionToLines(result).ToList();
        var orderFillReport = BuildOrderFillExecutionReport(source, plan, result);
        if (orderFillReport != null)
        {
            reply = $"{reply} {orderFillReport.Summary}";
            lines.AddRange(orderFillReport.Lines);
        }

        return new WebIntentResponse(reply, false, null, null, lines, PlanToJson(plan), GetModeLabel(plan));
    }

    private async Task<WebLlmStatus> GetLlmStatusAsync(CancellationToken cancellationToken)
    {
        var provider = (_config.LlmFallback.Provider ?? "ollama").Trim();
        var endpoint = _config.LlmFallback.Endpoint;
        if (!_config.LlmFallbackEnabled)
        {
            return new WebLlmStatus(false, false, provider, "Disabled", endpoint);
        }

        if (!Uri.TryCreate(endpoint, UriKind.Absolute, out var uri))
        {
            return new WebLlmStatus(true, false, provider, "Invalid endpoint.", endpoint);
        }

        if (!uri.IsLoopback && !_config.AllowNonLoopbackLlmEndpoint)
        {
            return new WebLlmStatus(true, false, provider, "Endpoint must be local (loopback).", endpoint);
        }

        try
        {
            using var client = new HttpClient { Timeout = _httpTimeout };
            using var probeCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            probeCts.CancelAfter(_httpTimeout);

            if (provider.Equals("ollama", StringComparison.OrdinalIgnoreCase))
            {
                return await ProbeOllamaStatusAsync(client, uri, provider, endpoint, probeCts.Token);
            }

            var probeUri = BuildLlmProbeUri(uri, provider);
            using var response = await client.GetAsync(probeUri, probeCts.Token);
            if (response.IsSuccessStatusCode)
            {
                return new WebLlmStatus(true, true, provider, "Reachable", endpoint);
            }

            return new WebLlmStatus(true, false, provider, $"HTTP {(int)response.StatusCode}", endpoint);
        }
        catch (Exception ex)
        {
            return new WebLlmStatus(true, false, provider, Compact(ex.Message, 120), endpoint);
        }
    }

    private async Task<WebLlmStatus> ProbeOllamaStatusAsync(HttpClient client, Uri endpointUri, string provider, string? endpoint, CancellationToken cancellationToken)
    {
        var baseUri = new Uri(endpointUri.GetLeftPart(UriPartial.Authority));
        var tagsUri = new Uri(baseUri, "/api/tags");
        using var response = await client.GetAsync(tagsUri, cancellationToken);
        if (!response.IsSuccessStatusCode)
        {
            return new WebLlmStatus(true, false, provider, $"HTTP {(int)response.StatusCode}", endpoint);
        }

        var configuredModel = (_config.LlmFallback.Model ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(configuredModel))
        {
            return new WebLlmStatus(true, true, provider, "Reachable (model not configured)", endpoint);
        }

        var installedModels = await ReadOllamaModelNamesAsync(response, cancellationToken);
        if (installedModels.Count == 0)
        {
            return new WebLlmStatus(true, false, provider, "Reachable but unable to read model list", endpoint);
        }

        var found = installedModels.Any(name => string.Equals(name, configuredModel, StringComparison.OrdinalIgnoreCase));
        if (found)
        {
            return new WebLlmStatus(true, true, provider, $"Reachable; model '{configuredModel}' found", endpoint);
        }

        var preview = string.Join(", ", installedModels.Take(5));
        if (installedModels.Count > 5)
        {
            preview += ", ...";
        }

        var message = string.IsNullOrWhiteSpace(preview)
            ? $"Reachable but model '{configuredModel}' not found"
            : $"Reachable but model '{configuredModel}' not found. Installed: {preview}";
        return new WebLlmStatus(true, false, provider, Compact(message, 220), endpoint);
    }

    private static async Task<List<string>> ReadOllamaModelNamesAsync(HttpResponseMessage response, CancellationToken cancellationToken)
    {
        try
        {
            await using var stream = await response.Content.ReadAsStreamAsync(cancellationToken);
            using var document = await JsonDocument.ParseAsync(stream, cancellationToken: cancellationToken);
            if (!document.RootElement.TryGetProperty("models", out var models) || models.ValueKind != JsonValueKind.Array)
            {
                return new List<string>();
            }

            var names = new List<string>();
            foreach (var item in models.EnumerateArray())
            {
                if (item.TryGetProperty("name", out var nameProp))
                {
                    var value = nameProp.GetString();
                    if (!string.IsNullOrWhiteSpace(value))
                    {
                        names.Add(value.Trim());
                    }
                }
            }

            return names;
        }
        catch
        {
            return new List<string>();
        }
    }

    private static Uri BuildLlmProbeUri(Uri endpoint, string provider)
    {
        if (provider.Equals("ollama", StringComparison.OrdinalIgnoreCase))
        {
            var baseUri = new Uri(endpoint.GetLeftPart(UriPartial.Authority));
            return new Uri(baseUri, "/api/tags");
        }

        return endpoint;
    }

    private string ValidateAndApplyContextLock(ActionPlan plan, WindowRef? activeWindow)
    {
        ContextLockState lockState;
        lock (_sync)
        {
            lockState = _contextLock;
        }

        if (!lockState.Enabled)
        {
            return string.Empty;
        }

        if (activeWindow == null)
        {
            return $"context lock active ({FormatContextLock(lockState)}), but no active window.";
        }

        if (!string.IsNullOrWhiteSpace(lockState.WindowId)
            && !string.Equals(lockState.WindowId, activeWindow.Id, StringComparison.OrdinalIgnoreCase))
        {
            return $"context lock active ({FormatContextLock(lockState)}).";
        }

        if (!string.IsNullOrWhiteSpace(lockState.AppId)
            && !string.Equals(lockState.AppId, activeWindow.AppId, StringComparison.OrdinalIgnoreCase))
        {
            return $"context lock active ({FormatContextLock(lockState)}).";
        }

        foreach (var step in plan.Steps)
        {
            if (string.IsNullOrWhiteSpace(step.ExpectedWindowId) && !string.IsNullOrWhiteSpace(lockState.WindowId))
            {
                step.ExpectedWindowId = lockState.WindowId;
            }

            if (string.IsNullOrWhiteSpace(step.ExpectedAppId) && !string.IsNullOrWhiteSpace(lockState.AppId))
            {
                step.ExpectedAppId = lockState.AppId;
            }
        }

        return string.Empty;
    }

    private async Task<WebChatResponse> AddGoalAsync(string goalText, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(goalText))
        {
            return WebChatResponse.Simple("Goal text is required.");
        }

        GoalState goal;
        lock (_sync)
        {
            var now = DateTimeOffset.UtcNow;
            goal = new GoalState
            {
                Id = Guid.NewGuid().ToString("N")[..8],
                Text = goalText.Trim(),
                Completed = false,
                CreatedAtUtc = now,
                UpdatedAtUtc = now,
                LastRunAtUtc = now,
                Attempts = 0,
                LastResult = null,
                Priority = GoalPriorityNormal,
                AutoRunEnabled = true
            };
            _goals.Add(goal);
        }

        await SaveGoalsToDiskAsync(cancellationToken);
        return WebChatResponse.Simple($"Goal added [{goal.Id}]: {goal.Text}");
    }

    private async Task<WebChatResponse> MarkGoalDoneAsync(string key, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(key))
        {
            return WebChatResponse.Simple("Goal id or text is required.");
        }

        GoalState? updated = null;
        lock (_sync)
        {
            var index = FindGoalIndexLocked(key);
            if (index >= 0)
            {
                var goal = _goals[index];
                goal.Completed = true;
                goal.UpdatedAtUtc = DateTimeOffset.UtcNow;
                goal.LastResult = "Marked as done by user.";
                updated = goal;
            }
        }

        if (updated == null)
        {
            return WebChatResponse.Simple("Goal not found.");
        }

        await SaveGoalsToDiskAsync(cancellationToken);
        return WebChatResponse.Simple($"Goal [{updated.Id}] marked as done.");
    }

    private async Task<WebChatResponse> RemoveGoalAsync(string key, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(key))
        {
            return WebChatResponse.Simple("Goal id or text is required.");
        }

        GoalState? removed = null;
        lock (_sync)
        {
            var index = FindGoalIndexLocked(key);
            if (index >= 0)
            {
                removed = _goals[index];
                _goals.RemoveAt(index);
            }
        }

        if (removed == null)
        {
            return WebChatResponse.Simple("Goal not found.");
        }

        await SaveGoalsToDiskAsync(cancellationToken);
        return WebChatResponse.Simple($"Goal [{removed.Id}] removed.");
    }

    private async Task<WebChatResponse> SetGoalPriorityAsync(string payload, CancellationToken cancellationToken)
    {
        var parts = payload.Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (parts.Length < 2)
        {
            return WebChatResponse.Simple("Use: goal priority <id|text> <low|normal|high>");
        }

        var priorityToken = parts[^1];
        var key = string.Join(' ', parts.Take(parts.Length - 1));
        if (!TryParseGoalPriority(priorityToken, out var priority))
        {
            return WebChatResponse.Simple("Priority must be: low, normal, or high.");
        }

        GoalState? updated = null;
        lock (_sync)
        {
            var index = FindGoalIndexLocked(key);
            if (index >= 0)
            {
                var goal = _goals[index];
                goal.Priority = priority;
                goal.UpdatedAtUtc = DateTimeOffset.UtcNow;
                updated = goal;
            }
        }

        if (updated == null)
        {
            return WebChatResponse.Simple("Goal not found.");
        }

        await SaveGoalsToDiskAsync(cancellationToken);
        return WebChatResponse.Simple($"Goal [{updated.Id}] priority set to {FormatGoalPriority(updated.Priority)}.");
    }

    private async Task<WebChatResponse> SetGoalAutoModeAsync(string payload, CancellationToken cancellationToken)
    {
        var parts = payload.Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (parts.Length < 2)
        {
            return WebChatResponse.Simple("Use: goal auto <id|text> <on|off>");
        }

        var modeToken = parts[^1].Trim().ToLowerInvariant();
        bool enabled;
        if (modeToken is "on" or "enable" or "enabled" or "true")
        {
            enabled = true;
        }
        else if (modeToken is "off" or "disable" or "disabled" or "false")
        {
            enabled = false;
        }
        else
        {
            return WebChatResponse.Simple("Auto mode must be on or off.");
        }

        var key = string.Join(' ', parts.Take(parts.Length - 1));
        GoalState? updated = null;
        lock (_sync)
        {
            var index = FindGoalIndexLocked(key);
            if (index >= 0)
            {
                var goal = _goals[index];
                goal.AutoRunEnabled = enabled;
                goal.UpdatedAtUtc = DateTimeOffset.UtcNow;
                updated = goal;
            }
        }

        if (updated == null)
        {
            return WebChatResponse.Simple("Goal not found.");
        }

        await SaveGoalsToDiskAsync(cancellationToken);
        return WebChatResponse.Simple($"Goal [{updated.Id}] auto mode {(updated.AutoRunEnabled ? "ON" : "OFF")}.");
    }

    private Task<WebChatResponse> BuildGoalPlanConfirmationAsync(string key, CancellationToken cancellationToken)
    {
        GoalState? goal;
        lock (_sync)
        {
            var index = FindGoalIndexLocked(key);
            goal = index >= 0 ? _goals[index] : null;
        }

        if (goal == null)
        {
            return Task.FromResult(WebChatResponse.Simple("Goal not found."));
        }

        if (goal.Completed)
        {
            return Task.FromResult(WebChatResponse.Simple($"Goal [{goal.Id}] is already completed."));
        }

        var plan = _planner.PlanFromIntent(goal.Text);
        if (IsUnrecognizedPlan(plan))
        {
            return Task.FromResult(WebChatResponse.Simple($"Goal [{goal.Id}] could not be mapped to a safe plan. Edit it or use a clearer command."));
        }

        var source = $"goal:{goal.Id} {goal.Text}";
        var token = CreatePendingAction(PendingActionType.ExecutePlan, source, plan, dryRun: false);
        var notice = GetRewriteNotice(plan);
        var prompt = string.IsNullOrWhiteSpace(notice)
            ? $"Goal [{goal.Id}] interpreted. Confirm execution?"
            : $"Goal [{goal.Id}] interpreted. {notice}. Confirm execution?";
        return Task.FromResult(WebChatResponse.ConfirmWithSteps(prompt, token, PlanToLines(plan), PlanToJson(plan), GetModeLabel(plan)));
    }

    private async Task<WebChatResponse> ContinueLatestGoalAsync(CancellationToken cancellationToken)
    {
        GoalState? goal;
        lock (_sync)
        {
            goal = _goals
                .Where(item => !item.Completed)
                .OrderByDescending(item => item.Priority)
                .ThenByDescending(item => item.UpdatedAtUtc)
                .FirstOrDefault();
        }

        if (goal == null)
        {
            return WebChatResponse.Simple("No open goals. Use `goal add <text>` first.");
        }

        return await BuildGoalPlanConfirmationAsync(goal.Id, cancellationToken);
    }

    private string FormatGoals()
    {
        List<GoalState> snapshot;
        lock (_sync)
        {
            snapshot = _goals
                .OrderBy(item => item.Completed)
                .ThenByDescending(item => item.Priority)
                .ThenByDescending(item => item.UpdatedAtUtc)
                .Take(12)
                .ToList();
        }

        if (snapshot.Count == 0)
        {
            return "No goals yet. Use `goal add <text>`.";
        }

        var header = $"Goal scheduler: {(_config.GoalSchedulerEnabled ? "on" : "off")} every {_config.GoalSchedulerIntervalSeconds}s";
        var lines = snapshot.Select(goal =>
        {
            var status = goal.Completed ? "done" : "open";
            var attempts = goal.Attempts > 0 ? $" attempts:{goal.Attempts}" : string.Empty;
            var priority = FormatGoalPriority(goal.Priority);
            var auto = goal.AutoRunEnabled ? "auto:on" : "auto:off";
            return $"[{goal.Id}] {status} [{priority}] {auto} - {goal.Text}{attempts}";
        });
        return $"{header}{Environment.NewLine}{string.Join(Environment.NewLine, lines)}";
    }

    private string FormatMemory()
    {
        List<IntentMemoryEntry> snapshot;
        lock (_sync)
        {
            snapshot = _intentMemory
                .OrderByDescending(item => item.TimestampUtc)
                .Take(10)
                .ToList();
        }

        if (snapshot.Count == 0)
        {
            return "Memory is empty.";
        }

        var lines = snapshot.Select(item =>
        {
            var status = item.Success ? "ok" : "fail";
            var perception = string.IsNullOrWhiteSpace(item.AfterWindow)
                ? string.Empty
                : $" | window:{item.AfterWindow}";
            return $"{item.TimestampUtc:HH:mm:ss} [{status}] {item.IntentSummary}{perception}";
        });
        return string.Join(Environment.NewLine, lines);
    }

    private async Task RecordIntentMemoryAsync(
        string source,
        ActionPlan plan,
        ExecutionResult result,
        ObservationSnapshot before,
        ObservationSnapshot after,
        CancellationToken cancellationToken)
    {
        IntentMemoryEntry entry;
        lock (_sync)
        {
            var summary = source.Trim();
            if (summary.Length > 180)
            {
                summary = $"{summary[..180]}...";
            }

            entry = new IntentMemoryEntry(
                TimestampUtc: DateTimeOffset.UtcNow,
                Source: source,
                IntentSummary: summary,
                Success: result.Success,
                PlanSteps: plan.Steps.Count,
                ResultMessage: Compact(result.Message, 160),
                BeforeWindow: before.WindowDisplay,
                AfterWindow: after.WindowDisplay);

            _intentMemory.Add(entry);
            if (_intentMemory.Count > 200)
            {
                _intentMemory.RemoveRange(0, _intentMemory.Count - 200);
            }
        }

        await SaveMemoryToDiskAsync(cancellationToken);
    }

    private async Task UpdateGoalProgressAsync(string source, ExecutionResult result, CancellationToken cancellationToken)
    {
        if (!TryGetGoalIdFromSource(source, out var goalId))
        {
            return;
        }

        lock (_sync)
        {
            var index = _goals.FindIndex(goal => string.Equals(goal.Id, goalId, StringComparison.OrdinalIgnoreCase));
            if (index < 0)
            {
                return;
            }

            var current = _goals[index];
            current.UpdatedAtUtc = DateTimeOffset.UtcNow;
            current.LastRunAtUtc = DateTimeOffset.UtcNow;
            current.Attempts += 1;
            current.LastResult = Compact(result.Message, 180);
            current.Completed = current.Completed || result.Success;
        }

        await SaveGoalsToDiskAsync(cancellationToken);
    }

    private static bool TryGetGoalIdFromSource(string source, out string goalId)
    {
        goalId = string.Empty;
        if (!source.StartsWith("goal:", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        var payload = source["goal:".Length..].Trim();
        if (string.IsNullOrWhiteSpace(payload))
        {
            return false;
        }

        goalId = payload.Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)[0];
        return !string.IsNullOrWhiteSpace(goalId);
    }

    private static string BuildPerceptionNote(ObservationSnapshot before, ObservationSnapshot after)
    {
        if (string.IsNullOrWhiteSpace(before.WindowDisplay) && string.IsNullOrWhiteSpace(after.WindowDisplay))
        {
            return string.Empty;
        }

        var beforeText = string.IsNullOrWhiteSpace(before.WindowDisplay) ? "n/a" : before.WindowDisplay;
        var afterText = string.IsNullOrWhiteSpace(after.WindowDisplay) ? "n/a" : after.WindowDisplay;
        if (string.Equals(beforeText, afterText, StringComparison.OrdinalIgnoreCase))
        {
            return $"Perception: {afterText}.";
        }

        return $"Perception: {beforeText} -> {afterText}.";
    }

    private async Task<RecoveryAttemptOutcome> TryAutoRecoverAsync(
        ActionPlan originalPlan,
        ExecutionResult initialResult,
        Executor executor,
        CancellationToken cancellationToken)
    {
        if (initialResult.Success)
        {
            return RecoveryAttemptOutcome.None;
        }

        for (var attempt = 0; attempt < _config.AutoRecoveryMaxAttempts; attempt++)
        {
            if (!TryBuildRecoveryPlan(originalPlan, initialResult, attempt, out var recoveryPlan, out var reason))
            {
                return RecoveryAttemptOutcome.None;
            }

            await WriteAuditAsync(
                "auto_recovery_attempt",
                $"Auto-recovery attempt {attempt + 1}: {reason}",
                cancellationToken,
                new { attempt = attempt + 1, reason, steps = recoveryPlan.Steps.Count });

            var recovered = await executor.ExecutePlanAsync(recoveryPlan, dryRun: false, cancellationToken);
            if (recovered.Success)
            {
                return new RecoveryAttemptOutcome(recovered, $"Auto-recovery succeeded ({attempt + 1}/{_config.AutoRecoveryMaxAttempts}).");
            }
        }

        return RecoveryAttemptOutcome.None;
    }

    private bool TryBuildRecoveryPlan(
        ActionPlan originalPlan,
        ExecutionResult failedResult,
        int attempt,
        out ActionPlan recoveryPlan,
        out string reason)
    {
        recoveryPlan = new ActionPlan();
        reason = string.Empty;

        var failedStep = failedResult.Steps.FirstOrDefault(step => !step.Success);
        if (failedStep == null || failedStep.Index < 0 || failedStep.Index >= originalPlan.Steps.Count)
        {
            return false;
        }

        var originalFailedStep = originalPlan.Steps[failedStep.Index];
        if (originalFailedStep.Type is not (ActionType.Find or ActionType.Click))
        {
            return false;
        }

        if (!IsRecoverableUiFailureMessage(failedStep.Message))
        {
            return false;
        }

        var steps = new List<PlanStep>();
        if (!string.IsNullOrWhiteSpace(originalFailedStep.ExpectedAppId))
        {
            steps.Add(new PlanStep { Type = ActionType.OpenApp, AppIdOrPath = originalFailedStep.ExpectedAppId });
        }

        var waitMs = Math.Clamp(_config.AutoRecoveryWaitMs * (attempt + 1), 100, 5000);
        steps.Add(new PlanStep { Type = ActionType.WaitFor, WaitFor = TimeSpan.FromMilliseconds(waitMs) });

        for (var i = failedStep.Index; i < originalPlan.Steps.Count; i++)
        {
            steps.Add(ClonePlanStep(originalPlan.Steps[i]));
        }

        var firstAction = steps.FirstOrDefault(step => step.Type is ActionType.Find or ActionType.Click);
        if (firstAction != null)
        {
            if (firstAction.Type == ActionType.Click && firstAction.Selector == null && !string.IsNullOrWhiteSpace(firstAction.Text))
            {
                firstAction.Selector = new Selector { NameContains = firstAction.Text };
            }

            var selector = firstAction.Selector;
            if (selector != null && !string.IsNullOrWhiteSpace(selector.NameContains))
            {
                var tokens = selector.NameContains
                    .Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                    .Where(token => token.Length >= 3)
                    .ToList();

                if (tokens.Count > 0)
                {
                    selector.NameContains = tokens[0];
                }
            }
        }

        reason = $"recovering {originalFailedStep.Type} from step {failedStep.Index + 1}";
        recoveryPlan = new ActionPlan
        {
            Intent = $"{originalPlan.Intent} [auto-recovery]",
            Steps = steps
        };
        return true;
    }

    private static bool IsRecoverableUiFailureMessage(string? message)
    {
        if (string.IsNullOrWhiteSpace(message))
        {
            return false;
        }

        return message.Contains("No target to click", StringComparison.OrdinalIgnoreCase)
            || message.Contains("No elements found", StringComparison.OrdinalIgnoreCase)
            || message.Contains("not found", StringComparison.OrdinalIgnoreCase);
    }

    private static ExecutionResult MergeExecutionResults(ExecutionResult primary, ExecutionResult recovered)
    {
        var merged = new ExecutionResult
        {
            Success = recovered.Success,
            Message = recovered.Success ? $"Recovered: {recovered.Message}" : primary.Message
        };

        merged.Steps.AddRange(primary.Steps);
        for (var i = 0; i < recovered.Steps.Count; i++)
        {
            var step = recovered.Steps[i];
            merged.Steps.Add(new StepResult
            {
                Index = primary.Steps.Count + i,
                Type = step.Type,
                Success = step.Success,
                Message = $"[recovery] {step.Message}",
                Data = step.Data
            });
        }

        return merged;
    }

    private static PlanStep ClonePlanStep(PlanStep step)
    {
        return new PlanStep
        {
            Type = step.Type,
            Selector = step.Selector == null
                ? null
                : new Selector
                {
                    Role = step.Selector.Role,
                    NameContains = step.Selector.NameContains,
                    AutomationId = step.Selector.AutomationId,
                    ClassName = step.Selector.ClassName,
                    AncestorNameContains = step.Selector.AncestorNameContains,
                    Index = step.Selector.Index,
                    WindowId = step.Selector.WindowId,
                    BoundsHint = step.Selector.BoundsHint == null
                        ? null
                        : new Rect
                        {
                            X = step.Selector.BoundsHint.X,
                            Y = step.Selector.BoundsHint.Y,
                            Width = step.Selector.BoundsHint.Width,
                            Height = step.Selector.BoundsHint.Height
                        }
                },
            ExpectedAppId = step.ExpectedAppId,
            ExpectedWindowId = step.ExpectedWindowId,
            Text = step.Text,
            Target = step.Target,
            AppIdOrPath = step.AppIdOrPath,
            Keys = step.Keys == null ? null : new List<string>(step.Keys),
            Point = step.Point == null
                ? null
                : new Rect
                {
                    X = step.Point.X,
                    Y = step.Point.Y,
                    Width = step.Point.Width,
                    Height = step.Point.Height
                },
            ElementId = step.ElementId,
            WaitFor = step.WaitFor,
            Note = step.Note
        };
    }

    private async Task<ObservationSnapshot> CaptureObservationAsync(CancellationToken cancellationToken)
    {
        try
        {
            var snapshot = await _contextProvider.GetSnapshotAsync(cancellationToken);
            var window = snapshot.ActiveWindow;
            if (window == null)
            {
                return ObservationSnapshot.Empty;
            }

            var title = string.IsNullOrWhiteSpace(window.Title) ? "<untitled>" : window.Title.Trim();
            var appId = string.IsNullOrWhiteSpace(window.AppId) ? "<unknown-app>" : window.AppId.Trim();
            return new ObservationSnapshot(window.Id ?? string.Empty, $"{appId} - {Compact(title, 80)}");
        }
        catch
        {
            return ObservationSnapshot.Empty;
        }
    }

    private int FindGoalIndexLocked(string key)
    {
        if (string.IsNullOrWhiteSpace(key))
        {
            return -1;
        }

        var trimmed = key.Trim();
        var byId = _goals.FindIndex(goal => string.Equals(goal.Id, trimmed, StringComparison.OrdinalIgnoreCase));
        if (byId >= 0)
        {
            return byId;
        }

        return _goals.FindIndex(goal => goal.Text.Contains(trimmed, StringComparison.OrdinalIgnoreCase));
    }

    private static bool TryParseGoalPriority(string token, out int priority)
    {
        priority = GoalPriorityNormal;
        var normalized = token.Trim().ToLowerInvariant();
        switch (normalized)
        {
            case "low":
            case "l":
                priority = GoalPriorityLow;
                return true;
            case "normal":
            case "medium":
            case "med":
            case "n":
                priority = GoalPriorityNormal;
                return true;
            case "high":
            case "h":
                priority = GoalPriorityHigh;
                return true;
            default:
                return false;
        }
    }

    private static string FormatGoalPriority(int priority)
    {
        return priority switch
        {
            <= GoalPriorityLow => "low",
            >= GoalPriorityHigh => "high",
            _ => "normal"
        };
    }

    private static string FormatExecution(ExecutionResult result)
    {
        var summary = $"Success: {result.Success}. {result.Message}";
        if (result.Steps.Count == 0)
        {
            return summary;
        }

        var details = string.Join(" ", result.Steps.Select(step => $"{step.Type}:{step.Success}"));
        return $"{summary} Completed Steps: {details}";
    }

    private static string AppendNotice(string reply, string? notice)
        => string.IsNullOrWhiteSpace(notice) ? reply : $"{reply} {notice}";

    private static string? GetRewriteNotice(ActionPlan plan)
    {
        var raw = plan.Steps.Select(step => step.Note)
            .FirstOrDefault(note => !string.IsNullOrWhiteSpace(note) && note.StartsWith("Rewritten intent:", StringComparison.OrdinalIgnoreCase));
        if (string.IsNullOrWhiteSpace(raw))
        {
            return null;
        }

        var cleaned = Regex.Replace(raw, @"\s*\|\s*llm-low-confidence:[0-9.]+", string.Empty, RegexOptions.IgnoreCase);
        cleaned = Regex.Replace(cleaned, @"\s*\|\s*llm-needs-clarification", string.Empty, RegexOptions.IgnoreCase);
        cleaned = Regex.Replace(cleaned, @"\s+\|\s+", " | ");
        cleaned = cleaned.Trim();

        var lowMatch = Regex.Match(raw, @"llm-low-confidence:(?<score>[0-9.]+)", RegexOptions.IgnoreCase);
        if (lowMatch.Success)
        {
            cleaned = $"{cleaned} | Low confidence ({lowMatch.Groups["score"].Value})";
        }

        if (raw.Contains("llm-needs-clarification", StringComparison.OrdinalIgnoreCase))
        {
            cleaned = $"{cleaned} | Clarification suggested";
        }

        return cleaned;
    }

    private static string GetModeLabel(ActionPlan plan)
        => string.IsNullOrWhiteSpace(GetRewriteNotice(plan)) ? "Mode: Rule-based" : "Mode: LLM interpreter";

    private static IReadOnlyList<string> PlanToLines(ActionPlan plan)
    {
        var lines = new List<string>();
        for (var i = 0; i < plan.Steps.Count; i++)
        {
            var step = plan.Steps[i];
            lines.Add($"{i + 1}. {step.Type}");
        }
        return lines;
    }

    private static IReadOnlyList<string> ExecutionToLines(ExecutionResult result)
    {
        var lines = new List<string>();
        for (var i = 0; i < result.Steps.Count; i++)
        {
            var step = result.Steps[i];
            var line = $"{i + 1}. {step.Type} => {step.Success} ({step.Message})";
            if (step.Data != null)
            {
                line = $"{line} data={ToInlineJson(step.Data, 280)}";
            }
            lines.Add(line);
        }
        return lines;
    }

    private static OrderFillExecutionReport? BuildOrderFillExecutionReport(string source, ActionPlan plan, ExecutionResult result)
    {
        if (!source.StartsWith("order-fill:", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        var fieldToStepIndexes = new Dictionary<string, List<int>>(StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < plan.Steps.Count; i++)
        {
            if (!TryGetFillFieldKey(plan.Steps[i], out var fieldKey))
            {
                continue;
            }

            if (!fieldToStepIndexes.TryGetValue(fieldKey, out var indexes))
            {
                indexes = new List<int>();
                fieldToStepIndexes[fieldKey] = indexes;
            }

            indexes.Add(i);
        }

        if (fieldToStepIndexes.Count == 0)
        {
            return null;
        }

        var resultByIndex = result.Steps
            .GroupBy(step => step.Index)
            .ToDictionary(group => group.Key, group => group.ToList());

        var filled = new List<string>();
        var missing = new List<string>();
        foreach (var field in fieldToStepIndexes.OrderBy(static item => item.Value.Min()))
        {
            var hasFilledStep = field.Value.Any(index =>
                resultByIndex.TryGetValue(index, out var steps)
                && steps.Any(IsEffectiveFillStepSuccess));

            var label = GetOrderFieldDisplayName(field.Key);
            if (hasFilledStep)
            {
                filled.Add(label);
            }
            else
            {
                missing.Add(label);
            }
        }

        var total = filled.Count + missing.Count;
        var summary = total == 0
            ? "Form fill report unavailable."
            : $"Form fill: {filled.Count}/{total} fields filled."
              + (missing.Count > 0 ? $" Missing: {string.Join(", ", missing)}." : string.Empty);

        var lines = new List<string>
        {
            $"Order fill report: {filled.Count}/{total} fields filled"
        };
        if (filled.Count > 0)
        {
            lines.Add($"Filled: {string.Join(", ", filled)}");
        }
        if (missing.Count > 0)
        {
            lines.Add($"Missing: {string.Join(", ", missing)}");
        }

        return new OrderFillExecutionReport(summary, lines);
    }

    private static bool IsEffectiveFillStepSuccess(StepResult step)
    {
        if (!step.Success)
        {
            return false;
        }

        if (step.Message.Contains("Optional step skipped", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        return step.Type is ActionType.SetValue or ActionType.TypeText;
    }

    private static bool TryGetFillFieldKey(PlanStep step, out string key)
    {
        key = string.Empty;
        var note = step.Note ?? string.Empty;
        if (string.IsNullOrWhiteSpace(note))
        {
            return false;
        }

        var parts = note.Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        foreach (var part in parts)
        {
            if (!part.StartsWith("fill-field:", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var value = part["fill-field:".Length..].Trim();
            if (string.IsNullOrWhiteSpace(value))
            {
                continue;
            }

            key = value;
            return true;
        }

        // Backward compatibility with older plans that only used optional groups.
        foreach (var part in parts)
        {
            if (!part.StartsWith("optional-group:", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var value = part["optional-group:".Length..].Trim();
            if (string.IsNullOrWhiteSpace(value)
                || value.Equals("first_focus", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            key = value;
            return true;
        }

        return false;
    }

    private static string GetOrderFieldDisplayName(string key)
    {
        return key.ToLowerInvariant() switch
        {
            "customer_name" => "Name",
            "customer_email" => "Email",
            "customer_phone" => "Phone",
            "order_number" => "Order number",
            "shipping_address" => "Shipping address",
            "billing_address" => "Billing address",
            "order_notes" => "Comments",
            _ => key.Replace('_', ' ')
        };
    }

    private static string ToInlineJson(object data, int maxChars)
    {
        try
        {
            var json = JsonSerializer.Serialize(data);
            return Compact(json, maxChars);
        }
        catch
        {
            return Compact(data.ToString() ?? "<data>", maxChars);
        }
    }

    private static string PlanToJson(ActionPlan plan)
    {
        try
        {
            return JsonSerializer.Serialize(plan, JsonOptions);
        }
        catch
        {
            return string.Empty;
        }
    }
    private static bool TryParseActionPlanJson(string? planJson, out ActionPlan? plan, out string error)
    {
        plan = null;
        error = string.Empty;
        var json = planJson?.Trim();
        if (string.IsNullOrWhiteSpace(json))
        {
            error = "Empty plan JSON.";
            return false;
        }

        try
        {
            plan = JsonSerializer.Deserialize<ActionPlan>(json, JsonOptions);
            if (plan == null || plan.Steps.Count == 0)
            {
                error = "Plan is empty.";
                plan = null;
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

    private async Task<WebChatResponse> TranslateWithLlmAsync(TranslationIntent intent, CancellationToken cancellationToken)
    {
        if (!_config.LlmFallbackEnabled)
        {
            return WebChatResponse.Simple("LLM is disabled. Enable it in config to use translation.");
        }

        var endpoint = _config.LlmFallback.Endpoint?.Trim();
        if (string.IsNullOrWhiteSpace(endpoint) || !Uri.TryCreate(endpoint, UriKind.Absolute, out var uri))
        {
            return WebChatResponse.Simple("LLM endpoint is not configured or invalid.");
        }

        if (!uri.IsLoopback && !_config.AllowNonLoopbackLlmEndpoint)
        {
            return WebChatResponse.Simple("LLM endpoint must be local unless remote LLM is enabled.");
        }

        var provider = (_config.LlmFallback.Provider ?? "ollama").Trim().ToLowerInvariant();
        uri = NormalizeLlmEndpoint(uri, provider);
        var prompt = BuildTranslationPrompt(intent);
        var maxTokens = ResolveTranslationMaxTokens(intent.Text, _config.LlmFallback.MaxTokens);

        try
        {
            await WriteAuditAsync("llm_translate_request", "Translation requested", cancellationToken, new
            {
                provider,
                model = _config.LlmFallback.Model,
                endpoint,
                target = intent.TargetLanguage,
                source = intent.SourceLanguage,
                input = _config.AuditLlmIncludeRawText ? intent.Text : "[redacted]",
                inputLength = intent.Text.Length
            });

            using var client = new HttpClient { Timeout = _httpTimeout };
            var translated = provider switch
            {
                "openai" => await CallOpenAiTranslationAsync(client, uri, prompt, maxTokens, cancellationToken),
                "llama.cpp" => await CallLlamaCppTranslationAsync(client, uri, prompt, maxTokens, cancellationToken),
                _ => await CallOllamaTranslationAsync(client, uri, prompt, maxTokens, cancellationToken)
            };

            translated = CleanTranslationOutput(translated);
            if (string.IsNullOrWhiteSpace(translated))
            {
                return WebChatResponse.Simple("Translation failed: model returned an empty response.");
            }

            await WriteAuditAsync("llm_translate_response", "Translation completed", cancellationToken, new
            {
                provider,
                model = _config.LlmFallback.Model,
                output = _config.AuditLlmIncludeRawText ? translated : "[redacted]",
                outputLength = translated.Length
            });

            return WebChatResponse.Simple($"Translation ({provider}):\n{translated}");
        }
        catch (Exception ex)
        {
            await WriteAuditAsync("llm_translate_error", "Translation failed", cancellationToken, new
            {
                provider,
                model = _config.LlmFallback.Model,
                error = ex.Message
            });
            return WebChatResponse.Simple($"Translation failed: {Compact(ex.Message, 120)}");
        }
    }

    private async Task<string?> SuggestSupportedCommandAsync(string userInput, CancellationToken cancellationToken)
    {
        if (!_config.LlmFallbackEnabled || string.IsNullOrWhiteSpace(userInput))
        {
            return null;
        }

        var endpoint = _config.LlmFallback.Endpoint?.Trim();
        if (string.IsNullOrWhiteSpace(endpoint) || !Uri.TryCreate(endpoint, UriKind.Absolute, out var uri))
        {
            return null;
        }

        if (!uri.IsLoopback && !_config.AllowNonLoopbackLlmEndpoint)
        {
            return null;
        }

        var provider = (_config.LlmFallback.Provider ?? "ollama").Trim().ToLowerInvariant();
        uri = NormalizeLlmEndpoint(uri, provider);
        var prompt = BuildCommandSuggestionPrompt(userInput);
        var maxTokens = ResolveSuggestionMaxTokens(_config.LlmFallback.MaxTokens);

        try
        {
            using var client = new HttpClient { Timeout = _httpTimeout };
            var raw = provider switch
            {
                "openai" => await CallOpenAiTranslationAsync(client, uri, prompt, maxTokens, cancellationToken),
                "llama.cpp" => await CallLlamaCppTranslationAsync(client, uri, prompt, maxTokens, cancellationToken),
                _ => await CallOllamaTranslationAsync(client, uri, prompt, maxTokens, cancellationToken)
            };

            return CleanSuggestedCommand(raw);
        }
        catch (Exception ex)
        {
            await WriteAuditAsync("llm_command_suggestion_error", "AI suggestion failed", cancellationToken, new
            {
                provider,
                endpoint,
                error = ex.Message
            });
            return null;
        }
    }

    private async Task<string?> CallOllamaTranslationAsync(HttpClient client, Uri endpoint, string prompt, int maxTokens, CancellationToken cancellationToken)
    {
        var payload = new
        {
            model = _config.LlmFallback.Model,
            prompt,
            stream = false,
            options = new { temperature = 0.1, num_predict = maxTokens }
        };

        using var response = await client.PostAsJsonAsync(endpoint, payload, JsonOptions, cancellationToken);
        if (!response.IsSuccessStatusCode)
        {
            throw await BuildLlmHttpExceptionAsync(response, cancellationToken);
        }

        await using var stream = await response.Content.ReadAsStreamAsync(cancellationToken);
        using var doc = await JsonDocument.ParseAsync(stream, cancellationToken: cancellationToken);
        return doc.RootElement.TryGetProperty("response", out var result) ? result.GetString() : null;
    }

    private async Task<string?> CallOllamaPromptAsync(HttpClient client, Uri endpoint, string prompt, int maxTokens, CancellationToken cancellationToken)
    {
        var payload = new
        {
            model = _config.LlmFallback.Model,
            prompt,
            stream = false,
            options = new { temperature = 0.1, num_predict = maxTokens }
        };

        using var response = await client.PostAsJsonAsync(endpoint, payload, JsonOptions, cancellationToken);
        if (!response.IsSuccessStatusCode)
        {
            throw await BuildLlmHttpExceptionAsync(response, cancellationToken);
        }

        await using var stream = await response.Content.ReadAsStreamAsync(cancellationToken);
        using var doc = await JsonDocument.ParseAsync(stream, cancellationToken: cancellationToken);
        return doc.RootElement.TryGetProperty("response", out var result) ? result.GetString() : null;
    }

    private async Task<string?> CallOpenAiTranslationAsync(HttpClient client, Uri endpoint, string prompt, int maxTokens, CancellationToken cancellationToken)
    {
        var payload = new
        {
            model = _config.LlmFallback.Model,
            messages = new[]
            {
                new { role = "system", content = "You are a precise translator. Return only translated text." },
                new { role = "user", content = prompt }
            },
            temperature = 0.1,
            max_tokens = maxTokens
        };

        using var response = await client.PostAsJsonAsync(endpoint, payload, JsonOptions, cancellationToken);
        if (!response.IsSuccessStatusCode)
        {
            throw await BuildLlmHttpExceptionAsync(response, cancellationToken);
        }

        await using var stream = await response.Content.ReadAsStreamAsync(cancellationToken);
        using var doc = await JsonDocument.ParseAsync(stream, cancellationToken: cancellationToken);
        if (!doc.RootElement.TryGetProperty("choices", out var choices)
            || choices.ValueKind != JsonValueKind.Array
            || choices.GetArrayLength() == 0)
        {
            return null;
        }

        var first = choices[0];
        if (first.TryGetProperty("message", out var message) && message.TryGetProperty("content", out var content))
        {
            return content.GetString();
        }

        return first.TryGetProperty("text", out var text) ? text.GetString() : null;
    }

    private async Task<string?> CallOpenAiPromptAsync(HttpClient client, Uri endpoint, string systemPrompt, string prompt, int maxTokens, CancellationToken cancellationToken)
    {
        var payload = new
        {
            model = _config.LlmFallback.Model,
            messages = new[]
            {
                new { role = "system", content = systemPrompt },
                new { role = "user", content = prompt }
            },
            temperature = 0.1,
            max_tokens = maxTokens,
            response_format = new { type = "json_object" }
        };

        using var response = await client.PostAsJsonAsync(endpoint, payload, JsonOptions, cancellationToken);
        if (!response.IsSuccessStatusCode)
        {
            throw await BuildLlmHttpExceptionAsync(response, cancellationToken);
        }

        await using var stream = await response.Content.ReadAsStreamAsync(cancellationToken);
        using var doc = await JsonDocument.ParseAsync(stream, cancellationToken: cancellationToken);
        if (!doc.RootElement.TryGetProperty("choices", out var choices)
            || choices.ValueKind != JsonValueKind.Array
            || choices.GetArrayLength() == 0)
        {
            return null;
        }

        var first = choices[0];
        if (first.TryGetProperty("message", out var message) && message.TryGetProperty("content", out var content))
        {
            return content.GetString();
        }

        return first.TryGetProperty("text", out var text) ? text.GetString() : null;
    }

    private async Task<string?> CallLlamaCppTranslationAsync(HttpClient client, Uri endpoint, string prompt, int maxTokens, CancellationToken cancellationToken)
    {
        var payload = new
        {
            prompt,
            n_predict = maxTokens,
            temperature = 0.1
        };

        using var response = await client.PostAsJsonAsync(endpoint, payload, JsonOptions, cancellationToken);
        if (!response.IsSuccessStatusCode)
        {
            throw await BuildLlmHttpExceptionAsync(response, cancellationToken);
        }

        await using var stream = await response.Content.ReadAsStreamAsync(cancellationToken);
        using var doc = await JsonDocument.ParseAsync(stream, cancellationToken: cancellationToken);
        if (doc.RootElement.TryGetProperty("content", out var content))
        {
            return content.GetString();
        }

        if (!doc.RootElement.TryGetProperty("choices", out var choices)
            || choices.ValueKind != JsonValueKind.Array
            || choices.GetArrayLength() == 0)
        {
            return null;
        }

        var first = choices[0];
        return first.TryGetProperty("text", out var text) ? text.GetString() : null;
    }

    private async Task<string?> CallLlamaCppPromptAsync(HttpClient client, Uri endpoint, string prompt, int maxTokens, CancellationToken cancellationToken)
    {
        var payload = new
        {
            prompt,
            n_predict = maxTokens,
            temperature = 0.1
        };

        using var response = await client.PostAsJsonAsync(endpoint, payload, JsonOptions, cancellationToken);
        if (!response.IsSuccessStatusCode)
        {
            throw await BuildLlmHttpExceptionAsync(response, cancellationToken);
        }

        await using var stream = await response.Content.ReadAsStreamAsync(cancellationToken);
        using var doc = await JsonDocument.ParseAsync(stream, cancellationToken: cancellationToken);
        if (doc.RootElement.TryGetProperty("content", out var content))
        {
            return content.GetString();
        }

        if (!doc.RootElement.TryGetProperty("choices", out var choices)
            || choices.ValueKind != JsonValueKind.Array
            || choices.GetArrayLength() == 0)
        {
            return null;
        }

        var first = choices[0];
        return first.TryGetProperty("text", out var text) ? text.GetString() : null;
    }

    private static Uri NormalizeLlmEndpoint(Uri uri, string provider)
    {
        if (!string.IsNullOrWhiteSpace(uri.AbsolutePath) && uri.AbsolutePath != "/")
        {
            return uri;
        }

        return provider switch
        {
            "openai" => new Uri(uri, "/v1/chat/completions"),
            "llama.cpp" => new Uri(uri, "/completion"),
            _ => new Uri(uri, "/api/generate")
        };
    }

    private static int ResolveTranslationMaxTokens(string input, int configuredMaxTokens)
    {
        var configured = Math.Clamp(configuredMaxTokens, 32, 4096);
        var estimated = Math.Clamp((int)Math.Ceiling((input?.Length ?? 0) / 2.8), 128, 4096);
        return Math.Max(configured, estimated);
    }

    private static int ResolveSuggestionMaxTokens(int configuredMaxTokens)
    {
        var configured = Math.Clamp(configuredMaxTokens, 32, 512);
        return Math.Clamp(Math.Max(96, configured / 2), 96, 256);
    }

    private static async Task<Exception> BuildLlmHttpExceptionAsync(HttpResponseMessage response, CancellationToken cancellationToken)
    {
        var detail = await TryReadLlmErrorAsync(response, cancellationToken);
        var status = (int)response.StatusCode;
        var reason = response.ReasonPhrase ?? "HTTP error";
        return new InvalidOperationException($"LLM HTTP {status} ({reason}): {detail}");
    }

    private static async Task<string> TryReadLlmErrorAsync(HttpResponseMessage response, CancellationToken cancellationToken)
    {
        try
        {
            var body = await response.Content.ReadAsStringAsync(cancellationToken);
            if (string.IsNullOrWhiteSpace(body))
            {
                return "empty response body";
            }

            var trimmed = body.Trim();
            try
            {
                using var doc = JsonDocument.Parse(trimmed);
                if (doc.RootElement.ValueKind == JsonValueKind.Object)
                {
                    if (doc.RootElement.TryGetProperty("error", out var errorProp))
                    {
                        return Compact(errorProp.ToString(), 220);
                    }

                    if (doc.RootElement.TryGetProperty("message", out var messageProp))
                    {
                        return Compact(messageProp.ToString(), 220);
                    }
                }
            }
            catch
            {
                // Best effort: if it's not JSON, return compact text body.
            }

            return Compact(trimmed, 220);
        }
        catch
        {
            return "unable to read response body";
        }
    }

    private static string BuildTranslationPrompt(TranslationIntent intent)
    {
        var source = string.IsNullOrWhiteSpace(intent.SourceLanguage)
            ? string.Empty
            : $"Source language: {intent.SourceLanguage}\n";

        return
            "You are a translation engine.\n" +
            "Translate exactly, preserving tone and line breaks.\n" +
            "Return only the translated text.\n" +
            $"Target language: {intent.TargetLanguage}\n" +
            source +
            "TEXT START\n" +
            intent.Text +
            "\nTEXT END";
    }

    private static string BuildCommandSuggestionPrompt(string userInput)
    {
        return
            "You convert natural language into one safe DesktopAgent command.\n" +
            "Return exactly one command, no explanation.\n" +
            "If you cannot map safely, return UNKNOWN.\n" +
            "Allowed command patterns include:\n" +
            "- status | arm | disarm | simulate presence | require presence\n" +
            "- run <intent> | dry-run <intent>\n" +
            "- open <app>\n" +
            "- search <query> on <chrome|edge|firefox>\n" +
            "- translate <text> to <language>\n" +
            "- order intake [<email text>] | order preview | order clear | order fill <url>\n" +
            "- take screenshot\n" +
            "- take screenshot single-screen\n" +
            "- record screen [and audio] for <duration>\n" +
            "- jiggle mouse for <duration>\n" +
            "If user asks to process order details from an email/message and fill a form, map to 'order intake' first.\n" +
            "If user already provides form URL and asks to compile/fill it from order content, map to 'order fill <url>'.\n" +
            "User request:\n" +
            userInput;
    }

    private static string? CleanTranslationOutput(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw))
        {
            return null;
        }

        var text = raw.Trim();
        if (text.StartsWith("translation:", StringComparison.OrdinalIgnoreCase))
        {
            text = text["translation:".Length..].Trim();
        }

        if (text.StartsWith("```", StringComparison.Ordinal))
        {
            text = text.Trim('`').Trim();
        }

        return string.IsNullOrWhiteSpace(text) ? null : text;
    }

    private static string? CleanSuggestedCommand(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw))
        {
            return null;
        }

        var line = raw.Replace("\r", "\n")
            .Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .FirstOrDefault() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(line))
        {
            return null;
        }

        line = line.Trim().Trim('`', '"', '\'');
        if (line.StartsWith("command:", StringComparison.OrdinalIgnoreCase))
        {
            line = line["command:".Length..].Trim();
        }

        if (line.Equals("unknown", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        return line;
    }

    private static bool TryParseTranslationIntent(string message, out TranslationIntent intent)
    {
        intent = default;
        if (string.IsNullOrWhiteSpace(message))
        {
            return false;
        }

        var text = message.Trim();
        var lower = text.ToLowerInvariant();
        if (!lower.Contains("traduc", StringComparison.Ordinal)
            && !lower.Contains("tradur", StringComparison.Ordinal)
            && !lower.Contains("translat", StringComparison.Ordinal))
        {
            return false;
        }

        if (TryParseTranslationHeadBody(text, out intent))
        {
            return true;
        }

        if (TryParseInlineTranslation(text, out intent))
        {
            return true;
        }

        if (TryParseImplicitTranslation(text, out intent))
        {
            return true;
        }

        return false;
    }

    private static bool TryParseTranslationHeadBody(string text, out TranslationIntent intent)
    {
        intent = default;
        var newlineIndex = text.IndexOf('\n');
        if (newlineIndex > 0)
        {
            var head = text[..newlineIndex].Trim().TrimEnd(':');
            var body = text[(newlineIndex + 1)..].Trim();
            if (TryParseLanguageHead(head, out var target, out var source) && !string.IsNullOrWhiteSpace(body))
            {
                intent = new TranslationIntent(body, target, source);
                return true;
            }
        }

        var colonIndex = text.IndexOf(':');
        if (colonIndex > 0)
        {
            var head = text[..colonIndex].Trim();
            var body = text[(colonIndex + 1)..].Trim();
            if (TryParseLanguageHead(head, out var target, out var source) && !string.IsNullOrWhiteSpace(body))
            {
                intent = new TranslationIntent(body, target, source);
                return true;
            }
        }

        return false;
    }

    private static bool TryParseInlineTranslation(string text, out TranslationIntent intent)
    {
        intent = default;
        var lowered = text.ToLowerInvariant();
        var inMarker = lowered.LastIndexOf(" in ", StringComparison.Ordinal);
        var toMarker = lowered.LastIndexOf(" to ", StringComparison.Ordinal);
        var marker = Math.Max(inMarker, toMarker);
        if (marker <= 0)
        {
            return false;
        }

        var markerLen = marker == inMarker ? 4 : 4;
        var before = text[..marker].Trim();
        var after = text[(marker + markerLen)..].Trim();
        if (string.IsNullOrWhiteSpace(after))
        {
            return false;
        }

        var sep = after.IndexOfAny(new[] { '?', ':', '\n', ';', '!' });
        var descriptor = sep >= 0 ? after[..sep].Trim() : after;
        var postText = sep >= 0 ? after[(sep + 1)..].Trim() : string.Empty;
        if (!TryParseLanguageDescriptor(descriptor, out var target, out var source))
        {
            return false;
        }

        var textBody = !string.IsNullOrWhiteSpace(postText)
            ? postText
            : NormalizeTranslationLeadIn(before);
        if (string.IsNullOrWhiteSpace(textBody))
        {
            return false;
        }

        intent = new TranslationIntent(textBody, target, source);
        return true;
    }

    private static bool TryParseImplicitTranslation(string text, out TranslationIntent intent)
    {
        intent = default;
        if (string.IsNullOrWhiteSpace(text))
        {
            return false;
        }

        var separatorIndex = text.IndexOfAny(new[] { '?', ':', '\n' });
        if (separatorIndex <= 0 || separatorIndex >= text.Length - 1)
        {
            return false;
        }

        var leadIn = text[..separatorIndex].Trim();
        var body = text[(separatorIndex + 1)..].Trim();
        if (string.IsNullOrWhiteSpace(body))
        {
            return false;
        }

        var target = InferDefaultTargetLanguage(leadIn, body);
        intent = new TranslationIntent(body, target, null);
        return true;
    }

    private static string NormalizeTranslationLeadIn(string value)
    {
        var text = value.Trim();
        text = Regex.Replace(text, "^(can\\s+you\\s+)?(please\\s+)?translate\\s+", string.Empty, RegexOptions.IgnoreCase);
        text = Regex.Replace(text, "^(puoi\\s+)?(per\\s+favore\\s+)?tradur(?:re|mi|re\\s+mi)?\\s+", string.Empty, RegexOptions.IgnoreCase);
        text = Regex.Replace(text, "^(pui\\s+)?(per\\s+favore\\s+)?tradur(?:re|mi|re\\s+mi)?\\s+", string.Empty, RegexOptions.IgnoreCase);
        text = Regex.Replace(text, "^(traduci|tradurre)\\s+", string.Empty, RegexOptions.IgnoreCase);
        return text.Trim().Trim('?', ':', '.', '!', ',', ';');
    }

    private static bool TryParseLanguageHead(string head, out string target, out string? source)
    {
        target = string.Empty;
        source = null;
        if (string.IsNullOrWhiteSpace(head))
        {
            return false;
        }

        var trimmed = head.Trim();
        if (trimmed.StartsWith("translate to ", StringComparison.OrdinalIgnoreCase))
        {
            return TryParseLanguageDescriptor(trimmed["translate to ".Length..], out target, out source);
        }
        if (trimmed.StartsWith("translate in ", StringComparison.OrdinalIgnoreCase))
        {
            return TryParseLanguageDescriptor(trimmed["translate in ".Length..], out target, out source);
        }
        if (trimmed.StartsWith("traduci in ", StringComparison.OrdinalIgnoreCase))
        {
            return TryParseLanguageDescriptor(trimmed["traduci in ".Length..], out target, out source);
        }
        if (trimmed.StartsWith("tradurre in ", StringComparison.OrdinalIgnoreCase))
        {
            return TryParseLanguageDescriptor(trimmed["tradurre in ".Length..], out target, out source);
        }
        if (trimmed.StartsWith("to ", StringComparison.OrdinalIgnoreCase))
        {
            return TryParseLanguageDescriptor(trimmed["to ".Length..], out target, out source);
        }
        if (trimmed.StartsWith("in ", StringComparison.OrdinalIgnoreCase))
        {
            return TryParseLanguageDescriptor(trimmed["in ".Length..], out target, out source);
        }

        return false;
    }

    private static bool TryParseLanguageDescriptor(string descriptor, out string target, out string? source)
    {
        target = string.Empty;
        source = null;
        if (string.IsNullOrWhiteSpace(descriptor))
        {
            return false;
        }

        var normalized = descriptor.Trim().Trim('?', ':', '.', '!', ',', ';');
        var lower = normalized.ToLowerInvariant();
        var marker = lower.IndexOf(" from ", StringComparison.Ordinal);
        var markerLen = 6;
        if (marker <= 0)
        {
            marker = lower.IndexOf(" da ", StringComparison.Ordinal);
            markerLen = 4;
        }

        if (marker > 0)
        {
            target = normalized[..marker].Trim();
            source = normalized[(marker + markerLen)..].Trim();
        }
        else
        {
            target = normalized;
        }

        if (string.IsNullOrWhiteSpace(target))
        {
            return false;
        }

        if (string.IsNullOrWhiteSpace(source))
        {
            source = null;
        }

        return true;
    }

    private static string InferDefaultTargetLanguage(string leadIn, string body)
    {
        var lead = (leadIn ?? string.Empty).ToLowerInvariant();
        var text = (body ?? string.Empty).ToLowerInvariant();

        if (lead.Contains("italian", StringComparison.Ordinal) || lead.Contains("italiano", StringComparison.Ordinal))
        {
            return "italian";
        }

        if (lead.Contains("english", StringComparison.Ordinal) || lead.Contains("inglese", StringComparison.Ordinal))
        {
            return "english";
        }

        var italianHints = new[]
        {
            "ciao", "sono", "come va", "grazie", "per favore", "buongiorno", "arrivederci", "oggi", "domani", "prego"
        };
        if (italianHints.Any(hint => text.Contains(hint, StringComparison.Ordinal)))
        {
            return "english";
        }

        var englishHints = new[]
        {
            "hello", "how are you", "thanks", "please", "good morning", "today", "tomorrow"
        };
        if (englishHints.Any(hint => text.Contains(hint, StringComparison.Ordinal)))
        {
            return "italian";
        }

        return "english";
    }

    private readonly record struct TranslationIntent(string Text, string TargetLanguage, string? SourceLanguage);

    private static bool IsDirectIntent(string normalized)
    {
        if (normalized.StartsWith("http://", StringComparison.Ordinal) || normalized.StartsWith("https://", StringComparison.Ordinal))
        {
            return true;
        }

        return normalized.StartsWith("open ", StringComparison.Ordinal)
            || normalized.StartsWith("start ", StringComparison.Ordinal)
            || normalized.StartsWith("launch ", StringComparison.Ordinal)
            || normalized.StartsWith("apri ", StringComparison.Ordinal)
            || normalized.StartsWith("avvia ", StringComparison.Ordinal)
            || normalized.StartsWith("run ", StringComparison.Ordinal)
            || normalized.StartsWith("find ", StringComparison.Ordinal)
            || normalized.StartsWith("click ", StringComparison.Ordinal)
            || normalized.StartsWith("type ", StringComparison.Ordinal)
            || normalized.StartsWith("search ", StringComparison.Ordinal)
            || normalized.StartsWith("cerca ", StringComparison.Ordinal)
            || normalized.StartsWith("record ", StringComparison.Ordinal)
            || normalized.StartsWith("start recording", StringComparison.Ordinal)
            || normalized.StartsWith("stop recording", StringComparison.Ordinal)
            || normalized.StartsWith("snapshot", StringComparison.Ordinal)
            || normalized.StartsWith("screenshot", StringComparison.Ordinal)
            || normalized.StartsWith("take screenshot", StringComparison.Ordinal)
            || normalized.StartsWith("take snapshot", StringComparison.Ordinal)
            || normalized.Contains(" and ", StringComparison.Ordinal)
            || normalized.Contains(" then ", StringComparison.Ordinal);
    }

    private static bool IsSafeAutoExecuteAction(ActionType type)
    {
        return type is ActionType.Find
            or ActionType.ReadText
            or ActionType.WaitFor
            or ActionType.WaitForText
            or ActionType.CaptureScreen
            or ActionType.OpenApp
            or ActionType.GetClipboard
            or ActionType.FileRead
            or ActionType.FileList;
    }

    private static bool IsUnrecognizedPlan(ActionPlan plan)
    {
        if (plan.Steps.Count == 0)
        {
            return true;
        }

        if (plan.Steps.Count == 1 && plan.Steps[0].Type == ActionType.ReadText)
        {
            var note = plan.Steps[0].Note ?? string.Empty;
            if (string.IsNullOrWhiteSpace(note))
            {
                return true;
            }

            if (note.StartsWith("Default to read text", StringComparison.OrdinalIgnoreCase)
                || note.StartsWith("Unrecognized", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }

    private static string Compact(string value, int maxChars)
        => value.Length <= maxChars ? value : $"{value[..maxChars]}...";

    private static WebChatResponse ToChatResponse(WebIntentResponse response)
    {
        return new WebChatResponse(
            response.Reply,
            response.NeedsConfirmation,
            response.Token,
            response.ActionLabel,
            response.Steps,
            response.PlanJson,
            response.ModeLabel);
    }

    private string CreatePendingAction(PendingActionType type, string source, ActionPlan? plan, bool dryRun)
    {
        var token = Guid.NewGuid().ToString("N");
        lock (_sync)
        {
            _pendingActions[token] = new PendingAction(type, source, plan, dryRun);
        }

        return token;
    }

    private PendingAction? TakePendingAction(string token)
    {
        lock (_sync)
        {
            if (!_pendingActions.TryGetValue(token, out var pending))
            {
                return null;
            }

            _pendingActions.Remove(token);
            return pending;
        }
    }

    private WebTaskItem? FindTask(string name)
    {
        var key = name?.Trim() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(key))
        {
            return null;
        }

        lock (_sync)
        {
            return _tasks.FirstOrDefault(task => string.Equals(task.Name, key, StringComparison.OrdinalIgnoreCase));
        }
    }

    private async Task<WebChatResponse> LockCurrentWindowAsync(CancellationToken cancellationToken)
    {
        var active = await _desktopClient.GetActiveWindowAsync(cancellationToken);
        if (active == null)
        {
            return WebChatResponse.Simple("Cannot lock current window: no active window detected.");
        }

        lock (_sync)
        {
            _contextLock = new ContextLockState(true, active.Id, active.AppId, active.Title);
        }

        await WriteAuditAsync("context_lock", $"Context locked to window {active.Id}", cancellationToken, new { active.Id, active.AppId, active.Title });
        return WebChatResponse.Simple($"Context locked to window '{active.Title}' [{active.AppId}] ({active.Id}).");
    }

    private async Task<WebChatResponse> LockTargetAsync(string target, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(target))
        {
            return WebChatResponse.Simple("Specify target: lock on current window | lock on <app>.");
        }

        var lower = target.ToLowerInvariant();
        if (lower is "current" or "current window" or "this window")
        {
            return await LockCurrentWindowAsync(cancellationToken);
        }

        if (lower.StartsWith("window ", StringComparison.Ordinal))
        {
            var windowId = target["window ".Length..].Trim();
            if (string.IsNullOrWhiteSpace(windowId))
            {
                return WebChatResponse.Simple("Specify window id: lock on window <id>.");
            }

            lock (_sync)
            {
                _contextLock = new ContextLockState(true, windowId, null, null);
            }

            await WriteAuditAsync("context_lock", $"Context locked to window id {windowId}", cancellationToken);
            return WebChatResponse.Simple($"Context locked to window id '{windowId}'.");
        }

        var appTarget = target;
        if (_appResolver.TryResolveApp(target, out var resolved))
        {
            appTarget = resolved;
        }

        lock (_sync)
        {
            _contextLock = new ContextLockState(true, null, appTarget, null);
        }

        await WriteAuditAsync("context_lock", $"Context locked to app {appTarget}", cancellationToken);
        return WebChatResponse.Simple($"Context locked to app '{appTarget}'.");
    }
    private WebChatResponse ListApps(string originalText, string normalized)
    {
        var rawArgs = normalized.StartsWith("list apps", StringComparison.Ordinal)
            ? originalText["list apps".Length..]
            : originalText["apps".Length..];

        var tokens = rawArgs.Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).ToList();
        var flags = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "allowed", "allowlist", "--allowed", "--allowed-only", "--allowlist" };
        var allowedOnly = tokens.RemoveAll(token => flags.Contains(token)) > 0;
        var query = string.Join(" ", tokens);
        var matches = _appResolver.Suggest(query, 12);
        if (allowedOnly)
        {
            if (_config.AllowedApps.Count == 0)
            {
                return WebChatResponse.Simple("Allowlist is empty: all apps are considered allowed.");
            }

            matches = matches.Where(match => match.IsAllowed).ToList();
        }

        if (matches.Count == 0)
        {
            return WebChatResponse.Simple("No apps found.");
        }

        var lines = matches.Select(match =>
        {
            var tag = match.IsAllowed ? " [allowed]" : string.Empty;
            return $"{match.Entry.Name}{tag} score={match.Score:0.00} ({match.Entry.Path})";
        }).ToList();

        return WebChatResponse.WithSteps("Top apps:", lines, null, null);
    }

    private string FormatContextLock()
    {
        ContextLockState lockState;
        lock (_sync)
        {
            lockState = _contextLock;
        }

        return FormatContextLock(lockState);
    }

    private static string FormatContextLock(ContextLockState lockState)
    {
        if (!lockState.Enabled)
        {
            return "Context lock: OFF.";
        }

        if (!string.IsNullOrWhiteSpace(lockState.WindowId))
        {
            var titlePart = string.IsNullOrWhiteSpace(lockState.WindowTitle) ? string.Empty : $" '{lockState.WindowTitle}'";
            return $"Context lock: window{titlePart} ({lockState.WindowId}).";
        }

        if (!string.IsNullOrWhiteSpace(lockState.AppId))
        {
            return $"Context lock: app '{lockState.AppId}'.";
        }

        return "Context lock: ON.";
    }

    private async Task ScheduleLoopAsync(CancellationToken cancellationToken)
    {
        using var timer = new PeriodicTimer(TimeSpan.FromSeconds(1));
        try
        {
            while (await timer.WaitForNextTickAsync(cancellationToken))
            {
                List<ScheduleState> dueSchedules;
                var now = DateTimeOffset.UtcNow;
                lock (_sync)
                {
                    dueSchedules = _schedules
                        .Where(item => item.Enabled && IsScheduleDue(item, now))
                        .ToList();
                }

                var scheduleUpdated = false;
                if (dueSchedules.Count > 0)
                {
                    foreach (var schedule in dueSchedules)
                    {
                        try
                        {
                            await RunTaskAsync(schedule.TaskName, dryRun: false, cancellationToken);
                        }
                        catch
                        {
                            // Keep scheduler resilient.
                        }

                        lock (_sync)
                        {
                            var index = _schedules.FindIndex(item => string.Equals(item.Id, schedule.Id, StringComparison.OrdinalIgnoreCase));
                            if (index >= 0)
                            {
                                _schedules[index] = _schedules[index] with
                                {
                                    LastRunAtUtc = now,
                                    UpdatedAt = now
                                };
                                scheduleUpdated = true;
                            }
                        }
                    }
                }

                var goalsUpdated = await RunDueGoalsAsync(now, cancellationToken);
                if (scheduleUpdated)
                {
                    await SaveSchedulesToDiskAsync(cancellationToken);
                }

                if (goalsUpdated)
                {
                    await SaveGoalsToDiskAsync(cancellationToken);
                }
            }
        }
        catch (OperationCanceledException)
        {
            // Expected at shutdown.
        }
    }

    private async Task<bool> RunDueGoalsAsync(DateTimeOffset now, CancellationToken cancellationToken)
    {
        if (!_config.GoalSchedulerEnabled || _killSwitch.IsTripped)
        {
            return false;
        }

        var interval = TimeSpan.FromSeconds(Math.Clamp(_config.GoalSchedulerIntervalSeconds, 10, 3600));
        if (_lastGoalSchedulerSweepUtc != DateTimeOffset.MinValue && now - _lastGoalSchedulerSweepUtc < interval)
        {
            return false;
        }

        _lastGoalSchedulerSweepUtc = now;

        List<GoalState> dueGoals;
        lock (_sync)
        {
            dueGoals = _goals
                .Where(goal => !goal.Completed && goal.AutoRunEnabled && IsGoalDue(goal, now, interval))
                .OrderByDescending(goal => goal.Priority)
                .ThenByDescending(goal => goal.UpdatedAtUtc)
                .Take(Math.Clamp(_config.GoalSchedulerMaxPerTick, 1, 10))
                .Select(CloneGoalState)
                .ToList();
        }

        if (dueGoals.Count == 0)
        {
            return false;
        }

        foreach (var goal in dueGoals)
        {
            try
            {
                var response = await ExecuteGoalInternalAsync(goal.Id, approvedByUser: false, cancellationToken);
                await WriteAuditAsync(
                    "goal_scheduler_run",
                    $"Goal scheduler executed [{goal.Id}] => {Compact(response.Reply, 120)}",
                    cancellationToken,
                    new { goal.Id, goal.Text, goal.Priority });
            }
            catch (Exception ex)
            {
                lock (_sync)
                {
                    var index = _goals.FindIndex(item => string.Equals(item.Id, goal.Id, StringComparison.OrdinalIgnoreCase));
                    if (index >= 0)
                    {
                        var item = _goals[index];
                        item.LastRunAtUtc = now;
                        item.UpdatedAtUtc = now;
                        item.Attempts += 1;
                        item.LastResult = Compact($"Scheduler failed: {ex.Message}", 180);
                    }
                }
            }
        }

        return true;
    }

    private static bool IsScheduleDue(ScheduleState schedule, DateTimeOffset now)
    {
        if (!schedule.Enabled)
        {
            return false;
        }

        if (schedule.IntervalSeconds.HasValue)
        {
            if (schedule.LastRunAtUtc.HasValue)
            {
                return now - schedule.LastRunAtUtc.Value >= TimeSpan.FromSeconds(Math.Max(1, schedule.IntervalSeconds.Value));
            }

            if (!schedule.StartAtUtc.HasValue)
            {
                return true;
            }

            return schedule.StartAtUtc.Value <= now;
        }

        if (schedule.LastRunAtUtc.HasValue)
        {
            return false;
        }

        if (!schedule.StartAtUtc.HasValue)
        {
            return true;
        }

        return schedule.StartAtUtc.Value <= now;
    }

    private static bool IsGoalDue(GoalState goal, DateTimeOffset now, TimeSpan interval)
    {
        if (goal.Completed || !goal.AutoRunEnabled)
        {
            return false;
        }

        if (!goal.LastRunAtUtc.HasValue)
        {
            return true;
        }

        return now - goal.LastRunAtUtc.Value >= interval;
    }

    private async Task<WebIntentResponse> ExecuteGoalInternalAsync(string key, bool approvedByUser, CancellationToken cancellationToken)
    {
        GoalState? goal;
        lock (_sync)
        {
            var index = FindGoalIndexLocked(key);
            goal = index >= 0 ? _goals[index] : null;
        }

        if (goal == null)
        {
            return new WebIntentResponse("Goal not found.", false, null, null, null, null, null);
        }

        if (goal.Completed)
        {
            return new WebIntentResponse($"Goal [{goal.Id}] is already completed.", false, null, null, null, null, null);
        }

        var plan = _planner.PlanFromIntent(goal.Text);
        if (IsUnrecognizedPlan(plan))
        {
            lock (_sync)
            {
                goal.LastRunAtUtc = DateTimeOffset.UtcNow;
                goal.UpdatedAtUtc = DateTimeOffset.UtcNow;
                goal.Attempts += 1;
                goal.LastResult = "Goal plan could not be recognized.";
            }

            await SaveGoalsToDiskAsync(cancellationToken);
            return new WebIntentResponse($"Goal [{goal.Id}] could not be mapped to a safe plan.", false, null, null, null, null, null);
        }

        var source = $"goal:{goal.Id} {goal.Text}";
        return await ExecutePlanInternalAsync(source, plan, dryRun: false, approvedByUser, cancellationToken);
    }

    private static GoalState CloneGoalState(GoalState goal)
    {
        return new GoalState
        {
            Id = goal.Id,
            Text = goal.Text,
            Completed = goal.Completed,
            CreatedAtUtc = goal.CreatedAtUtc,
            UpdatedAtUtc = goal.UpdatedAtUtc,
            LastRunAtUtc = goal.LastRunAtUtc,
            Attempts = goal.Attempts,
            LastResult = goal.LastResult,
            Priority = goal.Priority,
            AutoRunEnabled = goal.AutoRunEnabled
        };
    }

    private async Task SaveTasksToDiskAsync(CancellationToken cancellationToken)
    {
        List<WebTaskItem> snapshot;
        lock (_sync)
        {
            snapshot = _tasks.OrderByDescending(task => task.UpdatedAt).ToList();
        }

        var path = Path.GetFullPath(_config.TaskLibraryPath);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        await File.WriteAllTextAsync(path, JsonSerializer.Serialize(snapshot, JsonOptions), cancellationToken);
    }

    private void LoadTasksFromDisk()
    {
        var path = Path.GetFullPath(_config.TaskLibraryPath);
        if (!File.Exists(path))
        {
            return;
        }

        try
        {
            var json = File.ReadAllText(path);
            var parsed = JsonSerializer.Deserialize<List<WebTaskItem>>(json, JsonOptions) ?? new List<WebTaskItem>();
            lock (_sync)
            {
                _tasks.Clear();
                _tasks.AddRange(parsed);
            }
        }
        catch
        {
            // Keep running with empty task list.
        }
    }

    private async Task SaveSchedulesToDiskAsync(CancellationToken cancellationToken)
    {
        List<ScheduleState> snapshot;
        lock (_sync)
        {
            snapshot = _schedules.ToList();
        }

        var path = Path.GetFullPath(_config.ScheduleLibraryPath);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        await File.WriteAllTextAsync(path, JsonSerializer.Serialize(snapshot, JsonOptions), cancellationToken);
    }

    private void LoadSchedulesFromDisk()
    {
        var path = Path.GetFullPath(_config.ScheduleLibraryPath);
        if (!File.Exists(path))
        {
            return;
        }

        try
        {
            var json = File.ReadAllText(path);
            var parsed = JsonSerializer.Deserialize<List<ScheduleState>>(json, JsonOptions) ?? new List<ScheduleState>();
            lock (_sync)
            {
                _schedules.Clear();
                _schedules.AddRange(parsed);
            }
        }
        catch
        {
            // Keep running with empty schedule list.
        }
    }

    private async Task SaveGoalsToDiskAsync(CancellationToken cancellationToken)
    {
        List<GoalState> snapshot;
        lock (_sync)
        {
            snapshot = _goals
                .OrderBy(goal => goal.Completed)
                .ThenByDescending(goal => goal.Priority)
                .ThenByDescending(goal => goal.UpdatedAtUtc)
                .ToList();
        }

        var path = Path.GetFullPath(_goalLibraryPath);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        await File.WriteAllTextAsync(path, JsonSerializer.Serialize(snapshot, JsonOptions), cancellationToken);
    }

    private void LoadGoalsFromDisk()
    {
        var path = Path.GetFullPath(_goalLibraryPath);
        if (!File.Exists(path))
        {
            return;
        }

        try
        {
            var json = File.ReadAllText(path);
            var parsed = JsonSerializer.Deserialize<List<GoalState>>(json, JsonOptions) ?? new List<GoalState>();
            foreach (var goal in parsed)
            {
                if (string.IsNullOrWhiteSpace(goal.Id))
                {
                    goal.Id = Guid.NewGuid().ToString("N")[..8];
                }

                goal.Priority = Math.Clamp(goal.Priority, GoalPriorityLow, GoalPriorityHigh);
                if (goal.CreatedAtUtc == default)
                {
                    goal.CreatedAtUtc = DateTimeOffset.UtcNow;
                }

                if (goal.UpdatedAtUtc == default)
                {
                    goal.UpdatedAtUtc = goal.CreatedAtUtc;
                }
            }

            lock (_sync)
            {
                _goals.Clear();
                _goals.AddRange(parsed);
            }
        }
        catch
        {
            // Keep running with empty goals list.
        }
    }

    private async Task SaveMemoryToDiskAsync(CancellationToken cancellationToken)
    {
        List<IntentMemoryEntry> snapshot;
        lock (_sync)
        {
            snapshot = _intentMemory
                .OrderByDescending(entry => entry.TimestampUtc)
                .Take(200)
                .ToList();
        }

        var path = Path.GetFullPath(_memoryLibraryPath);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        await File.WriteAllTextAsync(path, JsonSerializer.Serialize(snapshot, JsonOptions), cancellationToken);
    }

    private void LoadMemoryFromDisk()
    {
        var path = Path.GetFullPath(_memoryLibraryPath);
        if (!File.Exists(path))
        {
            return;
        }

        try
        {
            var json = File.ReadAllText(path);
            var parsed = JsonSerializer.Deserialize<List<IntentMemoryEntry>>(json, JsonOptions) ?? new List<IntentMemoryEntry>();
            lock (_sync)
            {
                _intentMemory.Clear();
                _intentMemory.AddRange(parsed.OrderByDescending(item => item.TimestampUtc).Take(200));
            }
        }
        catch
        {
            // Keep running with empty memory.
        }
    }

    private async Task SaveConfigToDiskAsync(CancellationToken cancellationToken)
    {
        var path = Path.GetFullPath(_configPath);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        var json = JsonSerializer.Serialize(_config, JsonOptions);
        await File.WriteAllTextAsync(path, json, cancellationToken);
    }

    private static string ResolveConfigPath(string? configPath, string storageRoot)
    {
        Directory.CreateDirectory(storageRoot);
        var fallbackPath = Path.Combine(storageRoot, "agentsettings.json");

        if (!string.IsNullOrWhiteSpace(configPath))
        {
            var candidate = configPath.Trim();
            if (!Path.IsPathRooted(candidate))
            {
                // Relative paths are anchored under LocalAppData so updates don't overwrite user settings.
                var persistentPath = Path.GetFullPath(Path.Combine(storageRoot, candidate));
                var legacyAppPath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, candidate));
                if (!File.Exists(persistentPath) && File.Exists(legacyAppPath))
                {
                    TryCopyConfigFile(legacyAppPath, persistentPath);
                }

                return persistentPath;
            }

            var fullPath = Path.GetFullPath(candidate);
            if (CanWriteConfigPath(fullPath))
            {
                return fullPath;
            }

            TryCopyConfigFile(fullPath, fallbackPath);
            return fallbackPath;
        }

        var legacyDefaultPath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "agentsettings.json"));
        if (!File.Exists(fallbackPath) && File.Exists(legacyDefaultPath))
        {
            TryCopyConfigFile(legacyDefaultPath, fallbackPath);
        }

        return fallbackPath;
    }

    private static bool CanWriteConfigPath(string path)
    {
        try
        {
            var directory = Path.GetDirectoryName(path);
            if (string.IsNullOrWhiteSpace(directory))
            {
                return false;
            }

            Directory.CreateDirectory(directory);

            if (File.Exists(path))
            {
                var attributes = File.GetAttributes(path);
                if ((attributes & FileAttributes.ReadOnly) == FileAttributes.ReadOnly)
                {
                    return false;
                }
            }

            var probePath = Path.Combine(directory, $".write-probe-{Guid.NewGuid():N}.tmp");
            using (new FileStream(probePath, FileMode.CreateNew, FileAccess.Write, FileShare.None, 1, FileOptions.DeleteOnClose))
            {
                // Probe successful.
            }

            return true;
        }
        catch
        {
            return false;
        }
    }

    private static void TryCopyConfigFile(string sourcePath, string destinationPath)
    {
        try
        {
            if (!File.Exists(sourcePath))
            {
                return;
            }

            var directory = Path.GetDirectoryName(destinationPath);
            if (!string.IsNullOrWhiteSpace(directory))
            {
                Directory.CreateDirectory(directory);
            }

            File.Copy(sourcePath, destinationPath, overwrite: false);
        }
        catch
        {
            // Fallback file copy is best effort.
        }
    }

    private static AgentConfig LoadConfig(string configPath, string adapterEndpoint, string storageRoot)
    {
        var config = new AgentConfig();

        if (File.Exists(configPath))
        {
            var configuration = new ConfigurationBuilder()
                .AddJsonFile(configPath, optional: false)
                .AddEnvironmentVariables(prefix: "DESKTOP_AGENT_")
                .Build();
            configuration.Bind(config);
        }
        else
        {
            var legacyWebPath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "web", "appsettings.json"));
            var legacyCliPath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "cli", "appsettings.json"));
            if (File.Exists(legacyWebPath))
            {
                var configuration = new ConfigurationBuilder()
                    .AddJsonFile(legacyWebPath, optional: false)
                    .AddEnvironmentVariables(prefix: "DESKTOP_AGENT_")
                    .Build();
                configuration.Bind(config);
            }
            else if (File.Exists(legacyCliPath))
            {
                var configuration = new ConfigurationBuilder()
                    .AddJsonFile(legacyCliPath, optional: false)
                    .AddEnvironmentVariables(prefix: "DESKTOP_AGENT_")
                    .Build();
                configuration.Bind(config);
            }
            else
            {
                var configuration = new ConfigurationBuilder()
                    .AddEnvironmentVariables(prefix: "DESKTOP_AGENT_")
                    .Build();
                configuration.Bind(config);
            }
        }

        if (!string.IsNullOrWhiteSpace(adapterEndpoint))
        {
            config.AdapterEndpoint = adapterEndpoint;
        }

        AgentConfigSanitizer.Normalize(config);

        config.AuditLogPath = ResolveWritablePath(config.AuditLogPath, storageRoot, "audit.log.jsonl");
        config.AppIndexCachePath = ResolveWritablePath(config.AppIndexCachePath, storageRoot, "app-index.json");
        config.TaskLibraryPath = ResolveWritablePath(config.TaskLibraryPath, storageRoot, "tasks.library.json");
        config.ScheduleLibraryPath = ResolveWritablePath(config.ScheduleLibraryPath, storageRoot, "schedules.library.json");
        return config;
    }

    private static string ResolveWritablePath(string path, string storageRoot, string fallbackName)
    {
        var candidate = string.IsNullOrWhiteSpace(path) ? fallbackName : path.Trim();
        if (!Path.IsPathRooted(candidate))
        {
            candidate = Path.Combine(storageRoot, candidate);
        }

        if (IsProgramFilesPath(candidate))
        {
            candidate = Path.Combine(storageRoot, Path.GetFileName(candidate));
        }

        var full = Path.GetFullPath(candidate);
        Directory.CreateDirectory(Path.GetDirectoryName(full)!);
        return full;
    }

    private static bool IsProgramFilesPath(string path)
    {
        var full = Path.GetFullPath(path);
        var programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
        var programFilesX86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);
        return full.StartsWith(programFiles, StringComparison.OrdinalIgnoreCase)
            || (!string.IsNullOrWhiteSpace(programFilesX86) && full.StartsWith(programFilesX86, StringComparison.OrdinalIgnoreCase));
    }

    private async Task WriteAuditAsync(string type, string message, CancellationToken cancellationToken, object? data = null)
    {
        await _auditLog.WriteAsync(new AuditEvent
        {
            Timestamp = DateTimeOffset.UtcNow,
            EventType = type,
            Message = message,
            Data = data
        }, cancellationToken);
    }

    private static WebScheduleItem ToScheduleItem(ScheduleState schedule)
    {
        return new WebScheduleItem(
            schedule.Id,
            schedule.TaskName,
            schedule.StartAtUtc,
            schedule.IntervalSeconds,
            schedule.Enabled,
            schedule.UpdatedAt);
    }

    private static bool TryParseCommand(string commandLine, out string fileName, out string args)
    {
        fileName = string.Empty;
        args = string.Empty;
        if (string.IsNullOrWhiteSpace(commandLine))
        {
            return false;
        }

        var tokens = SplitCommandLine(commandLine);
        if (tokens.Count == 0)
        {
            return false;
        }

        fileName = tokens[0];
        args = tokens.Count > 1 ? string.Join(' ', tokens.Skip(1)) : string.Empty;
        return true;
    }

    private static List<string> SplitCommandLine(string commandLine)
    {
        var results = new List<string>();
        var current = new StringBuilder();
        var inQuotes = false;
        char quoteChar = '\0';

        foreach (var ch in commandLine)
        {
            if (inQuotes)
            {
                if (ch == quoteChar)
                {
                    inQuotes = false;
                }
                else
                {
                    current.Append(ch);
                }
                continue;
            }

            if (ch is '"' or '\'')
            {
                inQuotes = true;
                quoteChar = ch;
                continue;
            }

            if (char.IsWhiteSpace(ch))
            {
                if (current.Length > 0)
                {
                    results.Add(current.ToString());
                    current.Clear();
                }
                continue;
            }

            current.Append(ch);
        }

        if (current.Length > 0)
        {
            results.Add(current.ToString());
        }

        return results;
    }

    private sealed class AutoApproveConfirmation : IUserConfirmation
    {
        public Task<bool> ConfirmAsync(string message, CancellationToken cancellationToken) => Task.FromResult(true);
    }

    private sealed class RejectConfirmation : IUserConfirmation
    {
        public Task<bool> ConfirmAsync(string message, CancellationToken cancellationToken) => Task.FromResult(false);
    }

    private enum PendingActionType
    {
        ExecutePlan,
        SimulatePresence
    }

    private sealed record PendingAction(PendingActionType Type, string Source, ActionPlan? Plan, bool DryRun);
    private sealed record ClarificationChoice(string Label, string Intent, string Description);
    private sealed record ClarificationRequest(string Kind, string Input, IReadOnlyList<ClarificationChoice> Options);

    private sealed record ContextLockState(bool Enabled, string? WindowId, string? AppId, string? WindowTitle)
    {
        public static ContextLockState None => new(false, null, null, null);
    }

    private sealed class GoalState
    {
        public string Id { get; set; } = string.Empty;
        public string Text { get; set; } = string.Empty;
        public bool Completed { get; set; }
        public DateTimeOffset CreatedAtUtc { get; set; }
        public DateTimeOffset UpdatedAtUtc { get; set; }
        public DateTimeOffset? LastRunAtUtc { get; set; }
        public int Attempts { get; set; }
        public string? LastResult { get; set; }
        public int Priority { get; set; } = GoalPriorityNormal;
        public bool AutoRunEnabled { get; set; } = true;
    }

    private sealed record IntentMemoryEntry(
        DateTimeOffset TimestampUtc,
        string Source,
        string IntentSummary,
        bool Success,
        int PlanSteps,
        string ResultMessage,
        string? BeforeWindow,
        string? AfterWindow);

    private sealed record ObservationSnapshot(string WindowId, string WindowDisplay)
    {
        public static ObservationSnapshot Empty { get; } = new(string.Empty, string.Empty);
    }

    private sealed record RecoveryAttemptOutcome(ExecutionResult? RecoveredResult, string Note)
    {
        public static RecoveryAttemptOutcome None { get; } = new(null, string.Empty);
    }

    private sealed record OrderFillExecutionReport(string Summary, IReadOnlyList<string> Lines);

    private sealed record OrderDraft(
        string Id,
        DateTimeOffset CreatedAtUtc,
        string RawText,
        string PayloadJson,
        IReadOnlyList<string> SummaryLines);

    private sealed class OrderDraftPayload
    {
        public string? OrderNumber { get; set; }
        public string? OrderDate { get; set; }
        public OrderCustomerPayload? Customer { get; set; }
        public string? ShippingAddress { get; set; }
        public string? BillingAddress { get; set; }
        public List<OrderItemPayload>? Items { get; set; }
        public string? Notes { get; set; }
    }

    private sealed class OrderCustomerPayload
    {
        public string? Name { get; set; }
        public string? Email { get; set; }
        public string? Phone { get; set; }
    }

    private sealed class OrderItemPayload
    {
        public string? Sku { get; set; }
        public string? Name { get; set; }
        public decimal? Qty { get; set; }
        public decimal? UnitPrice { get; set; }
        public string? Currency { get; set; }
    }

    private sealed record DiscoveredFormField(
        string Key,
        string BestTextHint,
        string AutomationId,
        string Name,
        string Placeholder,
        string Type,
        string AutoComplete);

    private sealed record SmartOrderFillPlan(
        ActionPlan Plan,
        int MappedFields,
        int DiscoveredFields);

    private readonly record struct PlanSafetyResult(bool Allowed, bool AutoExecute, string Reason);

    private sealed record ScheduleState(
        string Id,
        string TaskName,
        DateTimeOffset? StartAtUtc,
        int? IntervalSeconds,
        bool Enabled,
        DateTimeOffset UpdatedAt,
        DateTimeOffset? LastRunAtUtc);
}
