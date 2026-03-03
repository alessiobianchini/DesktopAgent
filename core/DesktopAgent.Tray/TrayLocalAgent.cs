using System.Diagnostics;
using System.Net.Http;
using System.Reflection;
using System.Text;
using System.Text.Json;
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
    private readonly List<WebTaskItem> _tasks = new();
    private readonly List<ScheduleState> _schedules = new();
    private readonly List<GoalState> _goals = new();
    private readonly List<IntentMemoryEntry> _intentMemory = new();
    private readonly CancellationTokenSource _schedulerCts = new();
    private readonly Task _schedulerTask;

    private ContextLockState _contextLock = ContextLockState.None;
    private DateTimeOffset _lastGoalSchedulerSweepUtc = DateTimeOffset.MinValue;

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
        _ocrEngine = _config.OcrEnabled
            ? new TesseractOcrEngine(_config.Ocr.TesseractPath, _loggerFactory.CreateLogger<TesseractOcrEngine>())
            : new OcrEngineStub();
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
            return informational;
        }

        var fileVersion = assembly.GetCustomAttribute<AssemblyFileVersionAttribute>()?.Version;
        if (!string.IsNullOrWhiteSpace(fileVersion))
        {
            return fileVersion;
        }

        return assembly.GetName().Version?.ToString() ?? "unknown";
    }

    public async Task<WebChatResponse> SendChatAsync(string message, CancellationToken cancellationToken)
    {
        var text = message?.Trim() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(text))
        {
            return WebChatResponse.Simple("Enter a valid command.");
        }

        var normalized = text.ToLowerInvariant();

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
            return await BuildPlanConfirmationAsync(intent, dryRun: false);
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
            return await BuildPlanConfirmationAsync(intent, dryRun: false);
        }

        if (IsDirectIntent(normalized))
        {
            return await BuildPlanConfirmationAsync(text, dryRun: false);
        }

        var plan = _planner.PlanFromIntent(text);
        if (IsUnrecognizedPlan(plan))
        {
            return WebChatResponse.Simple("Available commands: status, kill, reset kill, lock status, lock on <current window|app>, unlock, profile <safe|balanced|power>, arm, disarm, simulate presence, require presence, list apps [query] [allowed], goals, goal add <text>, goal run <id>, goal done <id>, goal remove <id>, goal priority <id> <low|normal|high>, goal auto <id> <on|off>, goal scheduler on|off|every <sec>, continue goal, memory, run <intent>, dry-run <intent>. Plugin intents: take screenshot, record screen [and audio] for <duration>, start recording [screen] [with/without audio], stop recording, jiggle mouse for <duration>.");
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
            RequireConfirmation: _config.RequireConfirmation,
            MaxActionsPerSecond: _config.MaxActionsPerSecond,
            QuizSafeModeEnabled: _config.QuizSafeModeEnabled,
            OcrEnabled: _config.OcrEnabled,
            OcrRestartRequired: _ocrRestartRequired,
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

    private async Task<WebChatResponse> BuildPlanConfirmationAsync(string intent, bool dryRun)
    {
        var plan = _planner.PlanFromIntent(intent ?? string.Empty);
        if (IsUnrecognizedPlan(plan))
        {
            return WebChatResponse.Simple("I couldn't map that to a safe plan. Try a clearer command.");
        }

        if (dryRun)
        {
            var result = await ExecutePlanInternalAsync(intent ?? string.Empty, plan, dryRun: true, approvedByUser: false, CancellationToken.None);
            return ToChatResponse(result);
        }

        var token = CreatePendingAction(PendingActionType.ExecutePlan, intent ?? string.Empty, plan, dryRun: false);
        var notice = GetRewriteNotice(plan);
        var prompt = string.IsNullOrWhiteSpace(notice)
            ? "I interpreted your request. Confirm execution?"
            : $"I interpreted your request. {notice}. Confirm execution?";
        return WebChatResponse.ConfirmWithSteps(prompt, token, PlanToLines(plan), PlanToJson(plan), GetModeLabel(plan));
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
        return new WebIntentResponse(reply, false, null, null, ExecutionToLines(result), PlanToJson(plan), GetModeLabel(plan));
    }

    private async Task<WebLlmStatus> GetLlmStatusAsync(CancellationToken cancellationToken)
    {
        var provider = _config.LlmFallback.Provider;
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
        => plan.Steps.Select(step => step.Note)
            .FirstOrDefault(note => !string.IsNullOrWhiteSpace(note) && note.StartsWith("Rewritten intent:", StringComparison.OrdinalIgnoreCase));

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
                candidate = Path.Combine(AppContext.BaseDirectory, candidate);
            }

            var fullPath = Path.GetFullPath(candidate);
            if (CanWriteConfigPath(fullPath))
            {
                return fullPath;
            }

            TryCopyConfigFile(fullPath, fallbackPath);
            return fallbackPath;
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

    private sealed record ScheduleState(
        string Id,
        string TaskName,
        DateTimeOffset? StartAtUtc,
        int? IntervalSeconds,
        bool Enabled,
        DateTimeOffset UpdatedAt,
        DateTimeOffset? LastRunAtUtc);
}
