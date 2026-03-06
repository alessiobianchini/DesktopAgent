namespace DesktopAgent.Tray;

internal sealed class WebApiClient : IDisposable
{
    private readonly TrayLocalAgent _agent;

    public WebApiClient(string adapterEndpoint, TimeSpan timeout, string? configPath = null)
    {
        _agent = new TrayLocalAgent(adapterEndpoint, timeout, configPath);
    }

    public Task<WebChatResponse> SendChatAsync(string message, CancellationToken cancellationToken) => _agent.SendChatAsync(message, cancellationToken);
    public Task<WebChatResponse> ConfirmAsync(string token, bool approve, CancellationToken cancellationToken) => _agent.ConfirmAsync(token, approve, cancellationToken);
    public Task<WebStatusResponse?> GetStatusAsync(CancellationToken cancellationToken) => _agent.GetStatusAsync(cancellationToken);
    public Task<WebConfigResponse?> GetConfigAsync(CancellationToken cancellationToken) => _agent.GetConfigAsync(cancellationToken);
    public Task<WebConfigResponse?> SaveConfigAsync(WebConfigUpdate payload, CancellationToken cancellationToken) => _agent.SaveConfigAsync(payload, cancellationToken);
    public Task<WebTasksResponse?> GetTasksAsync(CancellationToken cancellationToken) => _agent.GetTasksAsync(cancellationToken);
    public Task<WebApiSimpleResponse?> SaveTaskAsync(WebTaskUpsertRequest payload, CancellationToken cancellationToken) => _agent.SaveTaskAsync(payload, cancellationToken);
    public Task<WebApiSimpleResponse?> DeleteTaskAsync(string name, CancellationToken cancellationToken) => _agent.DeleteTaskAsync(name, cancellationToken);
    public Task<WebIntentResponse?> RunTaskAsync(string name, bool dryRun, CancellationToken cancellationToken) => _agent.RunTaskAsync(name, dryRun, cancellationToken);
    public Task<WebSchedulesResponse?> GetSchedulesAsync(CancellationToken cancellationToken) => _agent.GetSchedulesAsync(cancellationToken);
    public Task<WebScheduleSaveResponse?> SaveScheduleAsync(WebScheduleUpsertRequest payload, CancellationToken cancellationToken) => _agent.SaveScheduleAsync(payload, cancellationToken);
    public Task<WebApiSimpleResponse?> DeleteScheduleAsync(string id, CancellationToken cancellationToken) => _agent.DeleteScheduleAsync(id, cancellationToken);
    public Task<WebApiSimpleResponse?> RunScheduleNowAsync(string id, CancellationToken cancellationToken) => _agent.RunScheduleNowAsync(id, cancellationToken);
    public Task<WebGoalsResponse?> GetGoalsAsync(CancellationToken cancellationToken) => _agent.GetGoalsAsync(cancellationToken);
    public Task<WebApiSimpleResponse?> AddGoalAsync(string text, CancellationToken cancellationToken) => _agent.AddGoalFromUiAsync(text, cancellationToken);
    public Task<WebApiSimpleResponse?> SetGoalAutoAsync(string id, bool enabled, CancellationToken cancellationToken) => _agent.SetGoalAutoFromUiAsync(id, enabled, cancellationToken);
    public Task<WebApiSimpleResponse?> MarkGoalDoneAsync(string id, CancellationToken cancellationToken) => _agent.MarkGoalDoneFromUiAsync(id, cancellationToken);
    public Task<WebApiSimpleResponse?> RemoveGoalAsync(string id, CancellationToken cancellationToken) => _agent.RemoveGoalFromUiAsync(id, cancellationToken);
    public Task<WebAuditResponse?> GetAuditAsync(int take, CancellationToken cancellationToken) => _agent.GetAuditAsync(take, cancellationToken);
    public Task<WebApiSimpleResponse?> RestartServerAsync(CancellationToken cancellationToken) => _agent.RestartServerAsync(cancellationToken);
    public Task<WebApiSimpleResponse?> RestartAdapterAsync(CancellationToken cancellationToken) => _agent.RestartAdapterAsync(cancellationToken);
    public Task<WebIntentResponse?> ExecuteIntentAsync(string intent, bool dryRun, CancellationToken cancellationToken) => _agent.ExecuteIntentAsync(intent, dryRun, cancellationToken);
    public Task WriteSystemAuditAsync(string eventType, string message, object? data, CancellationToken cancellationToken) => _agent.WriteSystemAuditAsync(eventType, message, data, cancellationToken);

    public async Task<string> GetStatusLineAsync(CancellationToken cancellationToken)
    {
        var status = await GetStatusAsync(cancellationToken);
        if (status?.Adapter == null)
        {
            return "status unavailable";
        }

        var armedLabel = status.Adapter.Armed ? "ARMED:ON" : "ARMED:OFF";
        var presenceLabel = status.Adapter.RequireUserPresence ? "PRESENCE:REQ" : "PRESENCE:OFF";
        var llmLabel = status.Llm == null
            ? "LLM:UNKNOWN"
            : (status.Llm.Enabled && status.Llm.Available ? "LLM:ON" : "LLM:OFF");
        var versionLabel = string.IsNullOrWhiteSpace(status.Version) ? "v:unknown" : $"v:{status.Version}";
        return $"{armedLabel} | {presenceLabel} | {llmLabel} | {versionLabel}";
    }

    public void Dispose() => _agent.Dispose();
}

internal sealed record WebChatResponse(
    string Reply,
    bool NeedsConfirmation,
    string? Token,
    string? ActionLabel,
    IReadOnlyList<string>? Steps,
    string? PlanJson,
    string? ModeLabel)
{
    public static WebChatResponse Error(string message)
        => new(message, false, null, null, null, null, null);

    public static WebChatResponse Simple(string message)
        => new(message, false, null, null, null, null, null);

    public static WebChatResponse WithSteps(string message, IReadOnlyList<string>? steps, string? planJson, string? modeLabel)
        => new(message, false, null, null, steps, planJson, modeLabel);

    public static WebChatResponse Confirm(string message, string token)
        => new(message, true, token, "Confirm", null, null, null);

    public static WebChatResponse ConfirmWithSteps(string message, string token, IReadOnlyList<string>? steps, string? planJson, string? modeLabel)
        => new(message, true, token, "Confirm", steps, planJson, modeLabel);
}

internal sealed record WebApiSimpleResponse(string? Message = null, bool? Success = null);

internal sealed record WebStatusResponse(
    string? Version,
    WebAdapterStatus? Adapter,
    WebLlmStatus? Llm,
    WebKillSwitchStatus? KillSwitch);

internal sealed record WebAdapterStatus(bool Armed, bool RequireUserPresence, string? Message);
internal sealed record WebLlmStatus(bool Enabled, bool Available, string? Provider, string? Message, string? Endpoint);
internal sealed record WebKillSwitchStatus(bool Tripped, string? Reason);

internal sealed record WebConfigResponse(
    bool ProfileModeEnabled,
    string? ActiveProfile,
    bool RequireConfirmation,
    int MaxActionsPerSecond,
    bool QuizSafeModeEnabled,
    bool OcrEnabled,
    bool OcrRestartRequired,
    string? MediaOutputDirectory,
    string? ScreenRecordingAudioBackendPreference,
    string? ScreenRecordingAudioDevice,
    bool ScreenRecordingPrimaryDisplayOnly,
    string? AdapterRestartCommand,
    string? AdapterRestartWorkingDir,
    int FindRetryCount,
    int FindRetryDelayMs,
    int PostCheckTimeoutMs,
    int PostCheckPollMs,
    int ClipboardHistoryMaxItems,
    IReadOnlyList<string>? FilesystemAllowedRoots,
    IReadOnlyList<string>? AllowedApps,
    Dictionary<string, string>? AppAliases,
    WebConfigLlm? Llm,
    bool AuditLlmInteractions,
    bool AuditLlmIncludeRawText);

internal sealed record WebConfigLlm(
    bool Enabled,
    bool AllowNonLoopbackEndpoint,
    string? Provider,
    string? Endpoint,
    string? Model,
    int TimeoutSeconds,
    int MaxTokens);

internal sealed record WebConfigUpdate(
    WebConfigUpdateLlm? Llm,
    bool? ProfileModeEnabled,
    string? ActiveProfile,
    bool? RequireConfirmation,
    int? MaxActionsPerSecond,
    bool? QuizSafeModeEnabled,
    bool? OcrEnabled,
    string? MediaOutputDirectory,
    string? ScreenRecordingAudioBackendPreference,
    string? ScreenRecordingAudioDevice,
    bool? ScreenRecordingPrimaryDisplayOnly,
    string? AdapterRestartCommand,
    string? AdapterRestartWorkingDir,
    int? FindRetryCount,
    int? FindRetryDelayMs,
    int? PostCheckTimeoutMs,
    int? PostCheckPollMs,
    int? ClipboardHistoryMaxItems,
    IReadOnlyList<string>? FilesystemAllowedRoots,
    IReadOnlyList<string>? AllowedApps,
    Dictionary<string, string>? AppAliases,
    bool? AuditLlmInteractions,
    bool? AuditLlmIncludeRawText);

internal sealed record WebConfigUpdateLlm(
    bool? Enabled,
    bool? AllowNonLoopbackEndpoint,
    string? Provider,
    string? Endpoint,
    string? Model,
    int? TimeoutSeconds,
    int? MaxTokens);

internal sealed record WebTasksResponse(IReadOnlyList<WebTaskItem>? Tasks);
internal sealed record WebTaskItem(string Name, string Intent, string? Description, string? PlanJson, DateTimeOffset UpdatedAt);
internal sealed record WebTaskUpsertRequest(string Name, string Intent, string? Description, string? PlanJson);

internal sealed record WebIntentResponse(
    string Reply,
    bool NeedsConfirmation,
    string? Token,
    string? ActionLabel,
    IReadOnlyList<string>? Steps,
    string? PlanJson,
    string? ModeLabel);

internal sealed record WebSchedulesResponse(IReadOnlyList<WebScheduleItem>? Schedules);
internal sealed record WebScheduleItem(
    string Id,
    string TaskName,
    DateTimeOffset? StartAtUtc,
    int? IntervalSeconds,
    bool Enabled,
    DateTimeOffset UpdatedAt);

internal sealed record WebScheduleUpsertRequest(
    string? Id,
    string TaskName,
    DateTimeOffset? StartAtUtc,
    int? IntervalSeconds,
    bool? Enabled);

internal sealed record WebScheduleSaveResponse(string? Message, WebScheduleItem? Schedule);

internal sealed record WebGoalsResponse(
    bool SchedulerEnabled,
    int SchedulerIntervalSeconds,
    IReadOnlyList<WebGoalItem>? Goals);

internal sealed record WebGoalItem(
    string Id,
    string Text,
    bool Completed,
    int Priority,
    bool AutoRunEnabled,
    int Attempts,
    DateTimeOffset UpdatedAtUtc,
    DateTimeOffset? LastRunAtUtc,
    string? LastResult);

internal sealed record WebAuditResponse(IReadOnlyList<string>? Lines);
