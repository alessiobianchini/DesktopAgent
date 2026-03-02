using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace DesktopAgent.Tray;

internal sealed class WebApiClient : IDisposable
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    private readonly HttpClient _httpClient;

    public WebApiClient(string webUiUrl, TimeSpan timeout)
    {
        var baseUri = webUiUrl.Trim().TrimEnd('/');
        _httpClient = new HttpClient
        {
            BaseAddress = new Uri($"{baseUri}/"),
            Timeout = timeout
        };
    }

    public async Task<WebChatResponse> SendChatAsync(string message, CancellationToken cancellationToken)
    {
        var payload = new { message };
        return await PostChatLikeAsync("api/chat", payload, cancellationToken);
    }

    public async Task<WebChatResponse> ConfirmAsync(string token, bool approve, CancellationToken cancellationToken)
    {
        var payload = new { token, approve };
        return await PostChatLikeAsync("api/confirm", payload, cancellationToken);
    }

    public async Task<WebStatusResponse?> GetStatusAsync(CancellationToken cancellationToken)
    {
        return await GetJsonAsync<WebStatusResponse>("api/status", cancellationToken);
    }

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

    public async Task<WebConfigResponse?> GetConfigAsync(CancellationToken cancellationToken)
    {
        return await GetJsonAsync<WebConfigResponse>("api/config", cancellationToken);
    }

    public async Task<WebConfigResponse?> SaveConfigAsync(WebConfigUpdate payload, CancellationToken cancellationToken)
    {
        return await PostJsonAsync<WebConfigUpdate, WebConfigResponse>("api/config", payload, cancellationToken);
    }

    public async Task<WebTasksResponse?> GetTasksAsync(CancellationToken cancellationToken)
    {
        return await GetJsonAsync<WebTasksResponse>("api/tasks", cancellationToken);
    }

    public async Task<WebApiSimpleResponse?> SaveTaskAsync(WebTaskUpsertRequest payload, CancellationToken cancellationToken)
    {
        return await PostJsonAsync<WebTaskUpsertRequest, WebApiSimpleResponse>("api/tasks", payload, cancellationToken);
    }

    public async Task<WebApiSimpleResponse?> DeleteTaskAsync(string name, CancellationToken cancellationToken)
    {
        using var response = await _httpClient.DeleteAsync($"api/tasks/{Uri.EscapeDataString(name)}", cancellationToken);
        return await ParseApiResponseAsync<WebApiSimpleResponse>(response, cancellationToken);
    }

    public async Task<WebIntentResponse?> RunTaskAsync(string name, bool dryRun, CancellationToken cancellationToken)
    {
        return await PostJsonAsync<object, WebIntentResponse>(
            $"api/tasks/{Uri.EscapeDataString(name)}/run",
            new { dryRun },
            cancellationToken);
    }

    public async Task<WebSchedulesResponse?> GetSchedulesAsync(CancellationToken cancellationToken)
    {
        return await GetJsonAsync<WebSchedulesResponse>("api/schedules", cancellationToken);
    }

    public async Task<WebScheduleSaveResponse?> SaveScheduleAsync(WebScheduleUpsertRequest payload, CancellationToken cancellationToken)
    {
        return await PostJsonAsync<WebScheduleUpsertRequest, WebScheduleSaveResponse>("api/schedules", payload, cancellationToken);
    }

    public async Task<WebApiSimpleResponse?> DeleteScheduleAsync(string id, CancellationToken cancellationToken)
    {
        using var response = await _httpClient.DeleteAsync($"api/schedules/{Uri.EscapeDataString(id)}", cancellationToken);
        return await ParseApiResponseAsync<WebApiSimpleResponse>(response, cancellationToken);
    }

    public async Task<WebApiSimpleResponse?> RunScheduleNowAsync(string id, CancellationToken cancellationToken)
    {
        return await PostJsonAsync<object, WebApiSimpleResponse>(
            $"api/schedules/{Uri.EscapeDataString(id)}/run",
            new { },
            cancellationToken);
    }

    public async Task<WebAuditResponse?> GetAuditAsync(int take, CancellationToken cancellationToken)
    {
        return await GetJsonAsync<WebAuditResponse>($"api/audit?take={take}", cancellationToken);
    }

    public async Task<WebApiSimpleResponse?> RestartServerAsync(CancellationToken cancellationToken)
    {
        return await PostJsonAsync<object, WebApiSimpleResponse>("api/restart", new { }, cancellationToken);
    }

    public async Task<WebApiSimpleResponse?> RestartAdapterAsync(CancellationToken cancellationToken)
    {
        return await PostJsonAsync<object, WebApiSimpleResponse>("api/adapter/restart", new { }, cancellationToken);
    }

    public async Task<WebIntentResponse?> ExecuteIntentAsync(string intent, bool dryRun, CancellationToken cancellationToken)
    {
        return await PostJsonAsync<object, WebIntentResponse>("api/intent", new { intent, dryRun }, cancellationToken);
    }

    public void Dispose()
    {
        _httpClient.Dispose();
    }

    private async Task<WebChatResponse> PostChatLikeAsync(string path, object payload, CancellationToken cancellationToken)
    {
        using var response = await PostRawJsonAsync(path, payload, cancellationToken);
        if (!response.IsSuccessStatusCode)
        {
            var error = await BuildErrorAsync(response, cancellationToken);
            return WebChatResponse.Error(error);
        }

        var parsed = await DeserializeResponseAsync<WebChatResponse>(response, cancellationToken);
        return parsed ?? WebChatResponse.Error("Empty response body.");
    }

    private async Task<T?> GetJsonAsync<T>(string path, CancellationToken cancellationToken)
    {
        using var response = await _httpClient.GetAsync(path, cancellationToken);
        if (!response.IsSuccessStatusCode)
        {
            return default;
        }

        return await DeserializeResponseAsync<T>(response, cancellationToken);
    }

    private async Task<TResponse?> PostJsonAsync<TRequest, TResponse>(string path, TRequest payload, CancellationToken cancellationToken)
    {
        using var response = await PostRawJsonAsync(path, payload, cancellationToken);
        return await ParseApiResponseAsync<TResponse>(response, cancellationToken);
    }

    private async Task<TResponse?> ParseApiResponseAsync<TResponse>(HttpResponseMessage response, CancellationToken cancellationToken)
    {
        if (!response.IsSuccessStatusCode)
        {
            var message = await BuildErrorAsync(response, cancellationToken);
            if (typeof(TResponse) == typeof(WebApiSimpleResponse))
            {
                return (TResponse?)(object)new WebApiSimpleResponse(message, false);
            }

            if (typeof(TResponse) == typeof(WebChatResponse))
            {
                return (TResponse?)(object)WebChatResponse.Error(message);
            }

            return default;
        }

        return await DeserializeResponseAsync<TResponse>(response, cancellationToken);
    }

    private async Task<HttpResponseMessage> PostRawJsonAsync<TRequest>(string path, TRequest payload, CancellationToken cancellationToken)
    {
        var content = new StringContent(
            JsonSerializer.Serialize(payload, JsonOptions),
            Encoding.UTF8,
            "application/json");

        return await _httpClient.PostAsync(path, content, cancellationToken);
    }

    private static async Task<T?> DeserializeResponseAsync<T>(HttpResponseMessage response, CancellationToken cancellationToken)
    {
        var body = await response.Content.ReadAsStringAsync(cancellationToken);
        if (string.IsNullOrWhiteSpace(body))
        {
            return default;
        }

        try
        {
            return JsonSerializer.Deserialize<T>(body, JsonOptions);
        }
        catch
        {
            return default;
        }
    }

    private static async Task<string> BuildErrorAsync(HttpResponseMessage response, CancellationToken cancellationToken)
    {
        var body = await response.Content.ReadAsStringAsync(cancellationToken);
        if (string.IsNullOrWhiteSpace(body))
        {
            return $"HTTP {(int)response.StatusCode}";
        }

        try
        {
            using var json = JsonDocument.Parse(body);
            if (json.RootElement.TryGetProperty("message", out var messageElement))
            {
                var message = messageElement.GetString();
                if (!string.IsNullOrWhiteSpace(message))
                {
                    return message!;
                }
            }
        }
        catch
        {
            // ignored
        }

        return $"HTTP {(int)response.StatusCode}: {Compact(body)}";
    }

    private static string Compact(string text)
    {
        const int maxChars = 250;
        if (text.Length <= maxChars)
        {
            return text;
        }

        return text[..maxChars] + "...";
    }
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
    {
        return new WebChatResponse(message, false, null, null, null, null, null);
    }
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

internal sealed record WebAuditResponse(IReadOnlyList<string>? Lines);
