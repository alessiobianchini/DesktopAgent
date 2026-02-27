using System.Net.Http;
using System.Text;
using System.Text.Json;

namespace DesktopAgent.Tray;

internal sealed class WebApiClient : IDisposable
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true
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

    public async Task<string> GetStatusLineAsync(CancellationToken cancellationToken)
    {
        using var response = await _httpClient.GetAsync("api/status", cancellationToken);
        var body = await response.Content.ReadAsStringAsync(cancellationToken);
        if (!response.IsSuccessStatusCode)
        {
            return $"HTTP {(int)response.StatusCode}";
        }

        try
        {
            using var json = JsonDocument.Parse(body);
            if (!json.RootElement.TryGetProperty("adapter", out var adapter))
            {
                return "status unavailable";
            }

            var armed = adapter.TryGetProperty("armed", out var armedProp) && armedProp.GetBoolean();
            var requirePresence = adapter.TryGetProperty("requireUserPresence", out var presenceProp) && presenceProp.GetBoolean();
            var armedLabel = armed ? "ARMED:ON" : "ARMED:OFF";
            var presenceLabel = requirePresence ? "PRESENCE:REQ" : "PRESENCE:OFF";
            return $"{armedLabel} | {presenceLabel}";
        }
        catch
        {
            return "status parse error";
        }
    }

    public void Dispose()
    {
        _httpClient.Dispose();
    }

    private async Task<WebChatResponse> PostChatLikeAsync(string path, object payload, CancellationToken cancellationToken)
    {
        var content = new StringContent(
            JsonSerializer.Serialize(payload),
            Encoding.UTF8,
            "application/json");

        using var response = await _httpClient.PostAsync(path, content, cancellationToken);
        var body = await response.Content.ReadAsStringAsync(cancellationToken);
        if (!response.IsSuccessStatusCode)
        {
            return WebChatResponse.Error($"HTTP {(int)response.StatusCode}: {Compact(body)}");
        }

        try
        {
            var parsed = JsonSerializer.Deserialize<WebChatResponse>(body, JsonOptions);
            return parsed ?? WebChatResponse.Error("Empty response body.");
        }
        catch (Exception ex)
        {
            return WebChatResponse.Error($"Response parse error: {ex.Message}");
        }
    }

    private static string Compact(string text)
    {
        const int maxChars = 200;
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
