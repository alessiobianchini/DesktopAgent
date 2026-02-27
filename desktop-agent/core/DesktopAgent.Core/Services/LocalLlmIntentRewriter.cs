using System.Diagnostics;
using System.Net.Http.Json;
using System.Text.Json;
using System.Text.RegularExpressions;
using DesktopAgent.Core.Abstractions;
using DesktopAgent.Core.Config;
using DesktopAgent.Core.Models;

namespace DesktopAgent.Core.Services;

public sealed class LocalLlmIntentRewriter : ILlmIntentRewriter
{
    private static readonly JsonSerializerOptions JsonOptions = new(JsonSerializerDefaults.Web);
    private readonly AgentConfig _config;
    private readonly IAuditLog _auditLog;
    private readonly HttpClient _client;

    public LocalLlmIntentRewriter(AgentConfig config, IAuditLog auditLog)
    {
        _config = config;
        _auditLog = auditLog;
        _client = new HttpClient
        {
            Timeout = TimeSpan.FromSeconds(Math.Max(1, _config.LlmFallback.TimeoutSeconds))
        };
    }

    public string? Rewrite(string input)
    {
        if (string.IsNullOrWhiteSpace(input))
        {
            return null;
        }

        var endpoint = _config.LlmFallback.Endpoint?.Trim();
        if (string.IsNullOrWhiteSpace(endpoint))
        {
            WriteAudit("llm_error", "LLM endpoint missing", new { provider = _config.LlmFallback.Provider, model = _config.LlmFallback.Model });
            return null;
        }

        if (!Uri.TryCreate(endpoint, UriKind.Absolute, out var uri) || !uri.IsLoopback)
        {
            WriteAudit("llm_error", "LLM endpoint invalid", new { provider = _config.LlmFallback.Provider, model = _config.LlmFallback.Model, endpoint });
            return null;
        }

        var prompt = BuildPrompt(input);
        var provider = _config.LlmFallback.Provider?.Trim().ToLowerInvariant();
        var stopwatch = Stopwatch.StartNew();
        WriteAudit("llm_request", "LLM rewrite requested", new
        {
            provider,
            model = _config.LlmFallback.Model,
            endpoint,
            input = ToAuditText(input),
            inputLength = input.Length
        });

        try
        {
            string? rewritten;
            if (provider == "ollama")
            {
                rewritten = CallOllama(uri, prompt);
            }
            else if (provider == "openai")
            {
                rewritten = CallOpenAiCompatible(uri, prompt);
            }
            else if (provider == "llama.cpp")
            {
                rewritten = CallLlamaCpp(uri, prompt);
            }
            else
            {
                rewritten = CallOllama(uri, prompt);
            }

            stopwatch.Stop();
            WriteAudit("llm_response", "LLM rewrite completed", new
            {
                provider,
                model = _config.LlmFallback.Model,
                latencyMs = stopwatch.ElapsedMilliseconds,
                success = !string.IsNullOrWhiteSpace(rewritten),
                output = ToAuditText(rewritten),
                outputLength = rewritten?.Length ?? 0
            });
            return rewritten;
        }
        catch (Exception ex)
        {
            stopwatch.Stop();
            WriteAudit("llm_error", "LLM rewrite failed", new
            {
                provider,
                model = _config.LlmFallback.Model,
                latencyMs = stopwatch.ElapsedMilliseconds,
                error = ex.Message
            });
            return null;
        }
    }

    private string? CallOllama(Uri uri, string prompt)
    {
        var payload = new
        {
            model = _config.LlmFallback.Model,
            prompt,
            stream = false,
            options = new { temperature = 0.2, num_predict = _config.LlmFallback.MaxTokens }
        };

        var response = _client.PostAsJsonAsync(uri, payload, JsonOptions).GetAwaiter().GetResult();
        if (!response.IsSuccessStatusCode)
        {
            return null;
        }

        using var stream = response.Content.ReadAsStream();
        using var doc = JsonDocument.Parse(stream);
        if (doc.RootElement.TryGetProperty("response", out var resp))
        {
            return CleanOutput(resp.GetString());
        }

        return null;
    }

    private string? CallOpenAiCompatible(Uri uri, string prompt)
    {
        var payload = new
        {
            model = _config.LlmFallback.Model,
            messages = new[]
            {
                new { role = "system", content = SystemPrompt() },
                new { role = "user", content = prompt }
            },
            temperature = 0.2,
            max_tokens = _config.LlmFallback.MaxTokens
        };

        var response = _client.PostAsJsonAsync(uri, payload, JsonOptions).GetAwaiter().GetResult();
        if (!response.IsSuccessStatusCode)
        {
            return null;
        }

        using var stream = response.Content.ReadAsStream();
        using var doc = JsonDocument.Parse(stream);
        if (doc.RootElement.TryGetProperty("choices", out var choices) && choices.GetArrayLength() > 0)
        {
            var first = choices[0];
            if (first.TryGetProperty("message", out var message) && message.TryGetProperty("content", out var content))
            {
                return CleanOutput(content.GetString());
            }
            if (first.TryGetProperty("text", out var text))
            {
                return CleanOutput(text.GetString());
            }
        }

        return null;
    }

    private string? CallLlamaCpp(Uri uri, string prompt)
    {
        var payload = new
        {
            prompt,
            n_predict = _config.LlmFallback.MaxTokens,
            temperature = 0.2,
            stop = new[] { "\n" }
        };

        var response = _client.PostAsJsonAsync(uri, payload, JsonOptions).GetAwaiter().GetResult();
        if (!response.IsSuccessStatusCode)
        {
            return null;
        }

        using var stream = response.Content.ReadAsStream();
        using var doc = JsonDocument.Parse(stream);
        if (doc.RootElement.TryGetProperty("content", out var content))
        {
            return CleanOutput(content.GetString());
        }
        if (doc.RootElement.TryGetProperty("choices", out var choices) && choices.GetArrayLength() > 0)
        {
            var first = choices[0];
            if (first.TryGetProperty("text", out var text))
            {
                return CleanOutput(text.GetString());
            }
        }

        return null;
    }

    private static string BuildPrompt(string input)
    {
        return $"{SystemPrompt()}\n" +
               "Examples:\n" +
               "User: pen notepad plus plus\n" +
               "Command: open notepad plus plus\n" +
               "User: move mouse for 2 munutes\n" +
               "Command: move mouse for 2 minutes\n" +
               "User: apri vs code e crea nuovo file\n" +
               "Command: open vs code and create new file\n" +
               "User: doppio clic su conferma\n" +
               "Command: double click conferma\n" +
               $"User: {input}\n" +
               "Command:";
    }

    private static string SystemPrompt()
    {
        return "You are a typo-tolerant command normalizer for desktop automation. " +
               "Translate user requests into a SINGLE executable command. " +
               "Correct obvious typos (examples: pen->open, munutes->minutes, notepadplusplus->notepad plus plus). " +
               "Use only these verbs/actions: open, find, click, double click, right click, drag <source> to <target>, type, press, save, save as <name> [in <folder>], new tab, close tab, close window, minimize window, maximize window, restore window, switch window, focus <app>, scroll up/down [n], page up, page down, home, end, wait until <text> [for <seconds>], copy, paste, undo, redo, select all, open url <url>, search <query> [on <browser>], browser back/forward/refresh/find in page, notify <text>, clipboard history, volume up/down/mute [n], brightness up/down [n], lock screen, create new file, move mouse for <duration>, jiggle mouse for <duration>. " +
               "Preserve numeric values and duration from user text exactly when present. " +
               "Output only the command in English or Italian. Do not add notes, brackets, explanations, or markdown. If unsure, reply with NONE.";
    }

    private static string? CleanOutput(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw))
        {
            return null;
        }

        var trimmed = raw.Trim();
        trimmed = trimmed.Trim('"', '\'');
        if (trimmed.StartsWith("command:", StringComparison.OrdinalIgnoreCase))
        {
            trimmed = trimmed.Substring("command:".Length).Trim();
        }
        if (trimmed.StartsWith("intent:", StringComparison.OrdinalIgnoreCase))
        {
            trimmed = trimmed.Substring("intent:".Length).Trim();
        }
        if (trimmed.Equals("none", StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        var line = trimmed.Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).FirstOrDefault();
        if (string.IsNullOrWhiteSpace(line))
        {
            return null;
        }

        // Remove model side-notes like: open notepad+plus [in Notepad++].
        var bracketIndex = line.IndexOf('[');
        if (bracketIndex >= 0)
        {
            line = line[..bracketIndex].Trim();
        }

        line = line.TrimEnd('.', ';', ':');
        line = Regex.Replace(line, "\\bnotepad\\+\\+\\b", "notepad plus plus", RegexOptions.IgnoreCase);
        line = Regex.Replace(line, "\\bnotepadplusplus\\b", "notepad plus plus", RegexOptions.IgnoreCase);
        line = line.Replace("notepad+plus", "notepad plus plus", StringComparison.OrdinalIgnoreCase);
        line = Regex.Replace(line, "\\bnotepad\\s+plus\\b(?!\\s+plus)", "notepad plus plus", RegexOptions.IgnoreCase);
        line = string.Join(' ', line.Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries));

        return string.IsNullOrWhiteSpace(line) ? null : line;
    }

    private string ToAuditText(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw))
        {
            return string.Empty;
        }

        return _config.AuditLlmIncludeRawText ? raw : "[redacted]";
    }

    private void WriteAudit(string eventType, string message, object data)
    {
        if (!_config.AuditLlmInteractions)
        {
            return;
        }

        try
        {
            _auditLog.WriteAsync(new AuditEvent
            {
                Timestamp = DateTimeOffset.UtcNow,
                EventType = eventType,
                Message = message,
                Data = data
            }, CancellationToken.None).GetAwaiter().GetResult();
        }
        catch
        {
            // Best effort: never fail intent parsing because audit logging failed.
        }
    }
}
