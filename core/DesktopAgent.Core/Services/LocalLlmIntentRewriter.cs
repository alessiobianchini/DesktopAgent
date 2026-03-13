using System.Diagnostics;
using System.Globalization;
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

    public LlmRewriteResult? Rewrite(string input)
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

        if (!Uri.TryCreate(endpoint, UriKind.Absolute, out var uri)
            || (!uri.IsLoopback && !_config.AllowNonLoopbackLlmEndpoint))
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
            string? rawOutput;
            if (provider == "ollama")
            {
                rawOutput = CallOllama(uri, prompt);
            }
            else if (provider == "openai")
            {
                rawOutput = CallOpenAiCompatible(uri, prompt);
            }
            else if (provider == "llama.cpp")
            {
                rawOutput = CallLlamaCpp(uri, prompt);
            }
            else
            {
                rawOutput = CallOllama(uri, prompt);
            }

            var rewritten = ParseResult(rawOutput);
            stopwatch.Stop();
            WriteAudit("llm_response", "LLM rewrite completed", new
            {
                provider,
                model = _config.LlmFallback.Model,
                latencyMs = stopwatch.ElapsedMilliseconds,
                success = rewritten != null,
                output = ToAuditText(rewritten?.RawOutput),
                translatedCommand = rewritten?.Command ?? string.Empty,
                confidence = rewritten?.Confidence ?? 0,
                needsClarification = rewritten?.NeedsClarification ?? false,
                outputLength = rewritten?.Command.Length ?? 0
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
            format = "json",
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
            return resp.GetString();
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
            max_tokens = _config.LlmFallback.MaxTokens,
            response_format = new { type = "json_object" }
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
                return content.GetString();
            }
            if (first.TryGetProperty("text", out var text))
            {
                return text.GetString();
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
            return content.GetString();
        }
        if (doc.RootElement.TryGetProperty("choices", out var choices) && choices.GetArrayLength() > 0)
        {
            var first = choices[0];
            if (first.TryGetProperty("text", out var text))
            {
                return text.GetString();
            }
        }

        return null;
    }

    private static string BuildPrompt(string input)
    {
        return $"{SystemPrompt()}\n" +
               "Output schema (strict JSON, no extra keys):\n" +
               "{\"command\":\"<normalized command>\",\"confidence\":0.0,\"needs_clarification\":false,\"clarification_question\":\"\"}\n" +
               "Examples:\n" +
               "User: pen notepad plus plus\n" +
               "{\"command\":\"open notepad plus plus\",\"confidence\":0.93,\"needs_clarification\":false,\"clarification_question\":\"\"}\n" +
               "User: move mouse for 2 munutes\n" +
               "{\"command\":\"move mouse for 2 minutes\",\"confidence\":0.95,\"needs_clarification\":false,\"clarification_question\":\"\"}\n" +
               "User: apri vs code e crea nuovo file\n" +
               "{\"command\":\"open vs code and then create new file\",\"confidence\":0.90,\"needs_clarification\":false,\"clarification_question\":\"\"}\n" +
               "User: puoi fare una cattura schermo e aprire notepad?\n" +
               "{\"command\":\"take screenshot and then open notepad\",\"confidence\":0.94,\"needs_clarification\":false,\"clarification_question\":\"\"}\n" +
               "User: apri teams\n" +
               "{\"command\":\"open teams\",\"confidence\":0.84,\"needs_clarification\":false,\"clarification_question\":\"\"}\n" +
               "User: open the chat app maybe teams or slack\n" +
               "{\"command\":\"open teams\",\"confidence\":0.52,\"needs_clarification\":true,\"clarification_question\":\"Do you want Teams or Slack?\"}\n" +
               $"User: {input}\n" +
               "JSON:";
    }

    private static string SystemPrompt()
    {
        return "You are a typo-tolerant command normalizer for desktop automation. " +
               "Translate user requests into executable commands. " +
               "Correct obvious typos (examples: pen->open, munutes->minutes, notepadplusplus->notepad plus plus). " +
               "Use only these verbs/actions: open, find, click, double click, right click, drag <source> to <target>, type, press, save, save as <name> [in <folder>], new tab, close tab, close window, minimize window, maximize window, restore window, switch window, focus <app>, scroll up/down [n], page up, page down, home, end, wait until <text> [for <seconds>], copy, paste, undo, redo, select all, open url <url>, search <query> [on <browser>], browser back/forward/refresh/find in page, notify <text>, clipboard history, volume up/down/mute [n], brightness up/down [n], lock screen, create new file, move mouse for <duration>, jiggle mouse for <duration>, record screen [and audio] for <duration>, start recording [screen] [with/without audio], stop recording, take screenshot [for each screen|single-screen]. " +
               "If there are multiple actions, output them in sequence using ' and then ' as separator. " +
               "When app is implied, infer the most likely app token (examples: teams, chrome, edge, vscode). " +
               "Preserve numeric values and duration from user text exactly when present. " +
               "Always output strict JSON only, with keys: command, confidence, needs_clarification, clarification_question. " +
               "Do not output markdown or prose. If unsure, set needs_clarification=true and confidence<0.60.";
    }

    private LlmRewriteResult? ParseResult(string? rawOutput)
    {
        var raw = rawOutput?.Trim();
        if (string.IsNullOrWhiteSpace(raw))
        {
            return null;
        }

        var parsed = TryParseJsonResult(raw);
        if (parsed != null)
        {
            return parsed;
        }

        var command = CleanOutput(raw);
        if (string.IsNullOrWhiteSpace(command))
        {
            return null;
        }

        return new LlmRewriteResult(command, 0.60, false, null, raw);
    }

    private static LlmRewriteResult? TryParseJsonResult(string raw)
    {
        var candidate = ExtractJsonPayload(raw);
        if (string.IsNullOrWhiteSpace(candidate))
        {
            return null;
        }

        try
        {
            using var doc = JsonDocument.Parse(candidate);
            var root = doc.RootElement;
            var command = root.TryGetProperty("command", out var cmd) ? cmd.GetString() : null;
            command = CleanOutput(command);
            if (string.IsNullOrWhiteSpace(command))
            {
                return null;
            }

            var confidence = 0.6;
            if (root.TryGetProperty("confidence", out var conf))
            {
                confidence = conf.ValueKind switch
                {
                    JsonValueKind.Number => conf.GetDouble(),
                    JsonValueKind.String when double.TryParse(conf.GetString(), NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed)
                        => parsed,
                    _ => confidence
                };
            }

            confidence = Math.Clamp(confidence, 0.0, 1.0);
            var needsClarification = root.TryGetProperty("needs_clarification", out var needs) && needs.ValueKind == JsonValueKind.True;
            var clarification = root.TryGetProperty("clarification_question", out var question)
                ? question.GetString()
                : null;

            return new LlmRewriteResult(command, confidence, needsClarification, clarification, raw);
        }
        catch
        {
            return null;
        }
    }

    private static string? ExtractJsonPayload(string raw)
    {
        var trimmed = raw.Trim();
        if (trimmed.StartsWith('{') && trimmed.EndsWith('}'))
        {
            return trimmed;
        }

        var start = trimmed.IndexOf('{');
        var end = trimmed.LastIndexOf('}');
        if (start < 0 || end <= start)
        {
            return null;
        }

        return trimmed.Substring(start, end - start + 1);
    }

    private static string? CleanOutput(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw))
        {
            return null;
        }

        var trimmed = raw.Trim();
        trimmed = ExtractCommandFromJson(trimmed) ?? trimmed;
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

        var lines = trimmed.Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Select(line => Regex.Replace(line, "^[\\-\\*\\d\\)\\.\\s]+", string.Empty).Trim())
            .Where(line => !string.IsNullOrWhiteSpace(line))
            .ToList();
        if (lines.Count == 0)
        {
            return null;
        }

        var line = string.Join(" and then ", lines);

        // Remove model side-notes like: open notepad+plus [in Notepad++].
        var bracketIndex = line.IndexOf('[');
        if (bracketIndex >= 0)
        {
            line = line[..bracketIndex].Trim();
        }

        line = Regex.Replace(line, "\\s*(?:;|->|=>|\\|)\\s*", " and then ", RegexOptions.IgnoreCase);
        line = Regex.Replace(line, "\\s*,\\s*(?=(open|find|click|double click|right click|drag|type|press|save|new tab|close tab|close window|minimize|maximize|restore|switch window|focus|scroll|page up|page down|home|end|wait until|copy|paste|undo|redo|select all|open url|search|browser back|browser forward|refresh|find in page|notify|clipboard history|volume|brightness|lock screen|create new file|move mouse|jiggle mouse|record screen|start recording|stop recording|take screenshot|snapshot)\\b)", " and then ", RegexOptions.IgnoreCase);
        line = line.TrimEnd('.', ';', ':');
        line = ExtractLikelyCommandSpan(line);
        line = Regex.Replace(line, "\\bnotepad\\+\\+\\b", "notepad plus plus", RegexOptions.IgnoreCase);
        line = Regex.Replace(line, "\\bnotepadplusplus\\b", "notepad plus plus", RegexOptions.IgnoreCase);
        line = line.Replace("notepad+plus", "notepad plus plus", StringComparison.OrdinalIgnoreCase);
        line = Regex.Replace(line, "\\bnotepad\\s+plus\\b(?!\\s+plus)", "notepad plus plus", RegexOptions.IgnoreCase);
        line = Regex.Replace(line, "\\bms\\s*teams\\b", "teams", RegexOptions.IgnoreCase);
        line = Regex.Replace(line, "\\bvisual\\s+studio\\s+code\\b", "vs code", RegexOptions.IgnoreCase);
        line = string.Join(' ', line.Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries));

        return string.IsNullOrWhiteSpace(line) ? null : line;
    }

    private static string? ExtractCommandFromJson(string text)
    {
        var value = text.Trim();
        if (!value.StartsWith('{'))
        {
            return null;
        }

        try
        {
            using var doc = JsonDocument.Parse(value);
            if (doc.RootElement.TryGetProperty("command", out var command))
            {
                return command.GetString();
            }
            if (doc.RootElement.TryGetProperty("intent", out var intent))
            {
                return intent.GetString();
            }
        }
        catch
        {
            return null;
        }

        return null;
    }

    private static string ExtractLikelyCommandSpan(string line)
    {
        var commandStart = Regex.Match(
            line,
            "(open|find|click|double click|right click|drag|type|press|save|new tab|close tab|close window|minimize|maximize|restore|switch window|focus|scroll|page up|page down|home|end|wait until|copy|paste|undo|redo|select all|open url|search|browser back|browser forward|refresh|find in page|notify|clipboard history|volume|brightness|lock screen|create new file|move mouse|jiggle mouse|record screen|start recording|stop recording|take screenshot|snapshot)\\b",
            RegexOptions.IgnoreCase);

        if (!commandStart.Success || commandStart.Index <= 0)
        {
            return line;
        }

        return line[commandStart.Index..].Trim();
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
