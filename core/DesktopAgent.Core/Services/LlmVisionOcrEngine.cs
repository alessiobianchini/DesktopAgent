using System.Net.Http.Json;
using System.Text.Json;
using DesktopAgent.Core.Abstractions;
using DesktopAgent.Core.Config;
using DesktopAgent.Core.Models;
using DesktopAgent.Proto;
using Microsoft.Extensions.Logging;

namespace DesktopAgent.Core.Services;

public sealed class LlmVisionOcrEngine : IOcrEngine
{
    private static readonly JsonSerializerOptions JsonOptions = new(JsonSerializerDefaults.Web);
    private readonly AgentConfig _config;
    private readonly ILogger<LlmVisionOcrEngine> _logger;

    public LlmVisionOcrEngine(AgentConfig config, ILogger<LlmVisionOcrEngine> logger)
    {
        _config = config;
        _logger = logger;
    }

    public string Name => "llm-vision";

    public async Task<IReadOnlyList<OcrTextRegion>> ReadTextAsync(byte[] pngBytes, CancellationToken cancellationToken)
    {
        if (!_config.LlmFallbackEnabled || pngBytes.Length == 0)
        {
            return Array.Empty<OcrTextRegion>();
        }

        var endpoint = _config.LlmFallback.Endpoint?.Trim();
        if (string.IsNullOrWhiteSpace(endpoint) || !Uri.TryCreate(endpoint, UriKind.Absolute, out var uri))
        {
            return Array.Empty<OcrTextRegion>();
        }

        if (!uri.IsLoopback && !_config.AllowNonLoopbackLlmEndpoint)
        {
            return Array.Empty<OcrTextRegion>();
        }

        var provider = (_config.LlmFallback.Provider ?? "ollama").Trim().ToLowerInvariant();
        var normalizedUri = NormalizeVisionEndpoint(uri, provider);
        var (width, height) = TryReadPngDimensions(pngBytes);
        var prompt = BuildVisionPrompt(width, height);
        var maxTokens = Math.Clamp(Math.Max(256, _config.LlmFallback.MaxTokens * 3), 256, 4096);

        try
        {
            using var client = new HttpClient { Timeout = TimeSpan.FromSeconds(Math.Max(4, _config.LlmFallback.TimeoutSeconds)) };
            var raw = provider switch
            {
                "openai" => await CallOpenAiVisionAsync(client, normalizedUri, prompt, pngBytes, maxTokens, cancellationToken),
                "ollama" => await CallOllamaVisionAsync(client, normalizedUri, prompt, pngBytes, maxTokens, cancellationToken),
                _ => null
            };

            if (string.IsNullOrWhiteSpace(raw))
            {
                return Array.Empty<OcrTextRegion>();
            }

            var regions = ParseRegions(raw, width, height);
            return regions.Count == 0 ? Array.Empty<OcrTextRegion>() : regions;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "LLM vision OCR failed");
            return Array.Empty<OcrTextRegion>();
        }
    }

    private async Task<string?> CallOllamaVisionAsync(
        HttpClient client,
        Uri endpoint,
        string prompt,
        byte[] pngBytes,
        int maxTokens,
        CancellationToken cancellationToken)
    {
        var payload = new
        {
            model = _config.LlmFallback.Model,
            stream = false,
            format = "json",
            options = new { temperature = 0, num_predict = maxTokens },
            messages = new[]
            {
                new
                {
                    role = "user",
                    content = prompt,
                    images = new[] { Convert.ToBase64String(pngBytes) }
                }
            }
        };

        using var response = await client.PostAsJsonAsync(endpoint, payload, JsonOptions, cancellationToken);
        if (!response.IsSuccessStatusCode)
        {
            return null;
        }

        await using var stream = await response.Content.ReadAsStreamAsync(cancellationToken);
        using var doc = await JsonDocument.ParseAsync(stream, cancellationToken: cancellationToken);
        if (doc.RootElement.TryGetProperty("message", out var message)
            && message.TryGetProperty("content", out var content))
        {
            return content.GetString();
        }

        if (doc.RootElement.TryGetProperty("response", out var responseContent))
        {
            return responseContent.GetString();
        }

        return null;
    }

    private async Task<string?> CallOpenAiVisionAsync(
        HttpClient client,
        Uri endpoint,
        string prompt,
        byte[] pngBytes,
        int maxTokens,
        CancellationToken cancellationToken)
    {
        var base64 = Convert.ToBase64String(pngBytes);
        var payload = new
        {
            model = _config.LlmFallback.Model,
            temperature = 0,
            max_tokens = maxTokens,
            response_format = new { type = "json_object" },
            messages = new object[]
            {
                new
                {
                    role = "system",
                    content = new object[]
                    {
                        new { type = "text", text = "You are an OCR engine. Return strict JSON only." }
                    }
                },
                new
                {
                    role = "user",
                    content = new object[]
                    {
                        new { type = "text", text = prompt },
                        new
                        {
                            type = "image_url",
                            image_url = new
                            {
                                url = $"data:image/png;base64,{base64}"
                            }
                        }
                    }
                }
            }
        };

        using var response = await client.PostAsJsonAsync(endpoint, payload, JsonOptions, cancellationToken);
        if (!response.IsSuccessStatusCode)
        {
            return null;
        }

        await using var stream = await response.Content.ReadAsStreamAsync(cancellationToken);
        using var doc = await JsonDocument.ParseAsync(stream, cancellationToken: cancellationToken);
        if (!doc.RootElement.TryGetProperty("choices", out var choices) || choices.GetArrayLength() == 0)
        {
            return null;
        }

        var first = choices[0];
        if (!first.TryGetProperty("message", out var message))
        {
            return null;
        }

        if (!message.TryGetProperty("content", out var content))
        {
            return null;
        }

        return content.GetString();
    }

    private static Uri NormalizeVisionEndpoint(Uri uri, string provider)
    {
        if (provider == "openai")
        {
            if (uri.AbsolutePath.Equals("/", StringComparison.Ordinal))
            {
                return new Uri(uri, "/v1/chat/completions");
            }

            if (uri.AbsolutePath.EndsWith("/v1/chat/completions", StringComparison.OrdinalIgnoreCase))
            {
                return uri;
            }

            return uri;
        }

        if (provider == "ollama")
        {
            var root = new Uri(uri.GetLeftPart(UriPartial.Authority));
            return new Uri(root, "/api/chat");
        }

        return uri;
    }

    private static string BuildVisionPrompt(int width, int height)
    {
        var sizeHint = width > 0 && height > 0 ? $"Image size: {width}x{height} pixels." : "Image size unknown.";
        return
            "Perform OCR on this screenshot.\n" +
            $"{sizeHint}\n" +
            "Return STRICT JSON only with schema:\n" +
            "{\"regions\":[{\"text\":\"...\",\"x\":0,\"y\":0,\"width\":0,\"height\":0,\"confidence\":0.0}]}\n" +
            "Coordinates must be pixel values in screenshot space. Include only readable text.";
    }

    private static List<OcrTextRegion> ParseRegions(string raw, int imageWidth, int imageHeight)
    {
        var json = TryExtractJson(raw);
        if (string.IsNullOrWhiteSpace(json))
        {
            return new List<OcrTextRegion>();
        }

        try
        {
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;
            var regionElements = new List<JsonElement>();
            if (root.ValueKind == JsonValueKind.Array)
            {
                regionElements.AddRange(root.EnumerateArray());
            }
            else if (root.ValueKind == JsonValueKind.Object)
            {
                if (root.TryGetProperty("regions", out var regions) && regions.ValueKind == JsonValueKind.Array)
                {
                    regionElements.AddRange(regions.EnumerateArray());
                }
                else
                {
                    regionElements.Add(root);
                }
            }

            var results = new List<OcrTextRegion>();
            foreach (var element in regionElements)
            {
                var text = ReadString(element, "text");
                if (string.IsNullOrWhiteSpace(text))
                {
                    continue;
                }

                if (!TryReadInt(element, "x", out var x)) continue;
                if (!TryReadInt(element, "y", out var y)) continue;
                if (!TryReadInt(element, "width", out var width)) continue;
                if (!TryReadInt(element, "height", out var height)) continue;

                if (width <= 0 || height <= 0)
                {
                    continue;
                }

                if (imageWidth > 0)
                {
                    x = Math.Clamp(x, 0, imageWidth - 1);
                    width = Math.Clamp(width, 1, Math.Max(1, imageWidth - x));
                }

                if (imageHeight > 0)
                {
                    y = Math.Clamp(y, 0, imageHeight - 1);
                    height = Math.Clamp(height, 1, Math.Max(1, imageHeight - y));
                }

                var confidence = ReadFloat(element, "confidence");
                if (confidence > 1.0f)
                {
                    confidence /= 100f;
                }

                results.Add(new OcrTextRegion
                {
                    Text = text.Trim(),
                    Confidence = Math.Clamp(confidence, 0f, 1f),
                    Bounds = new Rect { X = x, Y = y, Width = width, Height = height }
                });
            }

            return results;
        }
        catch
        {
            return new List<OcrTextRegion>();
        }
    }

    private static bool TryReadInt(JsonElement element, string name, out int value)
    {
        value = 0;
        if (!element.TryGetProperty(name, out var prop))
        {
            return false;
        }

        return prop.ValueKind switch
        {
            JsonValueKind.Number => prop.TryGetInt32(out value),
            JsonValueKind.String when int.TryParse(prop.GetString(), out var parsed) => (value = parsed) == parsed,
            _ => false
        };
    }

    private static float ReadFloat(JsonElement element, string name)
    {
        if (!element.TryGetProperty(name, out var prop))
        {
            return 0f;
        }

        return prop.ValueKind switch
        {
            JsonValueKind.Number => prop.TryGetSingle(out var number) ? number : 0f,
            JsonValueKind.String when float.TryParse(prop.GetString(), out var parsed) => parsed,
            _ => 0f
        };
    }

    private static string ReadString(JsonElement element, string name)
    {
        return element.TryGetProperty(name, out var prop) && prop.ValueKind == JsonValueKind.String
            ? prop.GetString() ?? string.Empty
            : string.Empty;
    }

    private static string? TryExtractJson(string raw)
    {
        var text = (raw ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(text))
        {
            return null;
        }

        if ((text.StartsWith("{", StringComparison.Ordinal) && text.EndsWith("}", StringComparison.Ordinal))
            || (text.StartsWith("[", StringComparison.Ordinal) && text.EndsWith("]", StringComparison.Ordinal)))
        {
            return text;
        }

        var objectStart = text.IndexOf('{');
        var objectEnd = text.LastIndexOf('}');
        if (objectStart >= 0 && objectEnd > objectStart)
        {
            return text.Substring(objectStart, objectEnd - objectStart + 1);
        }

        var arrayStart = text.IndexOf('[');
        var arrayEnd = text.LastIndexOf(']');
        if (arrayStart >= 0 && arrayEnd > arrayStart)
        {
            return text.Substring(arrayStart, arrayEnd - arrayStart + 1);
        }

        return null;
    }

    private static (int Width, int Height) TryReadPngDimensions(byte[] pngBytes)
    {
        if (pngBytes.Length < 24)
        {
            return (0, 0);
        }

        var signature = new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 };
        for (var i = 0; i < signature.Length; i++)
        {
            if (pngBytes[i] != signature[i])
            {
                return (0, 0);
            }
        }

        // IHDR width/height are at bytes 16..23 (big-endian)
        var width = (pngBytes[16] << 24) | (pngBytes[17] << 16) | (pngBytes[18] << 8) | pngBytes[19];
        var height = (pngBytes[20] << 24) | (pngBytes[21] << 16) | (pngBytes[22] << 8) | pngBytes[23];
        if (width <= 0 || height <= 0)
        {
            return (0, 0);
        }

        return (width, height);
    }
}

