using DesktopAgent.Core.Abstractions;
using DesktopAgent.Core.Config;
using Microsoft.Extensions.Logging;

namespace DesktopAgent.Core.Services;

public static class OcrEngineFactory
{
    public static IOcrEngine Create(AgentConfig config, ILoggerFactory loggerFactory)
    {
        if (!config.OcrEnabled)
        {
            return new OcrEngineStub();
        }

        var mode = NormalizeMode(config.Ocr.Engine);
        var tesseractPath = ResolveTesseractPath(config.Ocr.TesseractPath);
        var tesseract = new TesseractOcrEngine(tesseractPath, loggerFactory.CreateLogger<TesseractOcrEngine>());
        var llmVision = new LlmVisionOcrEngine(config, loggerFactory.CreateLogger<LlmVisionOcrEngine>());
        var llmAvailable = config.LlmFallbackEnabled && !string.IsNullOrWhiteSpace(config.LlmFallback.Endpoint);

        return mode switch
        {
            "ai" => llmAvailable
                ? llmVision
                : tesseract,
            "tesseract" => llmAvailable
                ? new ChainedOcrEngine(new IOcrEngine[] { tesseract, llmVision }, loggerFactory.CreateLogger<ChainedOcrEngine>())
                : tesseract,
            _ => llmAvailable
                ? new ChainedOcrEngine(new IOcrEngine[] { tesseract, llmVision }, loggerFactory.CreateLogger<ChainedOcrEngine>())
                : tesseract
        };
    }

    private static string ResolveTesseractPath(string? configured)
    {
        var candidate = (configured ?? string.Empty).Trim();
        if (!string.IsNullOrWhiteSpace(candidate))
        {
            if (Path.IsPathRooted(candidate) && File.Exists(candidate))
            {
                return candidate;
            }

            var fromPath = FindExecutableInPath(candidate);
            if (!string.IsNullOrWhiteSpace(fromPath))
            {
                return fromPath;
            }
        }

        foreach (var knownPath in GetKnownTesseractPaths())
        {
            if (File.Exists(knownPath))
            {
                return knownPath;
            }
        }

        return string.IsNullOrWhiteSpace(candidate) ? "tesseract" : candidate;
    }

    private static IEnumerable<string> GetKnownTesseractPaths()
    {
        if (OperatingSystem.IsWindows())
        {
            yield return @"C:\Program Files\Tesseract-OCR\tesseract.exe";
            yield return @"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe";
        }
        else if (OperatingSystem.IsMacOS())
        {
            yield return "/opt/homebrew/bin/tesseract";
            yield return "/usr/local/bin/tesseract";
            yield return "/usr/bin/tesseract";
        }
        else
        {
            yield return "/usr/bin/tesseract";
            yield return "/usr/local/bin/tesseract";
        }
    }

    private static string? FindExecutableInPath(string executable)
    {
        var path = Environment.GetEnvironmentVariable("PATH");
        if (string.IsNullOrWhiteSpace(path))
        {
            return null;
        }

        var fileNames = new List<string> { executable };
        if (OperatingSystem.IsWindows() && !executable.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
        {
            fileNames.Add($"{executable}.exe");
        }

        foreach (var segment in path.Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            foreach (var fileName in fileNames)
            {
                var fullPath = Path.Combine(segment, fileName);
                if (File.Exists(fullPath))
                {
                    return fullPath;
                }
            }
        }

        return null;
    }

    private static string NormalizeMode(string? mode)
    {
        return (mode ?? string.Empty).Trim().ToLowerInvariant() switch
        {
            "tesseract" => "tesseract",
            "ai" => "ai",
            _ => "auto"
        };
    }
}
