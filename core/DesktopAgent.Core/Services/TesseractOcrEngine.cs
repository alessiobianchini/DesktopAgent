using System.Diagnostics;
using DesktopAgent.Core.Abstractions;
using DesktopAgent.Core.Models;
using DesktopAgent.Proto;
using Microsoft.Extensions.Logging;

namespace DesktopAgent.Core.Services;

public sealed class TesseractOcrEngine : IOcrEngine
{
    private readonly string _tesseractPath;
    private readonly ILogger<TesseractOcrEngine> _logger;

    public TesseractOcrEngine(string tesseractPath, ILogger<TesseractOcrEngine> logger)
    {
        _tesseractPath = tesseractPath;
        _logger = logger;
    }

    public string Name => "tesseract";

    public async Task<IReadOnlyList<OcrTextRegion>> ReadTextAsync(byte[] pngBytes, CancellationToken cancellationToken)
    {
        var tempFile = Path.Combine(Path.GetTempPath(), $"da_ocr_{Guid.NewGuid():N}.png");
        await File.WriteAllBytesAsync(tempFile, pngBytes, cancellationToken);

        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = _tesseractPath,
                Arguments = $"\"{tempFile}\" stdout -l eng tsv",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            using var process = Process.Start(psi);
            if (process == null)
            {
                return Array.Empty<OcrTextRegion>();
            }

            var output = await process.StandardOutput.ReadToEndAsync(cancellationToken);
            var error = await process.StandardError.ReadToEndAsync(cancellationToken);
            await process.WaitForExitAsync(cancellationToken);

            if (process.ExitCode != 0)
            {
                _logger.LogWarning("Tesseract exited with code {Code}: {Error}", process.ExitCode, error);
                return Array.Empty<OcrTextRegion>();
            }

            return ParseTsv(output);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Tesseract OCR failed");
            return Array.Empty<OcrTextRegion>();
        }
        finally
        {
            try
            {
                File.Delete(tempFile);
            }
            catch
            {
                // ignore cleanup errors
            }
        }
    }

    private static IReadOnlyList<OcrTextRegion> ParseTsv(string tsv)
    {
        var lines = tsv.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        if (lines.Length <= 1)
        {
            return Array.Empty<OcrTextRegion>();
        }

        var regions = new List<OcrTextRegion>();
        for (var i = 1; i < lines.Length; i++)
        {
            var parts = lines[i].Split('\t');
            if (parts.Length < 12)
            {
                continue;
            }

            var text = parts[11].Trim();
            if (string.IsNullOrWhiteSpace(text))
            {
                continue;
            }

            if (!int.TryParse(parts[6], out var left)) continue;
            if (!int.TryParse(parts[7], out var top)) continue;
            if (!int.TryParse(parts[8], out var width)) continue;
            if (!int.TryParse(parts[9], out var height)) continue;
            _ = float.TryParse(parts[10], out var confidence);

            regions.Add(new OcrTextRegion
            {
                Text = text,
                Confidence = confidence,
                Bounds = new Rect { X = left, Y = top, Width = width, Height = height }
            });
        }

        return regions;
    }
}
