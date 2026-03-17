using DesktopAgent.Core.Abstractions;
using DesktopAgent.Core.Models;
using Microsoft.Extensions.Logging;

namespace DesktopAgent.Core.Services;

public sealed class ChainedOcrEngine : IOcrEngine
{
    private readonly IReadOnlyList<IOcrEngine> _engines;
    private readonly ILogger<ChainedOcrEngine> _logger;

    public ChainedOcrEngine(IEnumerable<IOcrEngine> engines, ILogger<ChainedOcrEngine> logger)
    {
        _engines = engines?.Where(engine => engine != null).ToList() ?? new List<IOcrEngine>();
        _logger = logger;
    }

    public string Name => _engines.Count == 0
        ? "chain-empty"
        : $"chain({string.Join("->", _engines.Select(engine => engine.Name))})";

    public async Task<IReadOnlyList<OcrTextRegion>> ReadTextAsync(byte[] pngBytes, CancellationToken cancellationToken)
    {
        if (_engines.Count == 0)
        {
            return Array.Empty<OcrTextRegion>();
        }

        foreach (var engine in _engines)
        {
            try
            {
                var regions = await engine.ReadTextAsync(pngBytes, cancellationToken);
                if (regions.Count > 0)
                {
                    return regions;
                }
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "OCR engine {Engine} failed, trying next fallback.", engine.Name);
            }
        }

        return Array.Empty<OcrTextRegion>();
    }
}

