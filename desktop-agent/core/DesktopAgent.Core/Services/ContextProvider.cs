using DesktopAgent.Core.Abstractions;
using DesktopAgent.Core.Config;
using DesktopAgent.Core.Models;
using DesktopAgent.Proto;
using Microsoft.Extensions.Logging;

namespace DesktopAgent.Core.Services;

public sealed class ContextProvider : IContextProvider
{
    private readonly IDesktopAdapterClient _client;
    private readonly IOcrEngine _ocrEngine;
    private readonly AgentConfig _config;
    private readonly ILogger<ContextProvider> _logger;

    public ContextProvider(IDesktopAdapterClient client, IOcrEngine ocrEngine, AgentConfig config, ILogger<ContextProvider> logger)
    {
        _client = client;
        _ocrEngine = ocrEngine;
        _config = config;
        _logger = logger;
    }

    public async Task<ContextSnapshot> GetSnapshotAsync(CancellationToken cancellationToken)
    {
        var snapshot = new ContextSnapshot();
        var activeWindow = await _client.GetActiveWindowAsync(cancellationToken);
        snapshot.ActiveWindow = activeWindow;

        if (activeWindow != null)
        {
            var uiTree = await _client.GetUiTreeAsync(activeWindow.Id, cancellationToken);
            snapshot.UiTree = uiTree;

            if (uiTree == null)
            {
                snapshot.Screenshot = await SafeCaptureAsync(activeWindow.Id, cancellationToken);
            }
        }

        return snapshot;
    }

    public async Task<FindResult> FindByTextAsync(string text, CancellationToken cancellationToken)
    {
        var result = new FindResult();
        var selector = new Selector { NameContains = text };
        var elements = await _client.FindElementsAsync(selector, cancellationToken);
        result.Elements.AddRange(elements);

        if (result.Elements.Count > 0)
        {
            return result;
        }

        if (!_config.OcrEnabled)
        {
            return result;
        }

        var screenshot = await SafeCaptureAsync(windowId: string.Empty, cancellationToken);
        if (screenshot == null || screenshot.Png.IsEmpty)
        {
            return result;
        }

        try
        {
            var regions = await _ocrEngine.ReadTextAsync(screenshot.Png.ToByteArray(), cancellationToken);
            var matches = regions
                .Where(r => r.Text.Contains(text, StringComparison.OrdinalIgnoreCase))
                .ToList();
            result.OcrMatches.AddRange(matches);
            return result;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "OCR failed");
            return result;
        }
    }

    private async Task<ScreenshotResponse?> SafeCaptureAsync(string windowId, CancellationToken cancellationToken)
    {
        try
        {
            var request = new ScreenshotRequest { WindowId = windowId };
            return await _client.CaptureScreenAsync(request, cancellationToken);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "CaptureScreen failed");
            return null;
        }
    }
}
