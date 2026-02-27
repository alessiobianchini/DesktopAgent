using DesktopAgent.Core.Abstractions;
using DesktopAgent.Core.Models;

namespace DesktopAgent.Core.Services;

public sealed class OcrEngineStub : IOcrEngine
{
    public string Name => "stub";

    public Task<IReadOnlyList<OcrTextRegion>> ReadTextAsync(byte[] pngBytes, CancellationToken cancellationToken)
    {
        return Task.FromResult<IReadOnlyList<OcrTextRegion>>(Array.Empty<OcrTextRegion>());
    }
}
