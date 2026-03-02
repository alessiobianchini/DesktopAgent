using DesktopAgent.Core.Models;

namespace DesktopAgent.Core.Abstractions;

public interface IOcrEngine
{
    Task<IReadOnlyList<OcrTextRegion>> ReadTextAsync(byte[] pngBytes, CancellationToken cancellationToken);
    string Name { get; }
}
