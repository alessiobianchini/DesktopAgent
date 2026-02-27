using DesktopAgent.Proto;

namespace DesktopAgent.Core.Models;

public sealed class OcrTextRegion
{
    public string Text { get; set; } = string.Empty;
    public Rect Bounds { get; set; } = new();
    public float Confidence { get; set; }
}

public sealed class FindResult
{
    public List<ElementRef> Elements { get; set; } = new();
    public List<OcrTextRegion> OcrMatches { get; set; } = new();
}
