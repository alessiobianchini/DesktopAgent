using DesktopAgent.Proto;

namespace DesktopAgent.Core.Models;

public sealed class ContextSnapshot
{
    public WindowRef? ActiveWindow { get; set; }
    public UiTree? UiTree { get; set; }
    public ScreenshotResponse? Screenshot { get; set; }
}
