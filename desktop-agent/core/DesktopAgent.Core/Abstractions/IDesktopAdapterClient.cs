using DesktopAgent.Proto;

namespace DesktopAgent.Core.Abstractions;

public interface IDesktopAdapterClient
{
    Task<WindowRef?> GetActiveWindowAsync(CancellationToken cancellationToken);
    Task<IReadOnlyList<WindowRef>> ListWindowsAsync(CancellationToken cancellationToken);
    Task<UiTree?> GetUiTreeAsync(string windowId, CancellationToken cancellationToken);
    Task<IReadOnlyList<ElementRef>> FindElementsAsync(Selector selector, CancellationToken cancellationToken);
    Task<ActionResult> InvokeElementAsync(string elementId, CancellationToken cancellationToken);
    Task<ActionResult> SetElementValueAsync(string elementId, string value, CancellationToken cancellationToken);
    Task<ActionResult> ClickPointAsync(int x, int y, CancellationToken cancellationToken);
    Task<ActionResult> TypeTextAsync(string text, CancellationToken cancellationToken);
    Task<ActionResult> KeyComboAsync(IEnumerable<string> keys, CancellationToken cancellationToken);
    Task<ActionResult> OpenAppAsync(string appIdOrPath, CancellationToken cancellationToken);
    Task<ScreenshotResponse> CaptureScreenAsync(ScreenshotRequest request, CancellationToken cancellationToken);
    Task<ClipboardResponse> GetClipboardAsync(CancellationToken cancellationToken);
    Task<ActionResult> SetClipboardAsync(string text, CancellationToken cancellationToken);
    Task<Status> ArmAsync(bool requireUserPresence, CancellationToken cancellationToken);
    Task<Status> DisarmAsync(CancellationToken cancellationToken);
    Task<Status> GetStatusAsync(CancellationToken cancellationToken);
}
