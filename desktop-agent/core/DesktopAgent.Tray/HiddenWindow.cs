using Avalonia;
using Avalonia.Controls;

namespace DesktopAgent.Tray;

internal sealed class HiddenWindow : Window
{
    public HiddenWindow()
    {
        Width = 1;
        Height = 1;
        Opacity = 0;
        ShowInTaskbar = false;
        CanResize = false;
        SystemDecorations = SystemDecorations.None;
        Position = new PixelPoint(-32000, -32000);
    }

    protected override void OnOpened(EventArgs e)
    {
        base.OnOpened(e);
        Hide();
    }
}
