using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;

namespace DesktopAgent.Tray;

internal sealed record PluginSetupProbe(
    bool WingetAvailable,
    bool FfmpegInstalled,
    bool TesseractInstalled,
    bool TesseractPrimaryPackageAvailable,
    bool TesseractFallbackPackageAvailable)
{
    public bool HasMissingTools => !FfmpegInstalled || !TesseractInstalled;
}

internal sealed record PluginSetupChoice(
    bool InstallRequested,
    bool InstallFfmpeg,
    bool InstallOcr,
    bool DontAskAgain);

internal sealed class PluginSetupWindow : Window
{
    private readonly CheckBox _ffmpeg;
    private readonly CheckBox _ocr;
    private readonly CheckBox _dontAskAgain;

    public PluginSetupWindow(PluginSetupProbe probe)
    {
        Title = "DesktopAgent Optional Plugins";
        Width = 560;
        Height = 330;
        MinWidth = 520;
        MinHeight = 280;
        WindowStartupLocation = WindowStartupLocation.CenterOwner;
        CanResize = false;

        var root = new Grid
        {
            RowDefinitions = new RowDefinitions("Auto,Auto,Auto,Auto,*"),
            Margin = new Thickness(16)
        };

        root.Children.Add(new TextBlock
        {
            Text = "Install optional local tools for advanced features.",
            FontSize = 16,
            FontWeight = Avalonia.Media.FontWeight.SemiBold
        });

        var info = new TextBlock
        {
            Text = probe.WingetAvailable
                ? "You can install now, or later from Tray menu: Install Optional Plugins."
                : "winget not detected. Automatic plugin install is unavailable on this machine.",
            Margin = new Thickness(0, 8, 0, 0),
            TextWrapping = Avalonia.Media.TextWrapping.Wrap
        };
        Grid.SetRow(info, 1);
        root.Children.Add(info);

        var pluginPanel = new StackPanel { Margin = new Thickness(0, 12, 0, 0), Spacing = 8 };
        _ffmpeg = new CheckBox
        {
            Content = probe.FfmpegInstalled
                ? "FFmpeg (already installed)"
                : "Install FFmpeg (screen recording)",
            IsEnabled = probe.WingetAvailable && !probe.FfmpegInstalled,
            IsChecked = probe.WingetAvailable && !probe.FfmpegInstalled
        };
        pluginPanel.Children.Add(_ffmpeg);

        var ocrCaption = probe.TesseractInstalled
            ? "OCR / Tesseract (already installed)"
            : "Install OCR / Tesseract (vision fallback)";
        _ocr = new CheckBox
        {
            Content = ocrCaption,
            IsEnabled = probe.WingetAvailable && !probe.TesseractInstalled,
            IsChecked = probe.WingetAvailable && !probe.TesseractInstalled
        };
        pluginPanel.Children.Add(_ocr);

        if (probe.WingetAvailable && !probe.TesseractInstalled)
        {
            var packageHint = probe.TesseractPrimaryPackageAvailable
                ? "Package source: UB-Mannheim.TesseractOCR"
                : (probe.TesseractFallbackPackageAvailable
                    ? "Package source: tesseract-ocr.tesseract (fallback)"
                    : "No known winget OCR package detected.");
            pluginPanel.Children.Add(new TextBlock
            {
                Text = packageHint,
                Margin = new Thickness(20, -2, 0, 0),
                Opacity = 0.8
            });
        }

        Grid.SetRow(pluginPanel, 2);
        root.Children.Add(pluginPanel);

        _dontAskAgain = new CheckBox
        {
            Content = "Don't show this wizard again",
            Margin = new Thickness(0, 10, 0, 0),
            IsChecked = false
        };
        Grid.SetRow(_dontAskAgain, 3);
        root.Children.Add(_dontAskAgain);

        var buttonPanel = new StackPanel
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right,
            Spacing = 8
        };

        var skip = new Button
        {
            Content = "Skip",
            MinWidth = 90
        };
        skip.Click += (_, _) => Close(new PluginSetupChoice(false, false, false, _dontAskAgain.IsChecked == true));
        buttonPanel.Children.Add(skip);

        var install = new Button
        {
            Content = "Install Selected",
            MinWidth = 130,
            IsEnabled = probe.WingetAvailable
        };
        install.Click += (_, _) => Close(new PluginSetupChoice(
            true,
            _ffmpeg.IsEnabled && _ffmpeg.IsChecked == true,
            _ocr.IsEnabled && _ocr.IsChecked == true,
            true));
        buttonPanel.Children.Add(install);

        Grid.SetRow(buttonPanel, 4);
        root.Children.Add(buttonPanel);

        Content = root;
    }

    public static async Task ShowResultAsync(string message, Window? owner)
    {
        var window = new Window
        {
            Title = "Plugin Setup Result",
            Width = 500,
            Height = 240,
            MinWidth = 460,
            MinHeight = 200,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            CanResize = false
        };

        var text = new TextBlock
        {
            Text = message,
            TextWrapping = Avalonia.Media.TextWrapping.Wrap,
            Margin = new Thickness(16)
        };

        var ok = new Button
        {
            Content = "OK",
            HorizontalAlignment = HorizontalAlignment.Right,
            Margin = new Thickness(0, 8, 16, 16),
            MinWidth = 80
        };
        ok.Click += (_, _) => window.Close();

        var panel = new Grid { RowDefinitions = new RowDefinitions("*,Auto") };
        panel.Children.Add(text);
        Grid.SetRow(ok, 1);
        panel.Children.Add(ok);
        window.Content = panel;

        if (owner != null)
        {
            await window.ShowDialog(owner);
        }
        else
        {
            window.Show();
        }
    }
}
