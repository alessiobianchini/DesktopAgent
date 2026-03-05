using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using System.Text;
using Velopack;

namespace DesktopAgent.Tray;

internal static class Program
{
    [STAThread]
    public static void Main(string[] args)
    {
        HookGlobalExceptionLogging();
        WriteStartupLog($"startup begin | pid={Environment.ProcessId} | exe={Environment.ProcessPath}");

        try
        {
            VelopackApp.Build().Run();
        }
        catch (Exception ex)
        {
            // Continue in non-Velopack installs (zip/Inno/etc.), but keep diagnostics.
            WriteStartupLog($"velopack init failed: {ex.Message}");
        }

        try
        {
            BuildAvaloniaApp().StartWithClassicDesktopLifetime(args, ShutdownMode.OnExplicitShutdown);
        }
        catch (Exception ex)
        {
            WriteStartupLog($"fatal startup exception: {ex}");
            throw;
        }
    }

    public static AppBuilder BuildAvaloniaApp()
    {
        return AppBuilder.Configure<App>()
            .UsePlatformDetect()
            .LogToTrace();
    }

    private static void HookGlobalExceptionLogging()
    {
        AppDomain.CurrentDomain.UnhandledException += (_, eventArgs) =>
        {
            if (eventArgs.ExceptionObject is Exception ex)
            {
                WriteStartupLog($"unhandled exception: {ex}");
            }
        };

        TaskScheduler.UnobservedTaskException += (_, eventArgs) =>
        {
            WriteStartupLog($"unobserved task exception: {eventArgs.Exception}");
            eventArgs.SetObserved();
        };
    }

    private static void WriteStartupLog(string message)
    {
        try
        {
            var root = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "DesktopAgent");
            Directory.CreateDirectory(root);
            var path = Path.Combine(root, "tray-startup.log");
            var line = $"[{DateTimeOffset.Now:yyyy-MM-dd HH:mm:ss.fff zzz}] {message}{Environment.NewLine}";
            File.AppendAllText(path, line, Encoding.UTF8);
        }
        catch
        {
            // best effort
        }
    }
}
