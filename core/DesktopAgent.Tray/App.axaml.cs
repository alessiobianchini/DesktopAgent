using System.Diagnostics;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.ApplicationLifetimes;
using Avalonia.Markup.Xaml;
using Avalonia.Platform;
using Avalonia.Threading;
using DesktopAgent.Core.Services;
using DesktopAgent.Proto;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Velopack;

namespace DesktopAgent.Tray;

public partial class App : Application
{
    private readonly CancellationTokenSource _shutdown = new();
    private TraySettings _settings = new();
    private ILoggerFactory? _loggerFactory;
    private DesktopGrpcClient? _client;
    private WebApiClient? _webApiClient;
    private QuickChatWindow? _quickChatWindow;

    private TrayIcon? _trayIcon;
    private TrayIcons? _trayIcons;
    private NativeMenuItem? _statusItem;
    private NativeMenuItem? _armItem;
    private NativeMenuItem? _disarmItem;
    private NativeMenuItem? _killItem;
    private NativeMenuItem? _resetKillItem;
    private NativeMenuItem? _lockWindowItem;
    private NativeMenuItem? _lockAppItem;
    private NativeMenuItem? _unlockItem;
    private NativeMenuItem? _profileSafeItem;
    private NativeMenuItem? _profileBalancedItem;
    private NativeMenuItem? _profilePowerItem;
    private NativeMenuItem? _refreshItem;
    private NativeMenuItem? _openQuickChatItem;
    private NativeMenuItem? _openWebUiItem;
    private NativeMenuItem? _updateStatusItem;
    private NativeMenuItem? _checkUpdatesItem;
    private NativeMenuItem? _applyUpdateItem;
    private NativeMenuItem? _exitItem;

    private Task? _pollingTask;
    private UpdateManager? _updateManager;
    private VelopackAsset? _pendingUpdate;
    private readonly SemaphoreSlim _updateGate = new(1, 1);
    private DateTimeOffset _lastUpdateCheck = DateTimeOffset.MinValue;
    private string _updateStatus = "updates disabled";
    private bool _updatesEnabled;

    public override void Initialize()
    {
        AvaloniaXamlLoader.Load(this);
    }

    public override void OnFrameworkInitializationCompleted()
    {
        _settings = LoadSettings();
        _loggerFactory = LoggerFactory.Create(logging =>
        {
            logging.SetMinimumLevel(LogLevel.Information);
            logging.AddConsole();
        });

        _client = new DesktopGrpcClient(
            _settings.AdapterEndpoint,
            _loggerFactory.CreateLogger<DesktopGrpcClient>());
        _webApiClient = new WebApiClient(
            _settings.WebUiUrl,
            TimeSpan.FromSeconds(_settings.ApiTimeoutSeconds));
        InitializeUpdater();

        ConfigureTray();

        if (ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
        {
            desktop.MainWindow = new HiddenWindow();
            desktop.Exit += (_, _) => DisposeResources();
        }

        _pollingTask = Task.Run(() => PollStatusAsync(_shutdown.Token));

        base.OnFrameworkInitializationCompleted();
    }

    private void ConfigureTray()
    {
        _statusItem = new NativeMenuItem("Status: initializing...") { IsEnabled = false };
        _updateStatusItem = new NativeMenuItem($"Updates: {_updateStatus}") { IsEnabled = false };
        _refreshItem = new NativeMenuItem("Refresh status");
        _refreshItem.Click += async (_, _) => await RefreshStatusAsync();

        _armItem = new NativeMenuItem("Arm (Require presence)");
        _armItem.Click += async (_, _) =>
        {
            if (_client == null)
            {
                return;
            }

            var status = await _client.ArmAsync(_settings.RequireUserPresenceOnArm, _shutdown.Token);
            UpdateStatusUi(status);
        };

        _disarmItem = new NativeMenuItem("Disarm");
        _disarmItem.Click += async (_, _) =>
        {
            if (_client == null)
            {
                return;
            }

            var status = await _client.DisarmAsync(_shutdown.Token);
            UpdateStatusUi(status);
        };

        _killItem = new NativeMenuItem("Kill Switch ON");
        _killItem.Click += async (_, _) => await RunQuickCommandAsync("kill");

        _resetKillItem = new NativeMenuItem("Kill Switch Reset");
        _resetKillItem.Click += async (_, _) => await RunQuickCommandAsync("reset kill");

        _lockWindowItem = new NativeMenuItem("Lock Current Window");
        _lockWindowItem.Click += async (_, _) => await RunQuickCommandAsync("lock on current window");

        _lockAppItem = new NativeMenuItem("Lock Current App");
        _lockAppItem.Click += async (_, _) => await RunQuickCommandAsync("lock on app");

        _unlockItem = new NativeMenuItem("Unlock Context");
        _unlockItem.Click += async (_, _) => await RunQuickCommandAsync("unlock");

        _profileSafeItem = new NativeMenuItem("Profile: Safe");
        _profileSafeItem.Click += async (_, _) => await RunQuickCommandAsync("profile safe");

        _profileBalancedItem = new NativeMenuItem("Profile: Balanced");
        _profileBalancedItem.Click += async (_, _) => await RunQuickCommandAsync("profile balanced");

        _profilePowerItem = new NativeMenuItem("Profile: Power");
        _profilePowerItem.Click += async (_, _) => await RunQuickCommandAsync("profile power");

        _openQuickChatItem = new NativeMenuItem("Open Quick Chat");
        _openQuickChatItem.Click += (_, _) => ShowQuickChat();

        _openWebUiItem = new NativeMenuItem("Open Web UI");
        _openWebUiItem.Click += (_, _) => OpenWebUi();

        _checkUpdatesItem = new NativeMenuItem("Check updates now");
        _checkUpdatesItem.Click += async (_, _) => await CheckForUpdatesAsync(manual: true, _shutdown.Token);

        _applyUpdateItem = new NativeMenuItem("Apply downloaded update");
        _applyUpdateItem.Click += (_, _) => ApplyPendingUpdate();

        _exitItem = new NativeMenuItem("Exit");
        _exitItem.Click += (_, _) =>
        {
            if (ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
            {
                desktop.Shutdown();
            }
            else
            {
                DisposeResources();
                Environment.Exit(0);
            }
        };

        var menu = new NativeMenu();
        menu.Add(_statusItem);
        menu.Add(_updateStatusItem);
        menu.Add(new NativeMenuItemSeparator());
        menu.Add(_refreshItem);
        menu.Add(_checkUpdatesItem);
        menu.Add(_applyUpdateItem);
        menu.Add(new NativeMenuItemSeparator());
        menu.Add(_armItem);
        menu.Add(_disarmItem);
        menu.Add(_killItem);
        menu.Add(_resetKillItem);
        menu.Add(_lockWindowItem);
        menu.Add(_lockAppItem);
        menu.Add(_unlockItem);
        menu.Add(_profileSafeItem);
        menu.Add(_profileBalancedItem);
        menu.Add(_profilePowerItem);
        menu.Add(new NativeMenuItemSeparator());
        menu.Add(_openQuickChatItem);
        menu.Add(_openWebUiItem);
        menu.Add(new NativeMenuItemSeparator());
        menu.Add(_exitItem);

        _trayIcon = new TrayIcon
        {
            ToolTipText = "DesktopAgent",
            IsVisible = true,
            Menu = menu
        };

        TrySetTrayIconImage(_trayIcon);
        _trayIcon.Clicked += (_, _) => ShowQuickChat();

        _trayIcons = new TrayIcons { _trayIcon };
        TrayIcon.SetIcons(this, _trayIcons);
    }

    private async Task PollStatusAsync(CancellationToken cancellationToken)
    {
        var seconds = Math.Clamp(_settings.StatusRefreshSeconds, 1, 60);
        var interval = TimeSpan.FromSeconds(seconds);
        using var timer = new PeriodicTimer(interval);

        await RefreshStatusAsync();
        await CheckForUpdatesAsync(manual: false, cancellationToken);

        try
        {
            while (await timer.WaitForNextTickAsync(cancellationToken))
            {
                await RefreshStatusAsync();
                await CheckForUpdatesAsync(manual: false, cancellationToken);
            }
        }
        catch (OperationCanceledException)
        {
            // Normal shutdown.
        }
    }

    private async Task RefreshStatusAsync()
    {
        if (_client == null)
        {
            return;
        }

        var status = await _client.GetStatusAsync(_shutdown.Token);
        UpdateStatusUi(status);
    }

    private async Task RunQuickCommandAsync(string command)
    {
        if (_webApiClient == null)
        {
            return;
        }

        try
        {
            await _webApiClient.SendChatAsync(command, _shutdown.Token);
        }
        catch
        {
            // ignore
        }

        await RefreshStatusAsync();
    }

    private void UpdateStatusUi(Status status)
    {
        var armed = status.Armed;
        var requirePresence = status.RequireUserPresence;
        var detail = string.IsNullOrWhiteSpace(status.Message)
            ? "ok"
            : Compact(status.Message, 64);
        var armedIcon = armed ? "[+]" : "[-]";
        var presenceIcon = requirePresence ? "[+]" : "[-]";

        var line = $"Status: {armedIcon} Armed {(armed ? "ON" : "OFF")} | {presenceIcon} Presence {(requirePresence ? "REQ" : "OFF")} | {detail}";
        var tooltip = $"DesktopAgent: A={(armed ? "ON" : "OFF")} P={(requirePresence ? "REQ" : "OFF")}";

        Dispatcher.UIThread.Post(() =>
        {
            if (_statusItem != null)
            {
                _statusItem.Header = line;
            }

            if (_updateStatusItem != null)
            {
                _updateStatusItem.Header = $"Updates: {_updateStatus}";
            }

            if (_armItem != null)
            {
                _armItem.IsEnabled = !armed;
            }

            if (_disarmItem != null)
            {
                _disarmItem.IsEnabled = armed;
            }

            if (_checkUpdatesItem != null)
            {
                _checkUpdatesItem.IsEnabled = _updatesEnabled;
            }

            if (_applyUpdateItem != null)
            {
                _applyUpdateItem.IsEnabled = _pendingUpdate != null;
            }

            if (_trayIcon != null)
            {
                _trayIcon.ToolTipText = tooltip;
            }
        });
    }

    private void InitializeUpdater()
    {
        if (!_settings.AutoUpdateEnabled || string.IsNullOrWhiteSpace(_settings.AutoUpdateSource))
        {
            _updatesEnabled = false;
            _updateStatus = "disabled";
            return;
        }

        try
        {
            _updateManager = new UpdateManager(_settings.AutoUpdateSource);
            _updatesEnabled = true;
            _pendingUpdate = _updateManager.UpdatePendingRestart;
            _updateStatus = _pendingUpdate == null ? "enabled" : $"ready {_pendingUpdate.Version}";

            if (!_updateManager.IsInstalled)
            {
                _updateStatus = "dev mode";
            }
        }
        catch (Exception ex)
        {
            _updatesEnabled = false;
            _updateStatus = $"error: {Compact(ex.Message, 36)}";
        }
    }

    private async Task CheckForUpdatesAsync(bool manual, CancellationToken cancellationToken)
    {
        if (!_updatesEnabled || _updateManager == null)
        {
            return;
        }

        var intervalMinutes = Math.Clamp(_settings.AutoUpdateCheckIntervalMinutes, 5, 1440);
        if (!manual && DateTimeOffset.UtcNow - _lastUpdateCheck < TimeSpan.FromMinutes(intervalMinutes))
        {
            return;
        }

        if (!await _updateGate.WaitAsync(0, cancellationToken))
        {
            return;
        }

        try
        {
            _lastUpdateCheck = DateTimeOffset.UtcNow;
            _updateStatus = "checking...";

            if (_updateManager.UpdatePendingRestart != null)
            {
                _pendingUpdate = _updateManager.UpdatePendingRestart;
                _updateStatus = $"ready {_pendingUpdate.Version}";
                return;
            }

            var update = await _updateManager.CheckForUpdatesAsync();
            if (update == null)
            {
                _pendingUpdate = null;
                _updateStatus = "up to date";
                return;
            }

            _updateStatus = "downloading...";
            await _updateManager.DownloadUpdatesAsync(update);
            _pendingUpdate = update.TargetFullRelease;
            _updateStatus = $"ready {_pendingUpdate.Version}";

            if (_settings.AutoUpdateAutoApply && _pendingUpdate != null)
            {
                ApplyPendingUpdate();
            }
        }
        catch (Exception ex)
        {
            _updateStatus = $"failed: {Compact(ex.Message, 36)}";
        }
        finally
        {
            _updateGate.Release();
        }
    }

    private void ApplyPendingUpdate()
    {
        if (_updateManager == null || _pendingUpdate == null)
        {
            return;
        }

        try
        {
            _shutdown.Cancel();
            _updateManager.ApplyUpdatesAndRestart(_pendingUpdate);
        }
        catch (Exception ex)
        {
            _updateStatus = $"apply failed: {Compact(ex.Message, 30)}";
        }
    }

    private void OpenWebUi()
    {
        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = _settings.WebUiUrl,
                UseShellExecute = true
            });
        }
        catch
        {
            // Ignore shell launch errors.
        }
    }

    private void ShowQuickChat()
    {
        if (_webApiClient == null)
        {
            return;
        }

        Dispatcher.UIThread.Post(() =>
        {
            if (_quickChatWindow == null)
            {
                _quickChatWindow = new QuickChatWindow(_webApiClient, _settings.WebUiUrl);
                _quickChatWindow.Closed += (_, _) => _quickChatWindow = null;
            }

            if (!_quickChatWindow.IsVisible)
            {
                _quickChatWindow.Show();
            }

            _quickChatWindow.WindowState = WindowState.Normal;
            _quickChatWindow.Activate();
        });
    }

    private void TrySetTrayIconImage(TrayIcon trayIcon)
    {
        try
        {
            var uri = new Uri("avares://DesktopAgent.Tray/Assets/tray.png");
            if (!AssetLoader.Exists(uri))
            {
                return;
            }

            using var iconStream = AssetLoader.Open(uri);
            trayIcon.Icon = new WindowIcon(iconStream);
        }
        catch
        {
            // If icon load fails, the platform default may be used.
        }
    }

    private static TraySettings LoadSettings()
    {
        var settings = new TraySettings();
        var configuration = new ConfigurationBuilder()
            .SetBasePath(AppContext.BaseDirectory)
            .AddJsonFile("appsettings.json", optional: true)
            .AddEnvironmentVariables(prefix: "DESKTOP_AGENT_TRAY_")
            .Build();

        configuration.Bind(settings);

        if (string.IsNullOrWhiteSpace(settings.AdapterEndpoint))
        {
            settings.AdapterEndpoint = "http://localhost:51877";
        }

        if (string.IsNullOrWhiteSpace(settings.WebUiUrl))
        {
            settings.WebUiUrl = "http://localhost:51878";
        }

        settings.StatusRefreshSeconds = Math.Clamp(settings.StatusRefreshSeconds, 1, 60);
        settings.ApiTimeoutSeconds = Math.Clamp(settings.ApiTimeoutSeconds, 1, 3600);
        settings.AutoUpdateCheckIntervalMinutes = Math.Clamp(settings.AutoUpdateCheckIntervalMinutes, 5, 1440);
        return settings;
    }

    private static string Compact(string value, int max)
    {
        if (value.Length <= max)
        {
            return value;
        }

        return value[..max] + "...";
    }

    private void DisposeResources()
    {
        _shutdown.Cancel();
        _quickChatWindow?.Close();
        _updateGate.Dispose();
        _client?.Dispose();
        _webApiClient?.Dispose();
        _loggerFactory?.Dispose();
    }
}

internal sealed class TraySettings
{
    public string AdapterEndpoint { get; set; } = "http://localhost:51877";
    public string WebUiUrl { get; set; } = "http://localhost:51878";
    public bool RequireUserPresenceOnArm { get; set; } = true;
    public int StatusRefreshSeconds { get; set; } = 5;
    public int ApiTimeoutSeconds { get; set; } = 600;
    public bool AutoUpdateEnabled { get; set; } = true;
    public string AutoUpdateSource { get; set; } = "https://github.com/alessiobianchini/DesktopAgent/releases/latest/download";
    public int AutoUpdateCheckIntervalMinutes { get; set; } = 60;
    public bool AutoUpdateAutoApply { get; set; } = false;
}
