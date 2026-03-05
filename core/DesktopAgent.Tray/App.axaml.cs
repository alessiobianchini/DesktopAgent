using System.Diagnostics;
using System.Text.Json;
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
using System.Reflection;
using System.Text.RegularExpressions;
using Velopack;

namespace DesktopAgent.Tray;

public partial class App : Application
{
    private readonly CancellationTokenSource _shutdown = new();
    private TraySettings _settings = new();
    private Process? _managedAdapterProcess;
    private ILoggerFactory? _loggerFactory;
    private DesktopGrpcClient? _client;
    private WebApiClient? _webApiClient;
    private QuickChatWindow? _quickChatWindow;
    private readonly object _timelineSync = new();
    private List<string> _timelineSession = new();
    private PluginSetupState _pluginSetupState = new();
    private string _pluginSetupStatePath = string.Empty;

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
    private NativeMenuItem? _pluginsSetupItem;
    private NativeMenuItem? _openWebUiItem;
    private NativeMenuItem? _updateStatusItem;
    private NativeMenuItem? _checkUpdatesItem;
    private NativeMenuItem? _applyUpdateItem;
    private NativeMenuItem? _exitItem;

    private Task? _pollingTask;
    private UpdateManager? _updateManager;
    private VelopackAsset? _pendingUpdate;
    private bool _isVelopackInstalled;
    private IReadOnlyList<string> _lastUpdateBlockers = Array.Empty<string>();
    private readonly SemaphoreSlim _updateGate = new(1, 1);
    private DateTimeOffset _lastUpdateCheck = DateTimeOffset.MinValue;
    private string _updateStatus = "updates disabled";
    private bool _updatesEnabled;
    private TrayVisualState _trayVisualState = TrayVisualState.Unknown;

    public override void Initialize()
    {
        AvaloniaXamlLoader.Load(this);
    }

    public override void OnFrameworkInitializationCompleted()
    {
        _settings = LoadSettings();
        _pluginSetupStatePath = ResolvePluginSetupStatePath();
        _pluginSetupState = LoadPluginSetupState(_pluginSetupStatePath);
        EnsureBackendProcesses();
        _loggerFactory = LoggerFactory.Create(logging =>
        {
            logging.SetMinimumLevel(LogLevel.Information);
            logging.AddConsole();
        });

        _client = new DesktopGrpcClient(
            _settings.AdapterEndpoint,
            _loggerFactory.CreateLogger<DesktopGrpcClient>());
        _webApiClient = new WebApiClient(
            _settings.AdapterEndpoint,
            TimeSpan.FromSeconds(_settings.ApiTimeoutSeconds),
            _settings.AgentConfigPath);
        InitializeUpdater();

        ConfigureTray();

        if (ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
        {
            desktop.MainWindow = new HiddenWindow();
            desktop.Exit += (_, _) => DisposeResources();
        }

        _pollingTask = Task.Run(() => PollStatusAsync(_shutdown.Token));
        _ = Task.Run(() => RunPluginWizardIfNeededAsync(force: false, _shutdown.Token));

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

        _pluginsSetupItem = new NativeMenuItem("Install Optional Plugins...");
        _pluginsSetupItem.Click += async (_, _) => await RunPluginWizardIfNeededAsync(force: true, _shutdown.Token);

        _openWebUiItem = new NativeMenuItem("Open Data Folder");
        _openWebUiItem.Click += (_, _) => OpenWebUi();

        _checkUpdatesItem = new NativeMenuItem("Check updates now");
        _checkUpdatesItem.Click += async (_, _) => await CheckForUpdatesAsync(manual: true, _shutdown.Token);

        _applyUpdateItem = new NativeMenuItem("Apply downloaded update") { IsEnabled = false };
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
        menu.Add(_pluginsSetupItem);
        menu.Add(_openWebUiItem);
        menu.Add(new NativeMenuItemSeparator());
        menu.Add(_exitItem);

        _trayIcon = new TrayIcon
        {
            ToolTipText = "DesktopAgent",
            IsVisible = true,
            Menu = menu
        };

        UpdateTrayIconVisual(armed: false);
        _trayIcon.Clicked += (_, _) => ShowQuickChat();

        _trayIcons = new TrayIcons { _trayIcon };
        TrayIcon.SetIcons(this, _trayIcons);
        RefreshUpdateUi();
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

            RefreshUpdateUi();

            if (_trayIcon != null)
            {
                _trayIcon.ToolTipText = tooltip;
            }

            UpdateTrayIconVisual(armed);
        });
    }

    private void InitializeUpdater()
    {
        if (!_settings.AutoUpdateEnabled || string.IsNullOrWhiteSpace(_settings.AutoUpdateSource))
        {
            _updatesEnabled = false;
            _updateStatus = "disabled";
            _ = WriteUpdateAuditAsync("update_init", "Updater disabled", new
            {
                enabled = _settings.AutoUpdateEnabled,
                source = _settings.AutoUpdateSource
            });
            return;
        }

        try
        {
            _updateManager = new UpdateManager(_settings.AutoUpdateSource);
            _isVelopackInstalled = _updateManager.IsInstalled;
            _updatesEnabled = _isVelopackInstalled;
            _pendingUpdate = _updateManager.UpdatePendingRestart;
            if (!_isVelopackInstalled)
            {
                _updateStatus = "disabled (not installed via Velopack)";
                _pendingUpdate = null;
            }
            else
            {
                _updateStatus = _pendingUpdate == null ? "enabled" : $"ready {_pendingUpdate.Version}";
            }

            _ = WriteUpdateAuditAsync("update_init", "Updater initialized", new
            {
                enabled = _updatesEnabled,
                source = _settings.AutoUpdateSource,
                isInstalled = _updateManager.IsInstalled,
                pending = _pendingUpdate?.Version.ToString()
            });
        }
        catch (Exception ex)
        {
            _updatesEnabled = false;
            _updateStatus = ex.Message.Contains("VelopackLocator", StringComparison.OrdinalIgnoreCase)
                ? "disabled (non-velopack install)"
                : $"error: {Compact(ex.Message, 36)}";
            _ = WriteUpdateAuditAsync("update_init_failed", "Updater initialization failed", new
            {
                source = _settings.AutoUpdateSource,
                error = ex.Message
            });
        }

        RefreshUpdateUi();
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
            await WriteUpdateAuditAsync("update_check_started", "Checking for updates", new
            {
                manual,
                source = _settings.AutoUpdateSource
            });

            _lastUpdateCheck = DateTimeOffset.UtcNow;
            _updateStatus = "checking...";

            if (_updateManager.UpdatePendingRestart != null)
            {
                _pendingUpdate = _updateManager.UpdatePendingRestart;
                _updateStatus = $"ready {_pendingUpdate.Version}";
                await WriteUpdateAuditAsync("update_pending", "Update already pending restart", new
                {
                    version = _pendingUpdate.Version.ToString(),
                    manual
                });
                return;
            }

            var update = await _updateManager.CheckForUpdatesAsync();
            if (update == null)
            {
                _pendingUpdate = null;
                _updateStatus = "up to date";
                await WriteUpdateAuditAsync("update_up_to_date", "No updates available", new
                {
                    manual
                });
                return;
            }

            _updateStatus = "downloading...";
            await _updateManager.DownloadUpdatesAsync(update);
            _pendingUpdate = update.TargetFullRelease;
            _updateStatus = $"ready {_pendingUpdate.Version}";
            await WriteUpdateAuditAsync("update_downloaded", "Update downloaded", new
            {
                manual,
                version = _pendingUpdate.Version.ToString(),
                file = _pendingUpdate.FileName
            });

            if (_settings.AutoUpdateAutoApply && _pendingUpdate != null)
            {
                ApplyPendingUpdate();
            }
        }
        catch (Exception ex)
        {
            _updateStatus = $"failed: {Compact(ex.Message, 36)}";
            await WriteUpdateAuditAsync("update_check_failed", "Update check failed", new
            {
                manual,
                source = _settings.AutoUpdateSource,
                error = ex.Message
            });
        }
        finally
        {
            _updateGate.Release();
            RefreshUpdateUi();
        }
    }

    private void ApplyPendingUpdate()
    {
        _lastUpdateBlockers = Array.Empty<string>();
        if (_updateManager == null || _pendingUpdate == null)
        {
            _ = WriteUpdateAuditAsync("update_apply_skipped", "No pending update to apply", null);
            return;
        }

        if (!_isVelopackInstalled)
        {
            _updateStatus = "apply blocked: not a Velopack-installed app";
            _ = WriteUpdateAuditAsync("update_apply_skipped", "Apply skipped because current app is not Velopack-installed", new
            {
                currentExe = Environment.ProcessPath
            });
            RefreshUpdateUi();
            return;
        }

        try
        {
            var installRoot = TryResolveInstallRoot();
            var killed = installRoot == null ? Array.Empty<string>() : TryTerminateProcessesInInstallRoot(installRoot);
            _ = WriteUpdateAuditAsync("update_apply_requested", "Applying downloaded update", new
            {
                version = _pendingUpdate.Version.ToString(),
                file = _pendingUpdate.FileName,
                currentExe = Environment.ProcessPath,
                adapterPid = _managedAdapterProcess?.Id,
                installRoot,
                killed
            });

            // Ensure sidecar processes don't keep files locked in the current app folder.
            TryStopManagedProcess(_managedAdapterProcess);
            _managedAdapterProcess = null;

            _shutdown.Cancel();
            Thread.Sleep(900);

            if (!string.IsNullOrWhiteSpace(installRoot))
            {
                var blockers = WaitForInstallRootToUnlock(installRoot!, timeoutMs: 5000);
                if (blockers.Count > 0)
                {
                    _lastUpdateBlockers = blockers.ToList();
                    _updateStatus = $"apply blocked: files in use ({blockers.Count})";
                    _ = WriteUpdateAuditAsync("update_apply_blocked", "Apply blocked by running processes in install root", new
                    {
                        installRoot,
                        blockers
                    });
                    RefreshUpdateUi();
                    return;
                }
            }

            TryApplyUpdateWithRetry(_updateManager, _pendingUpdate);
            _lastUpdateBlockers = Array.Empty<string>();
        }
        catch (Exception ex)
        {
            _updateStatus = $"apply failed: {Compact(ex.Message, 30)}";
            _ = WriteUpdateAuditAsync("update_apply_failed", "Apply update failed", new
            {
                version = _pendingUpdate.Version.ToString(),
                error = ex.Message
            });
            RefreshUpdateUi();
        }
    }

    private void TryApplyUpdateWithRetry(UpdateManager updateManager, VelopackAsset pendingUpdate)
    {
        try
        {
            updateManager.ApplyUpdatesAndRestart(pendingUpdate);
        }
        catch
        {
            var installRoot = TryResolveInstallRoot();
            if (!string.IsNullOrWhiteSpace(installRoot))
            {
                TryTerminateProcessesInInstallRoot(installRoot!);
                Thread.Sleep(800);
            }

            updateManager.ApplyUpdatesAndRestart(pendingUpdate);
        }
    }

    private void RefreshUpdateUi()
    {
        Dispatcher.UIThread.Post(() =>
        {
            if (_updateStatusItem != null)
            {
                _updateStatusItem.Header = $"Updates: {_updateStatus}";
            }

            if (_checkUpdatesItem != null)
            {
                _checkUpdatesItem.IsEnabled = _updatesEnabled;
            }

            if (_applyUpdateItem != null)
            {
                _applyUpdateItem.IsEnabled = CanApplyPendingUpdate();
            }
        });
    }

    private bool CanApplyPendingUpdate()
    {
        if (_pendingUpdate == null)
        {
            return false;
        }

        var current = ReadCurrentVersionString();
        if (string.IsNullOrWhiteSpace(current))
        {
            return true;
        }

        return IsVersionGreater(_pendingUpdate.Version.ToString(), current);
    }

    private static string ReadCurrentVersionString()
    {
        try
        {
            var exeDir = AppContext.BaseDirectory;
            var sqVersionPath = Path.Combine(exeDir, "sq.version");
            if (File.Exists(sqVersionPath))
            {
                var text = File.ReadAllText(sqVersionPath);
                var match = Regex.Match(text, "<version>([^<]+)</version>", RegexOptions.IgnoreCase);
                if (match.Success)
                {
                    return match.Groups[1].Value.Trim();
                }
            }
        }
        catch
        {
            // ignore
        }

        try
        {
            var asm = Assembly.GetExecutingAssembly();
            var info = asm.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion;
            if (!string.IsNullOrWhiteSpace(info))
            {
                return info.Trim();
            }
        }
        catch
        {
            // ignore
        }

        return string.Empty;
    }

    private static bool IsVersionGreater(string left, string right)
    {
        var l = NormalizeSemVerCore(left);
        var r = NormalizeSemVerCore(right);
        return CompareSemVerCore(l, r) > 0;
    }

    private static int CompareSemVerCore((int Major, int Minor, int Patch) a, (int Major, int Minor, int Patch) b)
    {
        var c = a.Major.CompareTo(b.Major);
        if (c != 0) return c;
        c = a.Minor.CompareTo(b.Minor);
        if (c != 0) return c;
        return a.Patch.CompareTo(b.Patch);
    }

    private static (int Major, int Minor, int Patch) NormalizeSemVerCore(string value)
    {
        var raw = (value ?? string.Empty).Trim();
        if (raw.StartsWith('v') || raw.StartsWith('V'))
        {
            raw = raw[1..];
        }

        var core = raw.Split('-', 2)[0];
        var parts = core.Split('.', StringSplitOptions.RemoveEmptyEntries);
        var major = parts.Length > 0 && int.TryParse(parts[0], out var m) ? m : 0;
        var minor = parts.Length > 1 && int.TryParse(parts[1], out var n) ? n : 0;
        var patch = parts.Length > 2 && int.TryParse(parts[2], out var p) ? p : 0;
        return (major, minor, patch);
    }

    private async Task WriteUpdateAuditAsync(string eventType, string message, object? data)
    {
        if (_webApiClient == null)
        {
            return;
        }

        try
        {
            await _webApiClient.WriteSystemAuditAsync(eventType, message, data, CancellationToken.None);
        }
        catch
        {
            // Best effort: never fail tray update flow for logging issues.
        }
    }

    private void OpenWebUi()
    {
        try
        {
            var dataRoot = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "DesktopAgent");
            Directory.CreateDirectory(dataRoot);
            Process.Start(new ProcessStartInfo
            {
                FileName = dataRoot,
                UseShellExecute = true
            });
        }
        catch
        {
            // Ignore shell launch errors.
        }
    }

    private async Task RunPluginWizardIfNeededAsync(bool force, CancellationToken cancellationToken)
    {
        if (!OperatingSystem.IsWindows())
        {
            return;
        }

        if (!_settings.ShowPluginWizardOnFirstRun && !force)
        {
            return;
        }

        if (!force && _pluginSetupState.Dismissed)
        {
            return;
        }

        while (true)
        {
            var probeBefore = ProbePluginEnvironment();
            if (!force && !probeBefore.HasMissingTools)
            {
                _pluginSetupState.Dismissed = true;
                SavePluginSetupState(_pluginSetupStatePath, _pluginSetupState);
                return;
            }

            PluginSetupChoice? choice = null;
            await Dispatcher.UIThread.InvokeAsync(async () =>
            {
                var window = new PluginSetupWindow(probeBefore);
                choice = await ShowPluginSetupWindowAsync(window);
            });

            if (choice == null)
            {
                return;
            }

            if (choice.DontAskAgain || choice.InstallRequested)
            {
                _pluginSetupState.Dismissed = true;
                SavePluginSetupState(_pluginSetupStatePath, _pluginSetupState);
            }

            if (!choice.InstallRequested)
            {
                return;
            }

            var ffmpegResult = WingetInstallResult.NotRequested("FFmpeg not selected.");
            var ocrResult = WingetInstallResult.NotRequested("OCR not selected.");

            if (choice.InstallFfmpeg && !probeBefore.FfmpegInstalled && probeBefore.WingetAvailable)
            {
                ffmpegResult = await InstallWingetPackageAsync("Gyan.FFmpeg", cancellationToken);
            }

            if (choice.InstallOcr && !probeBefore.TesseractInstalled && probeBefore.WingetAvailable)
            {
                WingetInstallResult? primaryAttempt = null;
                if (probeBefore.TesseractPrimaryPackageAvailable)
                {
                    primaryAttempt = await InstallWingetPackageAsync("UB-Mannheim.TesseractOCR", cancellationToken);
                    ocrResult = primaryAttempt;
                }

                if ((primaryAttempt == null || !primaryAttempt.Success) && probeBefore.TesseractFallbackPackageAvailable)
                {
                    var fallbackAttempt = await InstallWingetPackageAsync("tesseract-ocr.tesseract", cancellationToken);
                    ocrResult = fallbackAttempt.Success
                        ? fallbackAttempt
                        : WingetInstallResult.Failed(
                            primaryAttempt == null
                                ? fallbackAttempt.Message
                                : $"{primaryAttempt.Message}; fallback: {fallbackAttempt.Message}");
                }

                if (primaryAttempt == null && !probeBefore.TesseractFallbackPackageAvailable)
                {
                    ocrResult = WingetInstallResult.Failed("No OCR package available in winget.");
                }
            }

            var probeAfter = ProbePluginEnvironment();
            if (choice.InstallOcr && probeAfter.TesseractInstalled)
            {
                await EnableOcrIfDisabledAsync(cancellationToken);
            }

            var hasFailures = (choice.InstallFfmpeg && !probeAfter.FfmpegInstalled)
                || (choice.InstallOcr && !probeAfter.TesseractInstalled);

            var retry = false;
            await Dispatcher.UIThread.InvokeAsync(async () =>
            {
                var summary = BuildPluginInstallSummary(choice, probeBefore, probeAfter, ffmpegResult, ocrResult);
                retry = await PluginSetupWindow.ShowResultAsync(summary, ResolveDialogOwner(), hasFailures);
            });

            if (!retry)
            {
                return;
            }

            force = true;
        }
    }

    private async Task<PluginSetupChoice?> ShowPluginSetupWindowAsync(PluginSetupWindow window)
    {
        var owner = ResolveDialogOwner();
        if (owner != null && owner.IsVisible)
        {
            await window.ShowDialog(owner);
            return window.ResultChoice;
        }

        var tcs = new TaskCompletionSource<object?>();
        void ClosedHandler(object? sender, EventArgs args) => tcs.TrySetResult(null);
        window.Closed += ClosedHandler;
        window.Show();
        await tcs.Task;
        window.Closed -= ClosedHandler;
        return window.ResultChoice;
    }

    private async Task<WingetInstallResult> InstallWingetPackageAsync(string packageId, CancellationToken cancellationToken)
    {
        try
        {
            using var process = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = "cmd",
                    Arguments = $"/C winget install -e --id {packageId} --accept-package-agreements --accept-source-agreements --disable-interactivity",
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    WindowStyle = ProcessWindowStyle.Hidden
                }
            };

            process.Start();
            var timeout = TimeSpan.FromSeconds(Math.Clamp(_settings.PluginInstallTimeoutSeconds, 30, 1800));
            using var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            var waitTask = process.WaitForExitAsync(timeoutCts.Token);
            var delayTask = Task.Delay(timeout, cancellationToken);
            var completed = await Task.WhenAny(waitTask, delayTask);

            if (completed == delayTask)
            {
                TryStopManagedProcess(process);
                return WingetInstallResult.Timeout(packageId, timeout);
            }

            await waitTask;
            return process.ExitCode == 0
                ? WingetInstallResult.Succeeded(packageId)
                : WingetInstallResult.Failed(FormatWingetExitMessage(process.ExitCode, packageId));
        }
        catch (OperationCanceledException)
        {
            return WingetInstallResult.Failed($"Installation canceled for {packageId}");
        }
        catch
        {
            return WingetInstallResult.Failed($"Failed to install {packageId}");
        }
    }

    private static string FormatWingetExitMessage(int exitCode, string packageId)
    {
        var hex = $"0x{unchecked((uint)exitCode):X8}";
        if (unchecked((uint)exitCode) == 0x8A15002B)
        {
            return $"winget exit code {exitCode} ({hex}) for {packageId}. Source/package metadata issue. Run 'winget source reset --force' and retry.";
        }

        return $"winget exit code {exitCode} ({hex}) for {packageId}";
    }

    private async Task EnableOcrIfDisabledAsync(CancellationToken cancellationToken)
    {
        if (_webApiClient == null)
        {
            return;
        }

        try
        {
            var config = await _webApiClient.GetConfigAsync(cancellationToken);
            if (config?.OcrEnabled == false)
            {
                await _webApiClient.SaveConfigAsync(new WebConfigUpdate(
                    Llm: null,
                    ProfileModeEnabled: null,
                    ActiveProfile: null,
                    RequireConfirmation: null,
                    MaxActionsPerSecond: null,
                    QuizSafeModeEnabled: null,
                    OcrEnabled: true,
                    ScreenRecordingAudioBackendPreference: null,
                    ScreenRecordingAudioDevice: null,
                    ScreenRecordingPrimaryDisplayOnly: null,
                    AdapterRestartCommand: null,
                    AdapterRestartWorkingDir: null,
                    FindRetryCount: null,
                    FindRetryDelayMs: null,
                    PostCheckTimeoutMs: null,
                    PostCheckPollMs: null,
                    ClipboardHistoryMaxItems: null,
                    FilesystemAllowedRoots: null,
                    AllowedApps: null,
                    AppAliases: null,
                    AuditLlmInteractions: null,
                    AuditLlmIncludeRawText: null), cancellationToken);
            }
        }
        catch
        {
            // Best effort.
        }
    }

    private static PluginSetupProbe ProbePluginEnvironment()
    {
        var winget = CommandExists("winget");
        var ffmpeg = CommandExists("ffmpeg");
        var tesseract = CommandExists("tesseract");

        var primary = false;
        var fallback = false;
        if (winget && !tesseract)
        {
            primary = WingetPackageExists("UB-Mannheim.TesseractOCR");
            if (!primary)
            {
                fallback = WingetPackageExists("tesseract-ocr.tesseract");
            }
        }

        return new PluginSetupProbe(
            winget,
            ffmpeg,
            tesseract,
            primary,
            fallback);
    }

    private static bool CommandExists(string command)
    {
        try
        {
            using var process = Process.Start(new ProcessStartInfo
            {
                FileName = "cmd",
                Arguments = $"/C where {command}",
                UseShellExecute = false,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden
            });
            if (process == null)
            {
                return false;
            }

            process.WaitForExit(3000);
            return process.HasExited && process.ExitCode == 0;
        }
        catch
        {
            return false;
        }
    }

    private static bool WingetPackageExists(string packageId)
    {
        try
        {
            using var process = Process.Start(new ProcessStartInfo
            {
                FileName = "cmd",
                Arguments = $"/C winget show -e --id {packageId}",
                UseShellExecute = false,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden
            });
            if (process == null)
            {
                return false;
            }

            process.WaitForExit(8000);
            return process.HasExited && process.ExitCode == 0;
        }
        catch
        {
            return false;
        }
    }

    private static string BuildPluginInstallSummary(
        PluginSetupChoice choice,
        PluginSetupProbe before,
        PluginSetupProbe after,
        WingetInstallResult ffmpegResult,
        WingetInstallResult ocrResult)
    {
        var lines = new List<string>();

        lines.Add("Environment check:");
        lines.Add($"- winget: {(before.WingetAvailable ? "available" : "missing")}");
        lines.Add($"- FFmpeg: {(before.FfmpegInstalled ? "already installed" : "missing")} -> {(after.FfmpegInstalled ? "installed" : "missing")}");
        lines.Add($"- OCR/Tesseract: {(before.TesseractInstalled ? "already installed" : "missing")} -> {(after.TesseractInstalled ? "installed" : "missing")}");

        if (choice.InstallFfmpeg)
        {
            lines.Add($"FFmpeg install: {ffmpegResult.Message}");
        }

        if (choice.InstallOcr)
        {
            lines.Add($"OCR install: {ocrResult.Message}");
        }

        if (lines.Count == 0)
        {
            lines.Add("No plugin selected.");
        }

        if ((choice.InstallFfmpeg && !after.FfmpegInstalled) || (choice.InstallOcr && !after.TesseractInstalled))
        {
            lines.Add("Some tools are still missing. Click Retry to try again.");
        }

        return string.Join(Environment.NewLine, lines);
    }

    private Window? ResolveDialogOwner()
    {
        if (_quickChatWindow != null && _quickChatWindow.IsVisible)
        {
            return _quickChatWindow;
        }

        if (Current?.ApplicationLifetime is IClassicDesktopStyleApplicationLifetime desktop)
        {
            if (desktop.MainWindow != null && desktop.MainWindow.IsVisible)
            {
                return desktop.MainWindow;
            }
        }

        return null;
    }

    private static string ResolvePluginSetupStatePath()
    {
        var root = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "DesktopAgent");
        Directory.CreateDirectory(root);
        return Path.Combine(root, "tray-plugin-state.json");
    }

    private static PluginSetupState LoadPluginSetupState(string path)
    {
        try
        {
            if (!File.Exists(path))
            {
                return new PluginSetupState();
            }

            var json = File.ReadAllText(path);
            var state = JsonSerializer.Deserialize<PluginSetupState>(json);
            return state ?? new PluginSetupState();
        }
        catch
        {
            return new PluginSetupState();
        }
    }

    private static void SavePluginSetupState(string path, PluginSetupState state)
    {
        try
        {
            var json = JsonSerializer.Serialize(state, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(path, json);
        }
        catch
        {
            // Best effort.
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
                _quickChatWindow = new QuickChatWindow(
                    _webApiClient,
                    () => RunPluginWizardIfNeededAsync(force: true, _shutdown.Token),
                    () => CheckForUpdatesAsync(manual: true, _shutdown.Token),
                    ApplyPendingUpdate,
                    GetUpdateStatusLine,
                    GetChatUpdateBadge,
                    GetChatUpdateDetails,
                    LoadTimelineSession,
                    SaveTimelineSession);
                _quickChatWindow.Closed += (_, _) => _quickChatWindow = null;
            }

            if (!_quickChatWindow.IsVisible)
            {
                _quickChatWindow.Show();
            }

            _quickChatWindow.WindowState = WindowState.Maximized;
            _quickChatWindow.Activate();
        });
    }

    private void UpdateTrayIconVisual(bool armed)
    {
        if (_trayIcon == null)
        {
            return;
        }

        var next = armed ? TrayVisualState.Armed : TrayVisualState.Disarmed;
        if (next == _trayVisualState)
        {
            return;
        }

        if (TrySetTrayIconImage(_trayIcon, next))
        {
            _trayVisualState = next;
        }
    }

    private bool TrySetTrayIconImage(TrayIcon trayIcon, TrayVisualState state)
    {
        try
        {
            var iconFile = state switch
            {
                TrayVisualState.Armed => "tray-armed.png",
                TrayVisualState.Disarmed => "tray-disarmed.png",
                _ => "tray.png"
            };

            var uri = new Uri($"avares://DesktopAgent.Tray/Assets/{iconFile}");
            if (!AssetLoader.Exists(uri))
            {
                uri = new Uri("avares://DesktopAgent.Tray/Assets/tray.png");
                if (!AssetLoader.Exists(uri))
                {
                    return false;
                }
            }

            using var iconStream = AssetLoader.Open(uri);
            trayIcon.Icon = new WindowIcon(iconStream);
            return true;
        }
        catch
        {
            // If icon load fails, the platform default may be used.
            return false;
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

        if (string.IsNullOrWhiteSpace(settings.AgentConfigPath))
        {
            settings.AgentConfigPath = "agentsettings.json";
        }

        settings.StatusRefreshSeconds = Math.Clamp(settings.StatusRefreshSeconds, 1, 60);
        settings.ApiTimeoutSeconds = Math.Clamp(settings.ApiTimeoutSeconds, 1, 3600);
        settings.AutoUpdateCheckIntervalMinutes = Math.Clamp(settings.AutoUpdateCheckIntervalMinutes, 5, 1440);
        if (!string.IsNullOrWhiteSpace(settings.AutoUpdateSource) && !settings.AutoUpdateSource.EndsWith("/", StringComparison.Ordinal))
        {
            settings.AutoUpdateSource += "/";
        }
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

    private string GetUpdateStatusLine()
    {
        var pending = _pendingUpdate?.Version.ToString() ?? "none";
        var mode = _isVelopackInstalled ? "velopack" : "portable";
        var blockers = _lastUpdateBlockers.Count > 0 ? $" | Blockers: {_lastUpdateBlockers.Count}" : string.Empty;
        return $"Updates: {_updateStatus} | Pending: {pending} | Mode: {mode}{blockers}";
    }

    private ChatUpdateBadge GetChatUpdateBadge()
    {
        if (_pendingUpdate != null)
        {
            var canApply = CanApplyPendingUpdate();
            return canApply
                ? new ChatUpdateBadge(true, true, $"Update available: v{_pendingUpdate.Version}")
                : new ChatUpdateBadge(false, false, string.Empty);
        }

        if (_updatesEnabled && _updateStatus.StartsWith("downloading", StringComparison.OrdinalIgnoreCase))
        {
            return new ChatUpdateBadge(true, false, "Downloading update...");
        }

        if (_updatesEnabled && _updateStatus.StartsWith("checking", StringComparison.OrdinalIgnoreCase))
        {
            return new ChatUpdateBadge(true, false, "Checking for updates...");
        }

        return new ChatUpdateBadge(false, false, string.Empty);
    }

    private ChatUpdateDetails GetChatUpdateDetails()
    {
        if (_pendingUpdate != null)
        {
            var notes = TryReadStringProperty(_pendingUpdate, "ReleaseNotes", "Notes", "Description", "Summary");
            var details = string.IsNullOrWhiteSpace(notes)
                ? $"A downloaded update is ready to apply.{Environment.NewLine}Version: v{_pendingUpdate.Version}{Environment.NewLine}Source: {_settings.AutoUpdateSource}"
                : $"Version: v{_pendingUpdate.Version}{Environment.NewLine}{Environment.NewLine}{notes}";

            return new ChatUpdateDetails(
                HasUpdate: true,
                Title: $"Update v{_pendingUpdate.Version}",
                Details: details,
                CanApply: true,
                HasBlockers: _lastUpdateBlockers.Count > 0,
                Blockers: _lastUpdateBlockers.Count == 0 ? string.Empty : string.Join(Environment.NewLine, _lastUpdateBlockers));
        }

        return new ChatUpdateDetails(
            HasUpdate: false,
            Title: "Updates",
            Details: $"{_updateStatus}{Environment.NewLine}Source: {_settings.AutoUpdateSource}{Environment.NewLine}Mode: {(_isVelopackInstalled ? "Velopack installed" : "Portable/non-Velopack")}{Environment.NewLine}Exe: {Environment.ProcessPath}",
            CanApply: false,
            HasBlockers: _lastUpdateBlockers.Count > 0,
            Blockers: _lastUpdateBlockers.Count == 0 ? string.Empty : string.Join(Environment.NewLine, _lastUpdateBlockers));
    }

    private static string? TryReadStringProperty(object source, params string[] names)
    {
        var type = source.GetType();
        foreach (var name in names)
        {
            var property = type.GetProperty(name);
            var value = property?.GetValue(source) as string;
            if (!string.IsNullOrWhiteSpace(value))
            {
                return value.Trim();
            }
        }

        return null;
    }

    private IReadOnlyList<string> LoadTimelineSession()
    {
        lock (_timelineSync)
        {
            return _timelineSession.ToList();
        }
    }

    private void SaveTimelineSession(IReadOnlyList<string> lines)
    {
        lock (_timelineSync)
        {
            _timelineSession = lines
                .Where(line => !string.IsNullOrWhiteSpace(line))
                .TakeLast(140)
                .ToList();
        }
    }

    private void EnsureBackendProcesses()
    {
        var startedAny = false;
        if (_settings.AutoStartAdapter && !IsTcpPortOpen(_settings.AdapterEndpoint))
        {
            var command = string.IsNullOrWhiteSpace(_settings.AdapterStartCommand)
                ? BuildDefaultAdapterStartCommand()
                : _settings.AdapterStartCommand;
            _managedAdapterProcess = TryStartManagedProcess(command, ("DESKTOP_AGENT_PORT", ExtractPort(_settings.AdapterEndpoint)?.ToString()));
            startedAny = _managedAdapterProcess != null;
        }

        // Tray-only mode: web backend is optional and disabled by default.
        // If users still want the web host, they can launch it manually.

        if (startedAny)
        {
            Thread.Sleep(800);
        }
    }

    private static Process? TryStartManagedProcess(string command, params (string Key, string? Value)[] env)
    {
        if (!TryParseCommand(command, out var fileName, out var args))
        {
            return null;
        }

        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = fileName,
                Arguments = args,
                UseShellExecute = false,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden
            };

            foreach (var (key, value) in env)
            {
                if (!string.IsNullOrWhiteSpace(key) && !string.IsNullOrWhiteSpace(value))
                {
                    psi.Environment[key] = value;
                }
            }

            return Process.Start(psi);
        }
        catch
        {
            return null;
        }
    }

    private static void TryStopManagedProcess(Process? process)
    {
        if (process == null)
        {
            return;
        }

        try
        {
            if (!process.HasExited)
            {
                process.Kill(true);
            }
        }
        catch
        {
            // ignored
        }
    }

    private static string? TryResolveInstallRoot()
    {
        try
        {
            var exe = Environment.ProcessPath;
            if (string.IsNullOrWhiteSpace(exe))
            {
                return null;
            }

            var currentDir = Path.GetDirectoryName(exe);
            if (string.IsNullOrWhiteSpace(currentDir))
            {
                return null;
            }

            return Directory.GetParent(currentDir)?.FullName;
        }
        catch
        {
            return null;
        }
    }

    private static IReadOnlyList<string> TryTerminateProcessesInInstallRoot(string installRoot)
    {
        var killed = new List<string>();
        var currentPid = Environment.ProcessId;
        foreach (var process in Process.GetProcesses())
        {
            try
            {
                if (process.Id == currentPid)
                {
                    continue;
                }

                string? path;
                try
                {
                    path = process.MainModule?.FileName;
                }
                catch
                {
                    continue;
                }

                if (string.IsNullOrWhiteSpace(path))
                {
                    continue;
                }

                if (!path.StartsWith(installRoot, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (string.Equals(Path.GetFileName(path), "Update.exe", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                process.Kill(entireProcessTree: true);
                process.WaitForExit(4000);
                killed.Add($"{process.ProcessName}:{process.Id}");
            }
            catch
            {
                // best effort
            }
            finally
            {
                process.Dispose();
            }
        }

        return killed;
    }

    private static IReadOnlyList<string> WaitForInstallRootToUnlock(string installRoot, int timeoutMs)
    {
        var deadline = DateTimeOffset.UtcNow.AddMilliseconds(Math.Max(500, timeoutMs));
        IReadOnlyList<string> blockers = Array.Empty<string>();
        while (DateTimeOffset.UtcNow < deadline)
        {
            blockers = GetProcessesInInstallRoot(installRoot);
            if (blockers.Count == 0)
            {
                return blockers;
            }

            Thread.Sleep(250);
        }

        return blockers;
    }

    private static IReadOnlyList<string> GetProcessesInInstallRoot(string installRoot)
    {
        var found = new List<string>();
        var currentPid = Environment.ProcessId;
        foreach (var process in Process.GetProcesses())
        {
            try
            {
                if (process.Id == currentPid)
                {
                    continue;
                }

                var path = process.MainModule?.FileName;
                if (string.IsNullOrWhiteSpace(path))
                {
                    continue;
                }

                if (!path.StartsWith(installRoot, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (string.Equals(Path.GetFileName(path), "Update.exe", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                found.Add($"{process.ProcessName}:{process.Id}");
            }
            catch
            {
                // best effort
            }
            finally
            {
                process.Dispose();
            }
        }

        return found;
    }

    private static bool IsTcpPortOpen(string endpoint)
    {
        if (!Uri.TryCreate(endpoint, UriKind.Absolute, out var uri))
        {
            return false;
        }

        var port = ExtractPort(uri) ?? (uri.Scheme == "https" ? 443 : 80);
        var host = string.IsNullOrWhiteSpace(uri.Host) ? "127.0.0.1" : uri.Host;

        try
        {
            using var client = new System.Net.Sockets.TcpClient();
            var connectTask = client.ConnectAsync(host, port);
            var ok = connectTask.Wait(TimeSpan.FromMilliseconds(300));
            return ok && client.Connected;
        }
        catch
        {
            return false;
        }
    }

    private static int? ExtractPort(string endpoint)
    {
        if (!Uri.TryCreate(endpoint, UriKind.Absolute, out var uri))
        {
            return null;
        }

        return ExtractPort(uri);
    }

    private static int? ExtractPort(Uri uri)
    {
        return uri.IsDefaultPort ? null : uri.Port;
    }

    private static string BuildDefaultAdapterStartCommand()
    {
        var baseDir = AppContext.BaseDirectory;
        if (OperatingSystem.IsWindows())
        {
            var path = ResolveExistingPath(
                Path.GetFullPath(Path.Combine(baseDir, "adapter", "DesktopAgent.Adapter.Windows.exe")),
                Path.GetFullPath(Path.Combine(baseDir, "..", "adapter", "DesktopAgent.Adapter.Windows.exe")));
            return $"\"{path}\"";
        }

        var dllPath = ResolveExistingPath(
            Path.GetFullPath(Path.Combine(baseDir, "adapter", "DesktopAgent.Adapter.Windows.dll")),
            Path.GetFullPath(Path.Combine(baseDir, "..", "adapter", "DesktopAgent.Adapter.Windows.dll")));
        return $"dotnet \"{dllPath}\"";
    }

    private static string ResolveExistingPath(params string[] candidates)
    {
        foreach (var candidate in candidates)
        {
            if (File.Exists(candidate))
            {
                return candidate;
            }
        }

        return candidates.Length > 0 ? candidates[0] : string.Empty;
    }

    private static bool TryParseCommand(string commandLine, out string fileName, out string args)
    {
        fileName = string.Empty;
        args = string.Empty;
        if (string.IsNullOrWhiteSpace(commandLine))
        {
            return false;
        }

        var tokens = SplitCommandLine(commandLine);
        if (tokens.Count == 0)
        {
            return false;
        }

        fileName = tokens[0];
        if (tokens.Count > 1)
        {
            args = string.Join(' ', tokens.Skip(1));
        }

        return true;
    }

    private static List<string> SplitCommandLine(string commandLine)
    {
        var results = new List<string>();
        var current = new System.Text.StringBuilder();
        var inQuotes = false;
        char quoteChar = '\0';

        for (var i = 0; i < commandLine.Length; i++)
        {
            var ch = commandLine[i];
            if (inQuotes)
            {
                if (ch == quoteChar)
                {
                    inQuotes = false;
                }
                else
                {
                    current.Append(ch);
                }
                continue;
            }

            if (ch == '"' || ch == '\'')
            {
                inQuotes = true;
                quoteChar = ch;
                continue;
            }

            if (char.IsWhiteSpace(ch))
            {
                if (current.Length > 0)
                {
                    results.Add(current.ToString());
                    current.Clear();
                }
                continue;
            }

            current.Append(ch);
        }

        if (current.Length > 0)
        {
            results.Add(current.ToString());
        }

        return results;
    }

    private void DisposeResources()
    {
        _shutdown.Cancel();
        _quickChatWindow?.Close();
        _updateGate.Dispose();
        _client?.Dispose();
        _webApiClient?.Dispose();
        _loggerFactory?.Dispose();
        TryStopManagedProcess(_managedAdapterProcess);
    }
}

internal enum TrayVisualState
{
    Unknown = 0,
    Disarmed = 1,
    Armed = 2
}

internal sealed class TraySettings
{
    public string AdapterEndpoint { get; set; } = "http://localhost:51877";
    public string AgentConfigPath { get; set; } = "agentsettings.json";
    public bool RequireUserPresenceOnArm { get; set; } = true;
    public int StatusRefreshSeconds { get; set; } = 5;
    public int ApiTimeoutSeconds { get; set; } = 600;
    public bool AutoUpdateEnabled { get; set; } = true;
    public string AutoUpdateSource { get; set; } = "https://github.com/alessiobianchini/DesktopAgent/releases/latest/download/";
    public int AutoUpdateCheckIntervalMinutes { get; set; } = 60;
    public bool AutoUpdateAutoApply { get; set; } = false;
    public bool AutoStartAdapter { get; set; } = true;
    public string AdapterStartCommand { get; set; } = "";
    public bool ShowPluginWizardOnFirstRun { get; set; } = true;
    public int PluginInstallTimeoutSeconds { get; set; } = 420;
}

internal sealed class PluginSetupState
{
    public bool Dismissed { get; set; }
}

internal sealed record WingetInstallResult(bool Success, bool TimedOut, string Message)
{
    public static WingetInstallResult NotRequested(string message) => new(false, false, message);
    public static WingetInstallResult Succeeded(string packageId) => new(true, false, $"Installed {packageId} successfully.");
    public static WingetInstallResult Failed(string message) => new(false, false, message);
    public static WingetInstallResult Timeout(string packageId, TimeSpan timeout) => new(false, true, $"Timed out after {(int)timeout.TotalSeconds}s while installing {packageId}.");
}
