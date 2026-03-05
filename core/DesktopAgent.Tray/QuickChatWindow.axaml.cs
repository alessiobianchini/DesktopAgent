using System.Diagnostics;
using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Threading;

namespace DesktopAgent.Tray;

internal partial class QuickChatWindow : Window
{
    private readonly WebApiClient _apiClient;
    private readonly Func<Task>? _runFirstSetup;
    private readonly Func<Task>? _checkUpdatesNow;
    private readonly Action? _applyUpdateNow;
    private readonly Func<string>? _getUpdateStatus;
    private readonly Func<ChatUpdateBadge>? _getChatUpdateBadge;
    private readonly Func<ChatUpdateDetails>? _getChatUpdateDetails;
    private readonly Func<IReadOnlyList<string>>? _loadTimelineSession;
    private readonly Action<IReadOnlyList<string>>? _saveTimelineSession;
    private readonly Queue<string> _historyLines = new();
    private readonly List<string> _timelineEntries = new();
    private readonly List<string> _recentCommands = new();
    private readonly CancellationTokenSource _pollingCts = new();
    private readonly List<WebTaskItem> _taskItems = new();
    private readonly List<WebScheduleItem> _scheduleItems = new();
    private readonly List<WebGoalItem> _goalItems = new();
    private readonly List<GoalRow> _goalRows = new();
    private const int MaxHistoryLines = 500;

    private TextBox? _historyBox;
    private TextBox? _inputBox;
    private TabControl? _mainTabs;
    private TabItem? _chatTab;
    private Border? _chatTabBadgePanel;
    private TextBlock? _chatTabBadgeText;
    private TextBlock? _healthText;
    private ComboBox? _timelineFilterCombo;
    private TextBox? _timelineSearchBox;
    private Button? _cancelRequestButton;
    private Border? _chatUpdateBadgePanel;
    private TextBlock? _chatUpdateBadgeText;
    private Button? _chatUpdateDetailsButton;
    private Button? _chatApplyUpdateButton;
    private ListBox? _timelineList;
    private ListBox? _commandPaletteList;
    private Button? _useSuggestionButton;
    private Border? _busyPanel;
    private TextBlock? _busyText;
    private TextBlock? _statusText;
    private TextBlock? _versionText;
    private StackPanel? _confirmPanel;
    private TextBlock? _confirmText;
    private Button? _sendButton;
    private Button? _confirmButton;
    private Button? _cancelButton;
    private Button? _statusButton;
    private Button? _armButton;
    private Button? _disarmButton;
    private Button? _simPresenceButton;
    private Button? _reqPresenceButton;
    private Button? _killButton;
    private Button? _resetKillButton;
    private Button? _restartAdapterButton;
    private Button? _restartServerButton;
    private Button? _lockWindowButton;
    private Button? _lockAppButton;
    private Button? _unlockButton;
    private Button? _profileSafeButton;
    private Button? _profileBalancedButton;
    private Button? _profilePowerButton;
    private Button? _openWebButton;
    private Button? _copyButton;
    private Button? _clearButton;

    private CheckBox? _cfgLlmEnabled;
    private ComboBox? _cfgLlmProvider;
    private TextBox? _cfgLlmEndpoint;
    private TextBox? _cfgLlmModel;
    private TextBox? _cfgLlmTimeout;
    private TextBox? _cfgLlmMaxTokens;
    private CheckBox? _cfgLlmAllowRemote;
    private Button? _cfgLoadButton;
    private Button? _cfgSaveButton;
    private Button? _cfgTestLlmButton;
    private Button? _cfgRunFirstSetupButton;
    private Button? _cfgCheckUpdatesButton;
    private Button? _cfgApplyUpdateButton;
    private TextBlock? _cfgUpdateStatusText;
    private TextBox? _cfgStatusBox;

    private Button? _tasksRefreshButton;
    private Button? _tasksRunButton;
    private Button? _tasksDeleteButton;
    private ListBox? _tasksList;
    private TextBox? _taskNameInput;
    private TextBox? _taskIntentInput;
    private TextBox? _taskDescriptionInput;
    private Button? _taskSaveButton;
    private TextBlock? _tasksStatusText;

    private Button? _schedulesRefreshButton;
    private Button? _schedulesRunButton;
    private Button? _schedulesDeleteButton;
    private ListBox? _schedulesList;
    private TextBox? _scheduleIdInput;
    private TextBox? _scheduleTaskNameInput;
    private TextBox? _scheduleStartAtInput;
    private TextBox? _scheduleIntervalInput;
    private CheckBox? _scheduleEnabledInput;
    private Button? _scheduleSaveButton;
    private TextBlock? _schedulesStatusText;

    private Button? _goalsRefreshButton;
    private Button? _goalsToggleAutoButton;
    private Button? _goalsDoneButton;
    private Button? _goalsRemoveButton;
    private ListBox? _goalsList;
    private TextBox? _goalTextInput;
    private Button? _goalAddButton;
    private TextBlock? _goalsStatusText;

    private Button? _auditRefreshButton;
    private Button? _auditCopyButton;
    private Button? _auditClearButton;
    private TextBox? _auditBox;

    private string? _pendingToken;
    private string? _aiSuggestedCommand;
    private string? _pendingMessage;
    private string _lastSentMessage = string.Empty;
    private DateTimeOffset _lastSentAtUtc = DateTimeOffset.MinValue;
    private DateTimeOffset _lastUtilityProbeAtUtc = DateTimeOffset.MinValue;
    private bool _ffmpegAvailable;
    private CancellationTokenSource? _activeRequestCts;
    private CancellationTokenSource? _paletteDebounceCts;
    private CancellationTokenSource? _persistSessionCts;
    private bool _busy;
    private bool _loadingSessionState;
    private bool _suppressPaletteSelectionChanged;
    private bool _suppressGoalAutoEvents;
    private string _lastStatusLine = string.Empty;
    private int _unreadChatEvents;
    private int _unreadErrorEvents;
    private int _unreadInfoEvents;
    private const int MaxTimelineLines = 140;
    private const int MaxRecentCommands = 20;
    private readonly string _sessionStatePath = ResolveSessionStatePath();
    private bool? _lastOcrEnabled;
    private static readonly string[] BaseCommandPalette =
    {
        "status",
        "arm",
        "disarm",
        "simulate presence",
        "require presence",
        "open notepad",
        "open vscode",
        "search weather gubbio on chrome",
        "translate Ciao come va to english",
        "take screenshot",
        "record screen for 30 seconds",
        "jiggle mouse for 2 minutes",
        "run open edge and search meteo gubbio",
        "dry-run open calculator"
    };

    public QuickChatWindow(
        WebApiClient apiClient,
        Func<Task>? runFirstSetup = null,
        Func<Task>? checkUpdatesNow = null,
        Action? applyUpdateNow = null,
        Func<string>? getUpdateStatus = null,
        Func<ChatUpdateBadge>? getChatUpdateBadge = null,
        Func<ChatUpdateDetails>? getChatUpdateDetails = null,
        Func<IReadOnlyList<string>>? loadTimelineSession = null,
        Action<IReadOnlyList<string>>? saveTimelineSession = null)
    {
        _apiClient = apiClient;
        _runFirstSetup = runFirstSetup;
        _checkUpdatesNow = checkUpdatesNow;
        _applyUpdateNow = applyUpdateNow;
        _getUpdateStatus = getUpdateStatus;
        _getChatUpdateBadge = getChatUpdateBadge;
        _getChatUpdateDetails = getChatUpdateDetails;
        _loadTimelineSession = loadTimelineSession;
        _saveTimelineSession = saveTimelineSession;
        InitializeComponent();
        WireControls();
    }

    private void WireControls()
    {
        _historyBox = this.FindControl<TextBox>("HistoryBox");
        _inputBox = this.FindControl<TextBox>("InputBox");
        _mainTabs = this.FindControl<TabControl>("MainTabs");
        _chatTab = this.FindControl<TabItem>("ChatTab");
        _chatTabBadgePanel = this.FindControl<Border>("ChatTabBadgePanel");
        _chatTabBadgeText = this.FindControl<TextBlock>("ChatTabBadgeText");
        _healthText = this.FindControl<TextBlock>("HealthText");
        _timelineFilterCombo = this.FindControl<ComboBox>("TimelineFilterCombo");
        _timelineSearchBox = this.FindControl<TextBox>("TimelineSearchBox");
        _cancelRequestButton = this.FindControl<Button>("CancelRequestButton");
        _chatUpdateBadgePanel = this.FindControl<Border>("ChatUpdateBadgePanel");
        _chatUpdateBadgeText = this.FindControl<TextBlock>("ChatUpdateBadgeText");
        _chatUpdateDetailsButton = this.FindControl<Button>("ChatUpdateDetailsButton");
        _chatApplyUpdateButton = this.FindControl<Button>("ChatApplyUpdateButton");
        _timelineList = this.FindControl<ListBox>("TimelineList");
        _commandPaletteList = this.FindControl<ListBox>("CommandPaletteList");
        _useSuggestionButton = this.FindControl<Button>("UseSuggestionButton");
        _busyPanel = this.FindControl<Border>("BusyPanel");
        _busyText = this.FindControl<TextBlock>("BusyText");
        _statusText = this.FindControl<TextBlock>("StatusText");
        _versionText = this.FindControl<TextBlock>("VersionText");
        _confirmPanel = this.FindControl<StackPanel>("ConfirmPanel");
        _confirmText = this.FindControl<TextBlock>("ConfirmText");
        _sendButton = this.FindControl<Button>("SendButton");
        _confirmButton = this.FindControl<Button>("ConfirmButton");
        _cancelButton = this.FindControl<Button>("CancelButton");
        _statusButton = this.FindControl<Button>("StatusButton");
        _armButton = this.FindControl<Button>("ArmButton");
        _disarmButton = this.FindControl<Button>("DisarmButton");
        _simPresenceButton = this.FindControl<Button>("SimPresenceButton");
        _reqPresenceButton = this.FindControl<Button>("ReqPresenceButton");
        _killButton = this.FindControl<Button>("KillButton");
        _resetKillButton = this.FindControl<Button>("ResetKillButton");
        _restartAdapterButton = this.FindControl<Button>("RestartAdapterButton");
        _restartServerButton = this.FindControl<Button>("RestartServerButton");
        _lockWindowButton = this.FindControl<Button>("LockWindowButton");
        _lockAppButton = this.FindControl<Button>("LockAppButton");
        _unlockButton = this.FindControl<Button>("UnlockButton");
        _profileSafeButton = this.FindControl<Button>("ProfileSafeButton");
        _profileBalancedButton = this.FindControl<Button>("ProfileBalancedButton");
        _profilePowerButton = this.FindControl<Button>("ProfilePowerButton");
        _openWebButton = this.FindControl<Button>("OpenWebButton");
        _copyButton = this.FindControl<Button>("CopyButton");
        _clearButton = this.FindControl<Button>("ClearButton");

        _cfgLlmEnabled = this.FindControl<CheckBox>("CfgLlmEnabled");
        _cfgLlmProvider = this.FindControl<ComboBox>("CfgLlmProvider");
        _cfgLlmEndpoint = this.FindControl<TextBox>("CfgLlmEndpoint");
        _cfgLlmModel = this.FindControl<TextBox>("CfgLlmModel");
        _cfgLlmTimeout = this.FindControl<TextBox>("CfgLlmTimeout");
        _cfgLlmMaxTokens = this.FindControl<TextBox>("CfgLlmMaxTokens");
        _cfgLlmAllowRemote = this.FindControl<CheckBox>("CfgLlmAllowRemote");
        _cfgLoadButton = this.FindControl<Button>("CfgLoadButton");
        _cfgSaveButton = this.FindControl<Button>("CfgSaveButton");
        _cfgTestLlmButton = this.FindControl<Button>("CfgTestLlmButton");
        _cfgRunFirstSetupButton = this.FindControl<Button>("CfgRunFirstSetupButton");
        _cfgCheckUpdatesButton = this.FindControl<Button>("CfgCheckUpdatesButton");
        _cfgApplyUpdateButton = this.FindControl<Button>("CfgApplyUpdateButton");
        _cfgUpdateStatusText = this.FindControl<TextBlock>("CfgUpdateStatusText");
        _cfgStatusBox = this.FindControl<TextBox>("CfgStatusBox");

        _tasksRefreshButton = this.FindControl<Button>("TasksRefreshButton");
        _tasksRunButton = this.FindControl<Button>("TasksRunButton");
        _tasksDeleteButton = this.FindControl<Button>("TasksDeleteButton");
        _tasksList = this.FindControl<ListBox>("TasksList");
        _taskNameInput = this.FindControl<TextBox>("TaskNameInput");
        _taskIntentInput = this.FindControl<TextBox>("TaskIntentInput");
        _taskDescriptionInput = this.FindControl<TextBox>("TaskDescriptionInput");
        _taskSaveButton = this.FindControl<Button>("TaskSaveButton");
        _tasksStatusText = this.FindControl<TextBlock>("TasksStatusText");

        _schedulesRefreshButton = this.FindControl<Button>("SchedulesRefreshButton");
        _schedulesRunButton = this.FindControl<Button>("SchedulesRunButton");
        _schedulesDeleteButton = this.FindControl<Button>("SchedulesDeleteButton");
        _schedulesList = this.FindControl<ListBox>("SchedulesList");
        _scheduleIdInput = this.FindControl<TextBox>("ScheduleIdInput");
        _scheduleTaskNameInput = this.FindControl<TextBox>("ScheduleTaskNameInput");
        _scheduleStartAtInput = this.FindControl<TextBox>("ScheduleStartAtInput");
        _scheduleIntervalInput = this.FindControl<TextBox>("ScheduleIntervalInput");
        _scheduleEnabledInput = this.FindControl<CheckBox>("ScheduleEnabledInput");
        _scheduleSaveButton = this.FindControl<Button>("ScheduleSaveButton");
        _schedulesStatusText = this.FindControl<TextBlock>("SchedulesStatusText");

        _goalsRefreshButton = this.FindControl<Button>("GoalsRefreshButton");
        _goalsToggleAutoButton = this.FindControl<Button>("GoalsToggleAutoButton");
        _goalsDoneButton = this.FindControl<Button>("GoalsDoneButton");
        _goalsRemoveButton = this.FindControl<Button>("GoalsRemoveButton");
        _goalsList = this.FindControl<ListBox>("GoalsList");
        _goalTextInput = this.FindControl<TextBox>("GoalTextInput");
        _goalAddButton = this.FindControl<Button>("GoalAddButton");
        _goalsStatusText = this.FindControl<TextBlock>("GoalsStatusText");

        _auditRefreshButton = this.FindControl<Button>("AuditRefreshButton");
        _auditCopyButton = this.FindControl<Button>("AuditCopyButton");
        _auditClearButton = this.FindControl<Button>("AuditClearButton");
        _auditBox = this.FindControl<TextBox>("AuditBox");

        if (_sendButton != null)
        {
            _sendButton.Click += async (_, _) => await SendFromInputAsync();
        }

        if (_inputBox != null)
        {
            _inputBox.KeyDown += async (_, e) =>
            {
                if (e.Key == Key.Up && MovePaletteSelection(-1))
                {
                    e.Handled = true;
                    return;
                }

                if (e.Key == Key.Down && MovePaletteSelection(1))
                {
                    e.Handled = true;
                    return;
                }

                if (e.Key == Key.Tab && ApplyPaletteSuggestion())
                {
                    e.Handled = true;
                    return;
                }

                if (e.Key == Key.Enter)
                {
                    if (ApplyPaletteSuggestion())
                    {
                        e.Handled = true;
                    }

                    e.Handled = true;
                    await SendFromInputAsync();
                }
            };
            _inputBox.PropertyChanged += (_, args) =>
            {
                if (args.Property.Name == nameof(TextBox.Text))
                {
                    ScheduleCommandPaletteUpdate();
                    if (!_loadingSessionState)
                    {
                        SchedulePersistSessionState();
                    }
                }
            };
        }

        if (_timelineFilterCombo != null)
        {
            _timelineFilterCombo.SelectionChanged += (_, _) => ApplyTimelineFilter();
        }

        if (_timelineSearchBox != null)
        {
            _timelineSearchBox.PropertyChanged += (_, args) =>
            {
                if (args.Property.Name == nameof(TextBox.Text))
                {
                    ApplyTimelineFilter();
                }
            };
        }

        if (_commandPaletteList != null)
        {
            _commandPaletteList.SelectionChanged += (_, _) => OnCommandPaletteSelectionChanged();
        }

        if (_useSuggestionButton != null)
        {
            _useSuggestionButton.Click += async (_, _) => await UseAiSuggestionAsync();
        }
        if (_cancelRequestButton != null)
        {
            _cancelRequestButton.Click += (_, _) => CancelActiveRequest();
        }
        if (_chatUpdateDetailsButton != null)
        {
            _chatUpdateDetailsButton.Click += async (_, _) => await ShowUpdateDetailsAsync();
        }
        if (_chatApplyUpdateButton != null)
        {
            _chatApplyUpdateButton.Click += (_, _) => ApplyUpdateFromChat();
        }

        if (_mainTabs != null)
        {
            _mainTabs.SelectionChanged += (_, _) =>
            {
                if (IsChatTabSelected())
                {
                    ResetUnreadChatEvents();
                }
            };
        }

        if (_confirmButton != null)
        {
            _confirmButton.Click += async (_, _) => await ConfirmAsync(true);
        }

        if (_cancelButton != null)
        {
            _cancelButton.Click += async (_, _) => await ConfirmAsync(false);
        }

        HookCommandButton(_statusButton, "status");
        HookCommandButton(_armButton, "arm");
        HookCommandButton(_disarmButton, "disarm");
        HookCommandButton(_simPresenceButton, "simulate presence");
        HookCommandButton(_reqPresenceButton, "require presence");
        HookCommandButton(_killButton, "kill");
        HookCommandButton(_resetKillButton, "reset kill");
        if (_restartAdapterButton != null)
        {
            _restartAdapterButton.Click += async (_, _) => await RestartAdapterAsync();
        }
        if (_restartServerButton != null)
        {
            _restartServerButton.Click += async (_, _) => await RestartServerAsync();
        }
        HookCommandButton(_lockWindowButton, "lock on current window");
        HookCommandButton(_lockAppButton, "lock on app");
        HookCommandButton(_unlockButton, "unlock");
        HookCommandButton(_profileSafeButton, "profile safe");
        HookCommandButton(_profileBalancedButton, "profile balanced");
        HookCommandButton(_profilePowerButton, "profile power");

        if (_openWebButton != null)
        {
            _openWebButton.Click += (_, _) => OpenWebUi();
        }

        if (_copyButton != null)
        {
            _copyButton.Click += async (_, _) => await CopyTextAsync(_historyBox?.Text, "Conversation copied.");
        }

        if (_clearButton != null)
        {
            _clearButton.Click += (_, _) =>
            {
                _historyLines.Clear();
                SetText(_historyBox, string.Empty);
            };
        }

        if (_cfgLoadButton != null)
        {
            _cfgLoadButton.Click += async (_, _) => await LoadConfigAsync();
        }
        if (_cfgSaveButton != null)
        {
            _cfgSaveButton.Click += async (_, _) => await SaveConfigAsync();
        }
        if (_cfgTestLlmButton != null)
        {
            _cfgTestLlmButton.Click += async (_, _) => await TestLlmAsync();
        }
        if (_cfgRunFirstSetupButton != null)
        {
            _cfgRunFirstSetupButton.Click += async (_, _) => await RunFirstSetupAsync();
        }
        if (_cfgCheckUpdatesButton != null)
        {
            _cfgCheckUpdatesButton.Click += async (_, _) => await CheckUpdatesFromConfigAsync();
        }
        if (_cfgApplyUpdateButton != null)
        {
            _cfgApplyUpdateButton.Click += (_, _) => ApplyUpdateFromConfig();
        }

        if (_tasksRefreshButton != null)
        {
            _tasksRefreshButton.Click += async (_, _) => await LoadTasksAsync();
        }
        if (_taskSaveButton != null)
        {
            _taskSaveButton.Click += async (_, _) => await SaveTaskAsync();
        }
        if (_tasksRunButton != null)
        {
            _tasksRunButton.Click += async (_, _) => await RunSelectedTaskAsync();
        }
        if (_tasksDeleteButton != null)
        {
            _tasksDeleteButton.Click += async (_, _) => await DeleteSelectedTaskAsync();
        }

        if (_schedulesRefreshButton != null)
        {
            _schedulesRefreshButton.Click += async (_, _) => await LoadSchedulesAsync();
        }
        if (_scheduleSaveButton != null)
        {
            _scheduleSaveButton.Click += async (_, _) => await SaveScheduleAsync();
        }
        if (_schedulesRunButton != null)
        {
            _schedulesRunButton.Click += async (_, _) => await RunSelectedScheduleAsync();
        }
        if (_schedulesDeleteButton != null)
        {
            _schedulesDeleteButton.Click += async (_, _) => await DeleteSelectedScheduleAsync();
        }

        if (_goalsRefreshButton != null)
        {
            _goalsRefreshButton.Click += async (_, _) => await LoadGoalsAsync();
        }
        if (_goalsToggleAutoButton != null)
        {
            _goalsToggleAutoButton.Click += async (_, _) => await ToggleSelectedGoalAutoAsync();
        }
        if (_goalsDoneButton != null)
        {
            _goalsDoneButton.Click += async (_, _) => await MarkSelectedGoalDoneAsync();
        }
        if (_goalsRemoveButton != null)
        {
            _goalsRemoveButton.Click += async (_, _) => await RemoveSelectedGoalAsync();
        }
        if (_goalAddButton != null)
        {
            _goalAddButton.Click += async (_, _) => await AddGoalAsync();
        }

        if (_auditRefreshButton != null)
        {
            _auditRefreshButton.Click += async (_, _) => await LoadAuditAsync();
        }
        if (_auditCopyButton != null)
        {
            _auditCopyButton.Click += async (_, _) => await CopyTextAsync(_auditBox?.Text, "Audit copied.");
        }
        if (_auditClearButton != null)
        {
            _auditClearButton.Click += (_, _) => SetText(_auditBox, string.Empty);
        }

        Opened += OnOpened;
        Activated += (_, _) =>
        {
            if (IsChatTabSelected())
            {
                ResetUnreadChatEvents();
            }
        };
        Closed += (_, _) =>
        {
            _pollingCts.Cancel();
            _activeRequestCts?.Cancel();
            _activeRequestCts?.Dispose();
            _activeRequestCts = null;
            _paletteDebounceCts?.Cancel();
            _paletteDebounceCts?.Dispose();
            _paletteDebounceCts = null;
            _persistSessionCts?.Cancel();
            _persistSessionCts?.Dispose();
            _persistSessionCts = null;
            PersistSessionState();
        };
    }

    private void HookCommandButton(Button? button, string command)
    {
        if (button == null)
        {
            return;
        }

        button.Click += async (_, _) => await ExecuteQuickCommandAsync(command);
    }

    private async void OnOpened(object? sender, EventArgs e)
    {
        LoadSessionState();
        ResetUnreadChatEvents();
        AppendSystem("Quick chat ready.");
        LoadTimelineSession();
        RefreshChatUpdateBadge();
        UpdateCommandPalette();
        await RefreshStatusAsync();
        await LoadConfigAsync();
        RefreshUpdateStatusInConfig();
        await LoadTasksAsync();
        await LoadSchedulesAsync();
        await LoadGoalsAsync();
        await LoadAuditAsync();
        _ = Task.Run(() => PollStatusAsync(_pollingCts.Token));
    }

    private async Task PollStatusAsync(CancellationToken cancellationToken)
    {
        using var timer = new PeriodicTimer(TimeSpan.FromSeconds(3));
        try
        {
            while (await timer.WaitForNextTickAsync(cancellationToken))
            {
                await RefreshStatusAsync();
            }
        }
        catch (OperationCanceledException)
        {
            // Normal shutdown.
        }
    }

    private async Task SendFromInputAsync()
    {
        if (_inputBox == null)
        {
            return;
        }

        var text = _inputBox.Text?.Trim() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(text))
        {
            return;
        }

        text = SanitizeInputForCommand(text);
        if (string.IsNullOrWhiteSpace(text))
        {
            return;
        }

        _inputBox.Text = string.Empty;
        await SendMessageAsync(text);
    }

    private async Task ExecuteQuickCommandAsync(string command)
    {
        await SendMessageAsync(command);
    }

    private async Task SendMessageAsync(string message)
    {
        var normalized = message?.Trim() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(normalized))
        {
            return;
        }

        if (_busy)
        {
            _pendingMessage = normalized;
            _activeRequestCts?.Cancel();
            AppendSystem("Cancelling previous request...");
            return;
        }

        if (string.Equals(_lastSentMessage, normalized, StringComparison.OrdinalIgnoreCase)
            && DateTimeOffset.UtcNow - _lastSentAtUtc < TimeSpan.FromMilliseconds(700))
        {
            AppendSystem("Ignored duplicate command (debounced).");
            return;
        }

        _lastSentMessage = normalized;
        _lastSentAtUtc = DateTimeOffset.UtcNow;
        AddRecentCommand(normalized);

        _activeRequestCts = CancellationTokenSource.CreateLinkedTokenSource(_pollingCts.Token);
        var requestToken = _activeRequestCts.Token;

        SetBusy(true);
        AppendUser(normalized);
        SetTimeline(new[] { "[..] Waiting for response..." });
        try
        {
            // Run agent call off the UI thread because intent rewriting may perform
            // blocking network operations (LLM fallback) in the current implementation.
            var response = await Task.Run(
                () => _apiClient.SendChatAsync(normalized, requestToken),
                requestToken);
            RenderResponse(response);
        }
        catch (OperationCanceledException)
        {
            AppendSystem("Request canceled.");
        }
        catch (Exception ex)
        {
            AppendSystem($"Error: {ex.Message}");
        }
        finally
        {
            _activeRequestCts?.Dispose();
            _activeRequestCts = null;
            SetBusy(false);
            await RefreshStatusAsync();

            var next = _pendingMessage;
            _pendingMessage = null;
            if (!string.IsNullOrWhiteSpace(next))
            {
                await SendMessageAsync(next);
            }
        }
    }

    private async Task ConfirmAsync(bool approve)
    {
        if (_busy || string.IsNullOrWhiteSpace(_pendingToken))
        {
            return;
        }

        SetBusy(true);
        SetTimeline(new[] { "[..] Waiting for response..." });
        try
        {
            var response = await Task.Run(
                () => _apiClient.ConfirmAsync(_pendingToken, approve, CancellationToken.None),
                CancellationToken.None);
            RenderResponse(response);
        }
        catch (Exception ex)
        {
            AppendSystem($"Error: {ex.Message}");
        }
        finally
        {
            _pendingToken = null;
            ShowConfirm(false, null);
            SetBusy(false);
            await RefreshStatusAsync();
        }
    }

    private void RenderResponse(WebChatResponse response)
    {
        var reply = string.IsNullOrWhiteSpace(response.Reply) ? "<no reply>" : response.Reply;
        AppendAgent(reply);
        UpdateAiSuggestionState(reply);

        if (!string.IsNullOrWhiteSpace(response.ModeLabel))
        {
            AppendSystem(response.ModeLabel!);
        }

        if (response.Steps is { Count: > 0 })
        {
            SetTimeline(response.Steps);
            foreach (var step in response.Steps)
            {
                AppendSystem(step);
            }
        }
        else
        {
            SetTimeline(new[] { "[..] No step details." });
        }

        if (response.NeedsConfirmation && !string.IsNullOrWhiteSpace(response.Token))
        {
            _pendingToken = response.Token;
            var prompt = string.IsNullOrWhiteSpace(response.ActionLabel)
                ? "Confirmation required."
                : $"{response.ActionLabel} required.";
            ShowConfirm(true, prompt);
        }
        else
        {
            _pendingToken = null;
            ShowConfirm(false, null);
        }
    }

    private async Task RefreshStatusAsync()
    {
        try
        {
            var snapshot = await _apiClient.GetStatusAsync(CancellationToken.None);
            if (snapshot?.Adapter == null)
            {
                SetText(_statusText, "Status unavailable");
                return;
            }

            var llmLabel = snapshot.Llm == null
                ? "LLM:unknown"
                : snapshot.Llm.Enabled && snapshot.Llm.Available ? "LLM:on" : "LLM:off";
            var killLabel = snapshot.KillSwitch?.Tripped == true ? "KILL:on" : "KILL:off";
            var armedLabel = snapshot.Adapter.Armed ? "ARMED:on" : "ARMED:off";
            var presenceLabel = snapshot.Adapter.RequireUserPresence ? "PRESENCE:req" : "PRESENCE:off";
            var statusLine = $"{armedLabel} | {presenceLabel} | {llmLabel} | {killLabel}";
            if (!string.Equals(statusLine, _lastStatusLine, StringComparison.Ordinal))
            {
                _lastStatusLine = statusLine;
                SetText(_statusText, statusLine);
            }

            if (!string.IsNullOrWhiteSpace(snapshot.Version))
            {
                SetText(_versionText, $"Version: {snapshot.Version}");
            }

            RefreshChatUpdateBadge();
            await RefreshHealthPanelAsync(snapshot);
        }
        catch
        {
            SetText(_statusText, "Status unavailable");
            SetText(_healthText, "Health: adapter down | grpc down | llm ? | ocr ? | ffmpeg ?");
        }
    }

    private async Task RestartAdapterAsync()
    {
        try
        {
            var response = await _apiClient.RestartAdapterAsync(CancellationToken.None);
            AppendSystem(response?.Message ?? "Adapter restart requested.");
        }
        catch (Exception ex)
        {
            AppendSystem($"Adapter restart failed: {ex.Message}");
        }

        await Task.Delay(700);
        await RefreshStatusAsync();
    }

    private async Task RestartServerAsync()
    {
        try
        {
            var response = await _apiClient.RestartServerAsync(CancellationToken.None);
            AppendSystem(response?.Message ?? "Server restart requested.");
        }
        catch (Exception ex)
        {
            AppendSystem($"Server restart failed: {ex.Message}");
        }
    }

    private void SetTimeline(IEnumerable<string> lines)
    {
        var items = lines
            .Where(line => !string.IsNullOrWhiteSpace(line))
            .Select(FormatTimelineLine)
            .ToList();
        if (items.Count == 0)
        {
            items.Add("[..] No timeline.");
        }

        foreach (var item in items)
        {
            _timelineEntries.Add(WithTimestamp(item));
        }

        while (_timelineEntries.Count > MaxTimelineLines)
        {
            _timelineEntries.RemoveAt(0);
        }

        PersistTimelineSession();
        SchedulePersistSessionState();
        ApplyTimelineFilter();
    }

    private void ApplyTimelineFilter()
    {
        Dispatcher.UIThread.Post(() =>
        {
            if (_timelineList == null)
            {
                return;
            }

            var filter = (_timelineFilterCombo?.SelectedItem as ComboBoxItem)?.Content?.ToString() ?? "All";
            var query = _timelineSearchBox?.Text?.Trim() ?? string.Empty;

            IEnumerable<string> filtered = _timelineEntries;
            if (string.Equals(filter, "Success", StringComparison.OrdinalIgnoreCase))
            {
                filtered = filtered.Where(static line => line.Contains("[OK]", StringComparison.OrdinalIgnoreCase));
            }
            else if (string.Equals(filter, "Failed", StringComparison.OrdinalIgnoreCase))
            {
                filtered = filtered.Where(static line => line.Contains("[FAIL]", StringComparison.OrdinalIgnoreCase));
            }

            if (!string.IsNullOrWhiteSpace(query))
            {
                filtered = filtered.Where(line => line.Contains(query, StringComparison.OrdinalIgnoreCase));
            }

            var items = filtered.ToList();
            if (items.Count == 0)
            {
                items.Add("[..] No timeline.");
            }

            _timelineList.ItemsSource = items;
        });
    }

    private async Task ShowUpdateDetailsAsync()
    {
        var details = _getChatUpdateDetails?.Invoke()
            ?? new ChatUpdateDetails(false, "Updates", "No update information available.", false);

        var dialog = new Window
        {
            Title = details.Title,
            Width = 620,
            Height = 420,
            MinWidth = 520,
            MinHeight = 320,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            CanResize = true
        };

        var contentBox = new TextBox
        {
            IsReadOnly = true,
            AcceptsReturn = true,
            TextWrapping = Avalonia.Media.TextWrapping.Wrap,
            VerticalContentAlignment = Avalonia.Layout.VerticalAlignment.Top,
            Background = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#0F1520")),
            Foreground = new Avalonia.Media.SolidColorBrush(Avalonia.Media.Color.Parse("#E6EDF7")),
            BorderThickness = new Thickness(0),
            Text = string.IsNullOrWhiteSpace(details.Details) ? "No details available." : details.Details
        };

        var closeButton = new Button
        {
            Content = "Close",
            MinWidth = 90
        };
        closeButton.Click += (_, _) => dialog.Close();

        var applyButton = new Button
        {
            Content = "Apply now",
            MinWidth = 100,
            IsVisible = details.CanApply
        };
        applyButton.Click += (_, _) =>
        {
            ApplyUpdateFromChat();
            dialog.Close();
        };

        var footer = new StackPanel
        {
            Orientation = Avalonia.Layout.Orientation.Horizontal,
            Spacing = 8,
            HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Right
        };
        if (details.CanApply)
        {
            footer.Children.Add(applyButton);
        }

        footer.Children.Add(closeButton);

        var root = new Grid
        {
            RowDefinitions = new RowDefinitions("*,Auto"),
            Margin = new Thickness(12)
        };
        root.Children.Add(contentBox);
        Grid.SetRow(footer, 1);
        root.Children.Add(footer);
        dialog.Content = root;

        if (IsVisible)
        {
            await dialog.ShowDialog(this);
            return;
        }

        var tcs = new TaskCompletionSource<object?>();
        void ClosedHandler(object? sender, EventArgs args) => tcs.TrySetResult(null);
        dialog.Closed += ClosedHandler;
        dialog.Show();
        await tcs.Task;
        dialog.Closed -= ClosedHandler;
    }

    private void CancelActiveRequest()
    {
        if (!_busy)
        {
            return;
        }

        _activeRequestCts?.Cancel();
        AppendSystem("Cancel requested.");
    }

    private static string WithTimestamp(string line)
    {
        var stamp = DateTime.Now.ToString("HH:mm:ss", CultureInfo.InvariantCulture);
        return $"[{stamp}] {line}";
    }

    private void LoadTimelineSession()
    {
        if (_loadTimelineSession == null || _timelineEntries.Count > 0)
        {
            return;
        }

        try
        {
            var saved = _loadTimelineSession();
            if (saved == null || saved.Count == 0)
            {
                return;
            }

            _timelineEntries.Clear();
            _timelineEntries.AddRange(saved.Where(static line => !string.IsNullOrWhiteSpace(line)).TakeLast(MaxTimelineLines));
            ApplyTimelineFilter();
        }
        catch
        {
            // ignored
        }
    }

    private void PersistTimelineSession()
    {
        if (_saveTimelineSession == null)
        {
            return;
        }

        try
        {
            _saveTimelineSession(_timelineEntries.ToList());
        }
        catch
        {
            // ignored
        }
    }

    private static string FormatTimelineLine(string line)
    {
        var normalized = line.Trim();
        if (normalized.Contains("=> True", StringComparison.OrdinalIgnoreCase)
            || normalized.Contains(":True", StringComparison.OrdinalIgnoreCase))
        {
            return $"[OK] {normalized}";
        }

        if (normalized.Contains("=> False", StringComparison.OrdinalIgnoreCase)
            || normalized.Contains(":False", StringComparison.OrdinalIgnoreCase)
            || normalized.StartsWith("Error", StringComparison.OrdinalIgnoreCase))
        {
            return $"[FAIL] {normalized}";
        }

        return $"[..] {normalized}";
    }

    private void UpdateAiSuggestionState(string reply)
    {
        _aiSuggestedCommand = ExtractAiSuggestion(reply);
        UpdateUseSuggestionButton();
        UpdateCommandPalette();
    }

    private void UpdateUseSuggestionButton()
    {
        Dispatcher.UIThread.Post(() =>
        {
            if (_useSuggestionButton == null)
            {
                return;
            }

            var hasSuggestion = !string.IsNullOrWhiteSpace(_aiSuggestedCommand);
            _useSuggestionButton.IsEnabled = !_busy && hasSuggestion;
            _useSuggestionButton.Content = hasSuggestion ? "Use Suggestion" : "No Suggestion";
        });
    }

    private static string? ExtractAiSuggestion(string reply)
    {
        if (string.IsNullOrWhiteSpace(reply))
        {
            return null;
        }

        var match = Regex.Match(reply, @"AI suggestion:\s*(.+)", RegexOptions.IgnoreCase);
        if (!match.Success)
        {
            return null;
        }

        var value = match.Groups[1].Value
            .Replace("\r", " ")
            .Replace("\n", " ")
            .Trim();
        return string.IsNullOrWhiteSpace(value) ? null : value;
    }

    private void UpdateCommandPalette()
    {
        Dispatcher.UIThread.Post(() =>
        {
            if (_commandPaletteList == null)
            {
                return;
            }

            var input = _inputBox?.Text?.Trim() ?? string.Empty;
            var tokens = input.Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            var candidates = new List<string>();
            if (!string.IsNullOrWhiteSpace(_aiSuggestedCommand))
            {
                candidates.Add(_aiSuggestedCommand!);
            }

            candidates.AddRange(_recentCommands);
            candidates.AddRange(BaseCommandPalette);

            var entries = candidates
                .Where(cmd => !string.IsNullOrWhiteSpace(cmd))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Where(cmd => tokens.Length == 0 || tokens.All(token => cmd.Contains(token, StringComparison.OrdinalIgnoreCase)))
                .Take(8)
                .ToList();

            _suppressPaletteSelectionChanged = true;
            try
            {
                _commandPaletteList.ItemsSource = entries;
                _commandPaletteList.SelectedIndex = -1;
                _commandPaletteList.IsVisible = entries.Count > 0 && !_busy;
            }
            finally
            {
                _suppressPaletteSelectionChanged = false;
            }
        });
    }

    private void ScheduleCommandPaletteUpdate()
    {
        _paletteDebounceCts?.Cancel();
        _paletteDebounceCts?.Dispose();
        var cts = new CancellationTokenSource();
        _paletteDebounceCts = cts;

        _ = Task.Run(async () =>
        {
            try
            {
                await Task.Delay(120, cts.Token);
                if (!cts.IsCancellationRequested)
                {
                    UpdateCommandPalette();
                }
            }
            catch (OperationCanceledException)
            {
                // ignored
            }
        }, CancellationToken.None);
    }

    private void OnCommandPaletteSelectionChanged()
    {
        if (_suppressPaletteSelectionChanged)
        {
            return;
        }

        ApplyCommandPaletteSelection();
    }

    private void ApplyCommandPaletteSelection()
    {
        if (_commandPaletteList?.SelectedItem is not string selected || string.IsNullOrWhiteSpace(selected))
        {
            return;
        }

        if (_inputBox != null)
        {
            _inputBox.Text = selected;
            _inputBox.CaretIndex = _inputBox.Text?.Length ?? 0;
        }
    }

    private bool MovePaletteSelection(int direction)
    {
        if (_commandPaletteList == null || !_commandPaletteList.IsVisible || direction == 0)
        {
            return false;
        }

        var items = (_commandPaletteList.ItemsSource as IEnumerable<string>)?.ToList() ?? new List<string>();
        if (items.Count == 0)
        {
            return false;
        }

        var currentIndex = _commandPaletteList.SelectedIndex;
        if (currentIndex < 0)
        {
            currentIndex = direction > 0 ? -1 : 0;
        }

        var nextIndex = Math.Clamp(currentIndex + direction, 0, items.Count - 1);
        if (nextIndex == currentIndex && _commandPaletteList.SelectedIndex >= 0)
        {
            return true;
        }

        _commandPaletteList.SelectedIndex = nextIndex;
        return true;
    }

    private bool ApplyPaletteSuggestion()
    {
        if (_commandPaletteList == null || !_commandPaletteList.IsVisible)
        {
            return false;
        }

        if (_commandPaletteList.SelectedItem is string selected && !string.IsNullOrWhiteSpace(selected))
        {
            ApplyCommandPaletteSelection();
            return true;
        }

        var first = (_commandPaletteList.ItemsSource as IEnumerable<string>)?.FirstOrDefault();
        if (string.IsNullOrWhiteSpace(first))
        {
            return false;
        }

        _commandPaletteList.SelectedIndex = 0;
        ApplyCommandPaletteSelection();
        return true;
    }

    private async Task UseAiSuggestionAsync()
    {
        var suggestion = _aiSuggestedCommand?.Trim();
        if (string.IsNullOrWhiteSpace(suggestion) || _busy)
        {
            return;
        }

        await SendMessageAsync(suggestion);
    }

    private async Task LoadConfigAsync()
    {
        try
        {
            var config = await _apiClient.GetConfigAsync(CancellationToken.None);
            if (config?.Llm == null)
            {
                AppendConfigStatus("Config unavailable.");
                return;
            }

            if (_cfgLlmEnabled != null)
            {
                _cfgLlmEnabled.IsChecked = config.Llm.Enabled;
            }

            if (_cfgLlmAllowRemote != null)
            {
                _cfgLlmAllowRemote.IsChecked = config.Llm.AllowNonLoopbackEndpoint;
            }

            SetText(_cfgLlmEndpoint, config.Llm.Endpoint ?? string.Empty);
            SetText(_cfgLlmModel, config.Llm.Model ?? string.Empty);
            SetText(_cfgLlmTimeout, config.Llm.TimeoutSeconds.ToString(CultureInfo.InvariantCulture));
            SetText(_cfgLlmMaxTokens, config.Llm.MaxTokens.ToString(CultureInfo.InvariantCulture));

            SelectProvider(config.Llm.Provider ?? "ollama");
            AppendConfigStatus("Config loaded.");
            RefreshUpdateStatusInConfig();
        }
        catch (Exception ex)
        {
            AppendConfigStatus($"Config load failed: {ex.Message}");
        }
    }

    private async Task SaveConfigAsync()
    {
        try
        {
            var update = new WebConfigUpdate(
                Llm: new WebConfigUpdateLlm(
                    Enabled: _cfgLlmEnabled?.IsChecked ?? false,
                    AllowNonLoopbackEndpoint: _cfgLlmAllowRemote?.IsChecked ?? false,
                    Provider: GetSelectedProvider(),
                    Endpoint: _cfgLlmEndpoint?.Text?.Trim(),
                    Model: _cfgLlmModel?.Text?.Trim(),
                    TimeoutSeconds: ParseInt(_cfgLlmTimeout?.Text),
                    MaxTokens: ParseInt(_cfgLlmMaxTokens?.Text)),
                ProfileModeEnabled: null,
                ActiveProfile: null,
                RequireConfirmation: null,
                MaxActionsPerSecond: null,
                QuizSafeModeEnabled: null,
                OcrEnabled: null,
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
                AuditLlmIncludeRawText: null);

            var response = await _apiClient.SaveConfigAsync(update, CancellationToken.None);
            if (response == null)
            {
                AppendConfigStatus("Config save failed.");
                return;
            }

            AppendConfigStatus("Config saved.");
            await TestLlmAsync();
        }
        catch (Exception ex)
        {
            AppendConfigStatus($"Config save failed: {ex.Message}");
        }
    }

    private async Task TestLlmAsync()
    {
        try
        {
            var status = await _apiClient.GetStatusAsync(CancellationToken.None);
            var llm = status?.Llm;
            if (llm == null)
            {
                AppendConfigStatus("LLM status unavailable.");
                return;
            }

            var availability = llm.Enabled && llm.Available ? "available" : "not available";
            AppendConfigStatus($"LLM {availability}. Provider={llm.Provider}; Message={llm.Message}; Endpoint={llm.Endpoint}");
        }
        catch (Exception ex)
        {
            AppendConfigStatus($"LLM test failed: {ex.Message}");
        }
    }

    private async Task RunFirstSetupAsync()
    {
        if (_runFirstSetup == null)
        {
            AppendConfigStatus("First setup trigger is not available.");
            return;
        }

        try
        {
            AppendConfigStatus("Opening first setup wizard...");
            // Run setup flow off the UI thread because environment probes and
            // package checks may perform blocking process waits.
            await Task.Run(_runFirstSetup, CancellationToken.None);
            AppendConfigStatus("First setup completed.");
        }
        catch (Exception ex)
        {
            AppendConfigStatus($"First setup failed: {ex.Message}");
        }
    }

    private async Task CheckUpdatesFromConfigAsync()
    {
        if (_checkUpdatesNow == null)
        {
            AppendConfigStatus("Update check is not available.");
            return;
        }

        try
        {
            AppendConfigStatus("Checking updates...");
            await _checkUpdatesNow();
            RefreshUpdateStatusInConfig();
            RefreshChatUpdateBadge();
            AppendConfigStatus("Update check completed.");
        }
        catch (Exception ex)
        {
            AppendConfigStatus($"Update check failed: {ex.Message}");
        }
    }

    private void ApplyUpdateFromConfig()
    {
        if (_applyUpdateNow == null)
        {
            AppendConfigStatus("Apply update is not available.");
            return;
        }

        try
        {
            _applyUpdateNow();
            RefreshUpdateStatusInConfig();
            RefreshChatUpdateBadge();
            AppendConfigStatus("Apply update requested.");
        }
        catch (Exception ex)
        {
            AppendConfigStatus($"Apply update failed: {ex.Message}");
        }
    }

    private void RefreshUpdateStatusInConfig()
    {
        if (_cfgUpdateStatusText == null)
        {
            return;
        }

        var text = _getUpdateStatus?.Invoke() ?? "Updates: unavailable";
        SetText(_cfgUpdateStatusText, text);
        if (_cfgApplyUpdateButton != null)
        {
            var canApply = _getChatUpdateBadge?.Invoke().CanApply ?? false;
            _cfgApplyUpdateButton.IsEnabled = canApply && !_busy;
        }
    }

    private void RefreshChatUpdateBadge()
    {
        if (_chatUpdateBadgePanel == null || _chatUpdateBadgeText == null || _chatApplyUpdateButton == null || _chatUpdateDetailsButton == null)
        {
            return;
        }

        var badge = _getChatUpdateBadge?.Invoke() ?? new ChatUpdateBadge(false, false, string.Empty);
        Dispatcher.UIThread.Post(() =>
        {
            _chatUpdateBadgePanel.IsVisible = badge.Visible;
            _chatUpdateBadgeText.Text = badge.Text;
            _chatApplyUpdateButton.IsEnabled = badge.CanApply && !_busy;
            _chatUpdateDetailsButton.IsEnabled = !_busy;
        });
    }

    private void ApplyUpdateFromChat()
    {
        ApplyUpdateFromConfig();
        RefreshChatUpdateBadge();
    }

    private async Task LoadTasksAsync()
    {
        try
        {
            var response = await _apiClient.GetTasksAsync(CancellationToken.None);
            _taskItems.Clear();
            if (response?.Tasks != null)
            {
                _taskItems.AddRange(response.Tasks.OrderByDescending(t => t.UpdatedAt));
            }

            if (_tasksList != null)
            {
                _tasksList.ItemsSource = _taskItems.Select(FormatTask).ToList();
            }

            SetText(_tasksStatusText, $"Tasks: {_taskItems.Count}");
        }
        catch (Exception ex)
        {
            SetText(_tasksStatusText, $"Tasks load failed: {ex.Message}");
        }
    }

    private async Task SaveTaskAsync()
    {
        var name = _taskNameInput?.Text?.Trim() ?? string.Empty;
        var intent = _taskIntentInput?.Text?.Trim() ?? string.Empty;
        var description = _taskDescriptionInput?.Text?.Trim();
        if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(intent))
        {
            SetText(_tasksStatusText, "Task name and intent are required.");
            return;
        }

        try
        {
            var response = await _apiClient.SaveTaskAsync(new WebTaskUpsertRequest(name, intent, description, null), CancellationToken.None);
            SetText(_tasksStatusText, response?.Message ?? "Task saved.");
            await LoadTasksAsync();
        }
        catch (Exception ex)
        {
            SetText(_tasksStatusText, $"Task save failed: {ex.Message}");
        }
    }

    private async Task RunSelectedTaskAsync()
    {
        var selected = SelectedTask();
        if (selected == null)
        {
            SetText(_tasksStatusText, "Select a task.");
            return;
        }

        try
        {
            var result = await _apiClient.RunTaskAsync(selected.Name, false, CancellationToken.None);
            if (result == null)
            {
                SetText(_tasksStatusText, "Task run failed.");
                return;
            }

            AppendSystem($"Task '{selected.Name}' run: {result.Reply}");
            if (result.Steps != null)
            {
                foreach (var step in result.Steps)
                {
                    AppendSystem(step);
                }
            }
        }
        catch (Exception ex)
        {
            SetText(_tasksStatusText, $"Task run failed: {ex.Message}");
        }
    }

    private async Task DeleteSelectedTaskAsync()
    {
        var selected = SelectedTask();
        if (selected == null)
        {
            SetText(_tasksStatusText, "Select a task.");
            return;
        }

        try
        {
            var result = await _apiClient.DeleteTaskAsync(selected.Name, CancellationToken.None);
            SetText(_tasksStatusText, result?.Message ?? "Task deleted.");
            await LoadTasksAsync();
        }
        catch (Exception ex)
        {
            SetText(_tasksStatusText, $"Task delete failed: {ex.Message}");
        }
    }

    private async Task LoadSchedulesAsync()
    {
        try
        {
            var response = await _apiClient.GetSchedulesAsync(CancellationToken.None);
            _scheduleItems.Clear();
            if (response?.Schedules != null)
            {
                _scheduleItems.AddRange(response.Schedules.OrderByDescending(s => s.UpdatedAt));
            }

            if (_schedulesList != null)
            {
                _schedulesList.ItemsSource = _scheduleItems.Select(FormatSchedule).ToList();
            }

            SetText(_schedulesStatusText, $"Schedules: {_scheduleItems.Count}");
        }
        catch (Exception ex)
        {
            SetText(_schedulesStatusText, $"Schedules load failed: {ex.Message}");
        }
    }

    private async Task SaveScheduleAsync()
    {
        var taskName = _scheduleTaskNameInput?.Text?.Trim() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(taskName))
        {
            SetText(_schedulesStatusText, "Task name is required.");
            return;
        }

        DateTimeOffset? startAtUtc = null;
        var startRaw = _scheduleStartAtInput?.Text?.Trim();
        if (!string.IsNullOrWhiteSpace(startRaw))
        {
            if (!DateTimeOffset.TryParse(startRaw, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out var parsed))
            {
                SetText(_schedulesStatusText, "Invalid Start UTC format.");
                return;
            }
            startAtUtc = parsed.ToUniversalTime();
        }

        var interval = ParseInt(_scheduleIntervalInput?.Text);

        try
        {
            var request = new WebScheduleUpsertRequest(
                Id: string.IsNullOrWhiteSpace(_scheduleIdInput?.Text) ? null : _scheduleIdInput?.Text?.Trim(),
                TaskName: taskName,
                StartAtUtc: startAtUtc,
                IntervalSeconds: interval,
                Enabled: _scheduleEnabledInput?.IsChecked ?? true);

            var result = await _apiClient.SaveScheduleAsync(request, CancellationToken.None);
            SetText(_schedulesStatusText, result?.Message ?? "Schedule saved.");
            await LoadSchedulesAsync();
        }
        catch (Exception ex)
        {
            SetText(_schedulesStatusText, $"Schedule save failed: {ex.Message}");
        }
    }

    private async Task RunSelectedScheduleAsync()
    {
        var selected = SelectedSchedule();
        if (selected == null)
        {
            SetText(_schedulesStatusText, "Select a schedule.");
            return;
        }

        try
        {
            var result = await _apiClient.RunScheduleNowAsync(selected.Id, CancellationToken.None);
            SetText(_schedulesStatusText, result?.Message ?? "Schedule triggered.");
        }
        catch (Exception ex)
        {
            SetText(_schedulesStatusText, $"Schedule run failed: {ex.Message}");
        }
    }

    private async Task DeleteSelectedScheduleAsync()
    {
        var selected = SelectedSchedule();
        if (selected == null)
        {
            SetText(_schedulesStatusText, "Select a schedule.");
            return;
        }

        try
        {
            var result = await _apiClient.DeleteScheduleAsync(selected.Id, CancellationToken.None);
            SetText(_schedulesStatusText, result?.Message ?? "Schedule deleted.");
            await LoadSchedulesAsync();
        }
        catch (Exception ex)
        {
            SetText(_schedulesStatusText, $"Schedule delete failed: {ex.Message}");
        }
    }

    private async Task LoadGoalsAsync()
    {
        try
        {
            var response = await _apiClient.GetGoalsAsync(CancellationToken.None);
            _goalItems.Clear();
            _goalRows.Clear();
            if (response?.Goals != null)
            {
                _goalItems.AddRange(response.Goals);
                _goalRows.AddRange(response.Goals.Select(ToGoalRow));
            }

            if (_goalsList != null)
            {
                _suppressGoalAutoEvents = true;
                _goalsList.ItemsSource = _goalRows;
                _suppressGoalAutoEvents = false;
            }

            var scheduler = response == null
                ? "scheduler:unknown"
                : $"scheduler:{(response.SchedulerEnabled ? "on" : "off")} every {response.SchedulerIntervalSeconds}s";
            SetText(_goalsStatusText, $"Goals: {_goalItems.Count} | {scheduler}");
        }
        catch (Exception ex)
        {
            SetText(_goalsStatusText, $"Goals load failed: {ex.Message}");
        }
    }

    private async Task AddGoalAsync()
    {
        var text = _goalTextInput?.Text?.Trim() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(text))
        {
            SetText(_goalsStatusText, "Goal text is required.");
            return;
        }

        try
        {
            var result = await _apiClient.AddGoalAsync(text, CancellationToken.None);
            SetText(_goalsStatusText, result?.Message ?? "Goal added.");
            SetText(_goalTextInput, string.Empty);
            await LoadGoalsAsync();
        }
        catch (Exception ex)
        {
            SetText(_goalsStatusText, $"Goal add failed: {ex.Message}");
        }
    }

    private async Task ToggleSelectedGoalAutoAsync()
    {
        var selected = SelectedGoal();
        if (selected == null)
        {
            SetText(_goalsStatusText, "Select a goal.");
            return;
        }

        try
        {
            var enabled = !selected.AutoRunEnabled;
            var result = await _apiClient.SetGoalAutoAsync(selected.Id, enabled, CancellationToken.None);
            SetText(_goalsStatusText, result?.Message ?? "Goal updated.");
            await LoadGoalsAsync();
        }
        catch (Exception ex)
        {
            SetText(_goalsStatusText, $"Goal update failed: {ex.Message}");
        }
    }

    private async Task MarkSelectedGoalDoneAsync()
    {
        var selected = SelectedGoal();
        if (selected == null)
        {
            SetText(_goalsStatusText, "Select a goal.");
            return;
        }

        try
        {
            var result = await _apiClient.MarkGoalDoneAsync(selected.Id, CancellationToken.None);
            SetText(_goalsStatusText, result?.Message ?? "Goal marked as done.");
            await LoadGoalsAsync();
        }
        catch (Exception ex)
        {
            SetText(_goalsStatusText, $"Goal update failed: {ex.Message}");
        }
    }

    private async Task RemoveSelectedGoalAsync()
    {
        var selected = SelectedGoal();
        if (selected == null)
        {
            SetText(_goalsStatusText, "Select a goal.");
            return;
        }

        try
        {
            var result = await _apiClient.RemoveGoalAsync(selected.Id, CancellationToken.None);
            SetText(_goalsStatusText, result?.Message ?? "Goal removed.");
            await LoadGoalsAsync();
        }
        catch (Exception ex)
        {
            SetText(_goalsStatusText, $"Goal remove failed: {ex.Message}");
        }
    }

    private async Task LoadAuditAsync()
    {
        try
        {
            var response = await _apiClient.GetAuditAsync(200, CancellationToken.None);
            var lines = response?.Lines ?? Array.Empty<string>();
            SetText(_auditBox, string.Join(Environment.NewLine, lines));
        }
        catch (Exception ex)
        {
            SetText(_auditBox, $"Audit load failed: {ex.Message}");
        }
    }

    private WebTaskItem? SelectedTask()
    {
        var index = _tasksList?.SelectedIndex ?? -1;
        if (index < 0 || index >= _taskItems.Count)
        {
            return null;
        }

        return _taskItems[index];
    }

    private WebScheduleItem? SelectedSchedule()
    {
        var index = _schedulesList?.SelectedIndex ?? -1;
        if (index < 0 || index >= _scheduleItems.Count)
        {
            return null;
        }

        return _scheduleItems[index];
    }

    private WebGoalItem? SelectedGoal()
    {
        var row = _goalsList?.SelectedItem as GoalRow;
        if (row == null)
        {
            return null;
        }

        return _goalItems.FirstOrDefault(goal => string.Equals(goal.Id, row.Id, StringComparison.OrdinalIgnoreCase));
    }

    private void ShowConfirm(bool show, string? text)
    {
        Dispatcher.UIThread.Post(() =>
        {
            if (_confirmPanel != null)
            {
                _confirmPanel.IsVisible = show;
            }

            if (_confirmText != null && !string.IsNullOrWhiteSpace(text))
            {
                _confirmText.Text = text;
            }
        });
    }

    private void SetBusy(bool busy)
    {
        _busy = busy;
        Dispatcher.UIThread.Post(() =>
        {
            if (_busyPanel != null)
            {
                _busyPanel.IsVisible = busy;
            }
            if (_busyText != null)
            {
                _busyText.Text = busy ? "Waiting for agent response..." : string.Empty;
            }

            foreach (var button in AllButtons())
            {
                if (ReferenceEquals(button, _cancelRequestButton))
                {
                    button.IsEnabled = busy;
                    continue;
                }

                button.IsEnabled = !busy;
            }
        });
        UpdateUseSuggestionButton();
        UpdateCommandPalette();
        RefreshChatUpdateBadge();
        RefreshUpdateStatusInConfig();
    }

    private IEnumerable<Button> AllButtons()
    {
        return new[]
        {
            _sendButton, _confirmButton, _cancelButton, _statusButton, _armButton, _disarmButton, _simPresenceButton,
            _useSuggestionButton, _chatUpdateDetailsButton, _chatApplyUpdateButton, _cancelRequestButton,
            _reqPresenceButton, _killButton, _resetKillButton, _restartAdapterButton, _restartServerButton,
            _lockWindowButton, _lockAppButton, _unlockButton, _profileSafeButton,
            _profileBalancedButton, _profilePowerButton, _openWebButton, _copyButton, _clearButton,
            _cfgLoadButton, _cfgSaveButton, _cfgTestLlmButton, _cfgRunFirstSetupButton, _cfgCheckUpdatesButton, _cfgApplyUpdateButton, _tasksRefreshButton, _tasksRunButton,
            _tasksDeleteButton, _taskSaveButton, _schedulesRefreshButton, _schedulesRunButton, _schedulesDeleteButton,
            _scheduleSaveButton, _goalsRefreshButton, _goalsToggleAutoButton, _goalsDoneButton,
            _goalsRemoveButton, _goalAddButton, _auditRefreshButton, _auditCopyButton, _auditClearButton
        }.Where(button => button != null).Cast<Button>();
    }

    private bool IsChatTabSelected()
    {
        if (_mainTabs == null || _chatTab == null)
        {
            return true;
        }

        return ReferenceEquals(_mainTabs.SelectedItem, _chatTab);
    }

    private void ResetUnreadChatEvents()
    {
        _unreadChatEvents = 0;
        _unreadErrorEvents = 0;
        _unreadInfoEvents = 0;
        UpdateUnreadBadge();
    }

    private void UpdateUnreadBadge()
    {
        Dispatcher.UIThread.Post(() =>
        {
            if (_chatTabBadgePanel == null || _chatTabBadgeText == null)
            {
                return;
            }

            var hasUnread = _unreadChatEvents > 0;
            _chatTabBadgePanel.IsVisible = hasUnread;
            if (!hasUnread)
            {
                ToolTip.SetTip(_chatTabBadgePanel, null);
                return;
            }

            _chatTabBadgeText.Text = _unreadChatEvents > 99 ? "99+" : _unreadChatEvents.ToString(CultureInfo.InvariantCulture);

            var hasErrors = _unreadErrorEvents > 0;
            _chatTabBadgePanel.Background = new Avalonia.Media.SolidColorBrush(
                Avalonia.Media.Color.Parse(hasErrors ? "#8A2C34" : "#2C4F8A"));
            _chatTabBadgePanel.BorderBrush = new Avalonia.Media.SolidColorBrush(
                Avalonia.Media.Color.Parse(hasErrors ? "#D96B73" : "#4F79BA"));
            _chatTabBadgePanel.BorderThickness = new Thickness(1);
            ToolTip.SetTip(_chatTabBadgePanel, $"Unread: {_unreadChatEvents} (errors: {_unreadErrorEvents}, info: {_unreadInfoEvents})");
            _chatTabBadgeText.Text = hasErrors
                ? $"! {_chatTabBadgeText.Text}"
                : _chatTabBadgeText.Text;
        });
    }

    private void AppendUser(string text) => AppendLine("YOU", text);
    private void AppendAgent(string text) => AppendLine("AGENT", text);
    private void AppendSystem(string text) => AppendLine("SYSTEM", text);

    private async Task CopyTextAsync(string? text, string successMessage)
    {
        var value = text ?? string.Empty;
        if (string.IsNullOrWhiteSpace(value))
        {
            AppendSystem("Nothing to copy.");
            return;
        }

        try
        {
            var clipboard = TopLevel.GetTopLevel(this)?.Clipboard;
            if (clipboard == null)
            {
                AppendSystem("Clipboard unavailable.");
                return;
            }

            await clipboard.SetTextAsync(value);
            AppendSystem(successMessage);
        }
        catch (Exception ex)
        {
            AppendSystem($"Copy failed: {ex.Message}");
        }
    }

    private void AppendLine(string role, string text)
    {
        Dispatcher.UIThread.Post(() =>
        {
            var timestamp = DateTime.Now.ToString("HH:mm:ss", CultureInfo.InvariantCulture);
            var cleanText = text.Replace("\r", string.Empty).Replace('\n', ' ').Trim();
            _historyLines.Enqueue($"[{timestamp}] {role,-6} {cleanText}");
            while (_historyLines.Count > MaxHistoryLines)
            {
                _historyLines.Dequeue();
            }

            SetText(_historyBox, string.Join(Environment.NewLine, _historyLines));
            if (_historyBox != null)
            {
                _historyBox.CaretIndex = _historyBox.Text?.Length ?? 0;
            }

            if (!string.Equals(role, "YOU", StringComparison.OrdinalIgnoreCase) && !IsChatTabSelected())
            {
                _unreadChatEvents = Math.Min(_unreadChatEvents + 1, 999);
                if (IsErrorLine(cleanText))
                {
                    _unreadErrorEvents = Math.Min(_unreadErrorEvents + 1, 999);
                }
                else
                {
                    _unreadInfoEvents = Math.Min(_unreadInfoEvents + 1, 999);
                }

                UpdateUnreadBadge();
            }

            SchedulePersistSessionState();
        });
    }

    private static bool IsErrorLine(string line)
    {
        if (string.IsNullOrWhiteSpace(line))
        {
            return false;
        }

        return line.Contains("error", StringComparison.OrdinalIgnoreCase)
            || line.Contains("failed", StringComparison.OrdinalIgnoreCase)
            || line.Contains("blocked", StringComparison.OrdinalIgnoreCase)
            || line.Contains("denied", StringComparison.OrdinalIgnoreCase)
            || line.Contains("unavailable", StringComparison.OrdinalIgnoreCase)
            || line.Contains("[FAIL]", StringComparison.OrdinalIgnoreCase);
    }

    private void AppendConfigStatus(string line)
    {
        Dispatcher.UIThread.Post(() =>
        {
            var current = _cfgStatusBox?.Text ?? string.Empty;
            var stamp = DateTime.Now.ToString("HH:mm:ss", CultureInfo.InvariantCulture);
            var entry = $"[{stamp}] {line}";
            SetText(_cfgStatusBox, string.IsNullOrWhiteSpace(current) ? entry : $"{current}{Environment.NewLine}{entry}");
            if (_cfgStatusBox != null)
            {
                _cfgStatusBox.CaretIndex = _cfgStatusBox.Text?.Length ?? 0;
            }
        });
    }

    private static void SetText(TextBox? textBox, string value)
    {
        if (textBox == null)
        {
            return;
        }

        textBox.Text = value;
    }

    private static void SetText(TextBlock? textBlock, string value)
    {
        if (textBlock == null)
        {
            return;
        }

        textBlock.Text = value;
    }

    private void SelectProvider(string provider)
    {
        if (_cfgLlmProvider == null)
        {
            return;
        }

        var normalized = provider.Trim().ToLowerInvariant();
        for (var i = 0; i < _cfgLlmProvider.ItemCount; i++)
        {
            if (_cfgLlmProvider.Items[i] is ComboBoxItem item &&
                string.Equals(item.Content?.ToString(), normalized, StringComparison.OrdinalIgnoreCase))
            {
                _cfgLlmProvider.SelectedIndex = i;
                return;
            }
        }

        _cfgLlmProvider.SelectedIndex = 0;
    }

    private string GetSelectedProvider()
    {
        if (_cfgLlmProvider?.SelectedItem is ComboBoxItem combo && combo.Content != null)
        {
            return combo.Content.ToString() ?? "ollama";
        }

        return "ollama";
    }

    private static int? ParseInt(string? value)
    {
        if (int.TryParse(value?.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed))
        {
            return parsed;
        }

        return null;
    }

    private static string FormatTask(WebTaskItem task)
    {
        var description = string.IsNullOrWhiteSpace(task.Description) ? task.Intent : task.Description;
        return $"{task.Name} | {description} | {task.UpdatedAt:yyyy-MM-dd HH:mm}";
    }

    private static string FormatSchedule(WebScheduleItem schedule)
    {
        var start = schedule.StartAtUtc?.ToString("yyyy-MM-dd HH:mm 'UTC'", CultureInfo.InvariantCulture) ?? "immediate";
        var interval = schedule.IntervalSeconds.HasValue ? $"{schedule.IntervalSeconds.Value}s" : "once";
        var enabled = schedule.Enabled ? "enabled" : "disabled";
        return $"{schedule.Id} | {schedule.TaskName} | {start} | {interval} | {enabled}";
    }

    private static string FormatGoal(WebGoalItem goal)
    {
        var status = goal.Completed ? "done" : "open";
        var priority = goal.Priority switch
        {
            <= 0 => "low",
            >= 2 => "high",
            _ => "normal"
        };
        var auto = goal.AutoRunEnabled ? "auto:on" : "auto:off";
        var attempts = goal.Attempts > 0 ? $" attempts:{goal.Attempts}" : string.Empty;
        return $"[{goal.Id}] {status} [{priority}] {auto} - {goal.Text}{attempts}";
    }

    private static GoalRow ToGoalRow(WebGoalItem goal)
    {
        return new GoalRow
        {
            Id = goal.Id,
            AutoRunEnabled = goal.AutoRunEnabled,
            Label = FormatGoal(goal)
        };
    }

    private async void OnGoalAutoChecked(object? sender, RoutedEventArgs e)
    {
        await UpdateGoalAutoFromRowAsync(sender, true);
    }

    private async void OnGoalAutoUnchecked(object? sender, RoutedEventArgs e)
    {
        await UpdateGoalAutoFromRowAsync(sender, false);
    }

    private async Task UpdateGoalAutoFromRowAsync(object? sender, bool enabled)
    {
        if (_suppressGoalAutoEvents || sender is not CheckBox checkBox || checkBox.DataContext is not GoalRow row)
        {
            return;
        }

        try
        {
            var result = await _apiClient.SetGoalAutoAsync(row.Id, enabled, CancellationToken.None);
            SetText(_goalsStatusText, result?.Message ?? "Goal updated.");
            await LoadGoalsAsync();
        }
        catch (Exception ex)
        {
            SetText(_goalsStatusText, $"Goal update failed: {ex.Message}");
        }
    }

    private void AddRecentCommand(string command)
    {
        var normalized = command.Trim();
        if (string.IsNullOrWhiteSpace(normalized))
        {
            return;
        }

        _recentCommands.RemoveAll(existing => string.Equals(existing, normalized, StringComparison.OrdinalIgnoreCase));
        _recentCommands.Insert(0, normalized);
        if (_recentCommands.Count > MaxRecentCommands)
        {
            _recentCommands.RemoveRange(MaxRecentCommands, _recentCommands.Count - MaxRecentCommands);
        }

        SchedulePersistSessionState();
    }

    private async Task RefreshHealthPanelAsync(WebStatusResponse snapshot)
    {
        var adapterUp = snapshot.Adapter != null;
        var grpcUp = adapterUp;
        var llmEnabled = snapshot.Llm?.Enabled == true;
        var llmOk = llmEnabled && snapshot.Llm?.Available == true;

        if (_lastUtilityProbeAtUtc == DateTimeOffset.MinValue || DateTimeOffset.UtcNow - _lastUtilityProbeAtUtc >= TimeSpan.FromSeconds(20))
        {
            _lastUtilityProbeAtUtc = DateTimeOffset.UtcNow;
            _ffmpegAvailable = await Task.Run(() => CommandExists("ffmpeg"), CancellationToken.None);

            try
            {
                var config = await _apiClient.GetConfigAsync(CancellationToken.None);
                if (config != null)
                {
                    _lastOcrEnabled = config.OcrEnabled;
                }
            }
            catch
            {
                // ignored
            }
        }

        var ocrLabel = _lastOcrEnabled switch
        {
            true => "[+]",
            false => "[-]",
            _ => "[?]"
        };
        var ffmpegLabel = _ffmpegAvailable ? "[+]" : "[-]";
        var llmLabel = !llmEnabled ? "[-]" : llmOk ? "[+]" : "[!]";
        var text = $"Health: adapter {(adapterUp ? "[+]" : "[-]")} | grpc {(grpcUp ? "[+]" : "[-]")} | llm {llmLabel} | ocr {ocrLabel} | ffmpeg {ffmpegLabel}";
        SetText(_healthText, text);
    }

    private static bool CommandExists(string command)
    {
        try
        {
            var checker = OperatingSystem.IsWindows() ? "where" : "which";
            using var process = Process.Start(new ProcessStartInfo
            {
                FileName = checker,
                Arguments = command,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden
            });
            if (process == null)
            {
                return false;
            }

            process.WaitForExit(2500);
            return process.HasExited && process.ExitCode == 0;
        }
        catch
        {
            return false;
        }
    }

    private void LoadSessionState()
    {
        try
        {
            if (!File.Exists(_sessionStatePath))
            {
                return;
            }

            var json = File.ReadAllText(_sessionStatePath);
            var state = JsonSerializer.Deserialize<ChatSessionState>(json);
            if (state == null)
            {
                return;
            }

            _loadingSessionState = true;

            _historyLines.Clear();
            foreach (var line in state.History.TakeLast(MaxHistoryLines))
            {
                if (!string.IsNullOrWhiteSpace(line))
                {
                    _historyLines.Enqueue(line);
                }
            }

            _timelineEntries.Clear();
            _timelineEntries.AddRange(state.Timeline.Where(static line => !string.IsNullOrWhiteSpace(line)).TakeLast(MaxTimelineLines));

            _recentCommands.Clear();
            _recentCommands.AddRange(state.RecentCommands.Where(static line => !string.IsNullOrWhiteSpace(line)).Take(MaxRecentCommands));

            SetText(_historyBox, string.Join(Environment.NewLine, _historyLines));
            if (_inputBox != null && !string.IsNullOrWhiteSpace(state.InputText))
            {
                _inputBox.Text = state.InputText;
                _inputBox.CaretIndex = _inputBox.Text?.Length ?? 0;
            }

            ApplyTimelineFilter();
            UpdateCommandPalette();
        }
        catch
        {
            // ignored
        }
        finally
        {
            _loadingSessionState = false;
        }
    }

    private void SchedulePersistSessionState()
    {
        if (_loadingSessionState)
        {
            return;
        }

        _persistSessionCts?.Cancel();
        _persistSessionCts?.Dispose();
        var cts = new CancellationTokenSource();
        _persistSessionCts = cts;

        _ = Task.Run(async () =>
        {
            try
            {
                await Task.Delay(250, cts.Token);
                if (!cts.IsCancellationRequested)
                {
                    PersistSessionState();
                }
            }
            catch (OperationCanceledException)
            {
                // ignored
            }
        }, CancellationToken.None);
    }

    private void PersistSessionState()
    {
        if (_loadingSessionState)
        {
            return;
        }

        try
        {
            var state = new ChatSessionState(
                _historyLines.TakeLast(MaxHistoryLines).ToList(),
                _timelineEntries.TakeLast(MaxTimelineLines).ToList(),
                _recentCommands.Take(MaxRecentCommands).ToList(),
                _inputBox?.Text ?? string.Empty);

            var parent = Path.GetDirectoryName(_sessionStatePath);
            if (!string.IsNullOrWhiteSpace(parent))
            {
                Directory.CreateDirectory(parent);
            }

            var json = JsonSerializer.Serialize(state, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(_sessionStatePath, json);
        }
        catch
        {
            // ignored
        }
    }

    private static string ResolveSessionStatePath()
    {
        var root = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "DesktopAgent");
        Directory.CreateDirectory(root);
        return Path.Combine(root, "tray-chat-session.json");
    }

    private static string SanitizeInputForCommand(string input)
    {
        var value = input.Trim();
        for (var i = 0; i < 3; i++)
        {
            var cleaned = Regex.Replace(value, "^(you|agent|system)\\s*:\\s*", string.Empty, RegexOptions.IgnoreCase).Trim();
            if (string.Equals(cleaned, value, StringComparison.Ordinal))
            {
                break;
            }

            value = cleaned;
        }

        return value;
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

    private sealed class GoalRow
    {
        public string Id { get; init; } = string.Empty;
        public string Label { get; init; } = string.Empty;
        public bool AutoRunEnabled { get; init; }
    }
}

internal sealed record ChatSessionState(
    List<string> History,
    List<string> Timeline,
    List<string> RecentCommands,
    string InputText);

internal sealed record ChatUpdateBadge(bool Visible, bool CanApply, string Text);
internal sealed record ChatUpdateDetails(bool HasUpdate, string Title, string Details, bool CanApply);
