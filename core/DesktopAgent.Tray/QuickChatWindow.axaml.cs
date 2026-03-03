using System.Diagnostics;
using System.Globalization;
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
    private readonly Queue<string> _historyLines = new();
    private readonly CancellationTokenSource _pollingCts = new();
    private readonly List<WebTaskItem> _taskItems = new();
    private readonly List<WebScheduleItem> _scheduleItems = new();
    private readonly List<WebGoalItem> _goalItems = new();
    private readonly List<GoalRow> _goalRows = new();
    private const int MaxHistoryLines = 500;

    private TextBox? _historyBox;
    private TextBox? _inputBox;
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
    private bool _busy;
    private bool _suppressGoalAutoEvents;
    private string _lastStatusLine = string.Empty;

    public QuickChatWindow(WebApiClient apiClient)
    {
        _apiClient = apiClient;
        InitializeComponent();
        WireControls();
    }

    private void WireControls()
    {
        _historyBox = this.FindControl<TextBox>("HistoryBox");
        _inputBox = this.FindControl<TextBox>("InputBox");
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
                if (e.Key == Key.Enter)
                {
                    e.Handled = true;
                    await SendFromInputAsync();
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
        Closed += (_, _) => _pollingCts.Cancel();
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
        AppendSystem("Quick chat ready.");
        await RefreshStatusAsync();
        await LoadConfigAsync();
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
        if (_busy)
        {
            return;
        }

        SetBusy(true);
        AppendUser(message);
        try
        {
            // Run agent call off the UI thread because intent rewriting may perform
            // blocking network operations (LLM fallback) in the current implementation.
            var response = await Task.Run(
                () => _apiClient.SendChatAsync(message, CancellationToken.None),
                CancellationToken.None);
            RenderResponse(response);
        }
        catch (Exception ex)
        {
            AppendSystem($"Error: {ex.Message}");
        }
        finally
        {
            SetBusy(false);
            await RefreshStatusAsync();
        }
    }

    private async Task ConfirmAsync(bool approve)
    {
        if (_busy || string.IsNullOrWhiteSpace(_pendingToken))
        {
            return;
        }

        SetBusy(true);
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

        if (!string.IsNullOrWhiteSpace(response.ModeLabel))
        {
            AppendSystem(response.ModeLabel!);
        }

        if (response.Steps is { Count: > 0 })
        {
            foreach (var step in response.Steps)
            {
                AppendSystem(step);
            }
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
        }
        catch
        {
            SetText(_statusText, "Status unavailable");
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
            foreach (var button in AllButtons())
            {
                button.IsEnabled = !busy;
            }
        });
    }

    private IEnumerable<Button> AllButtons()
    {
        return new[]
        {
            _sendButton, _confirmButton, _cancelButton, _statusButton, _armButton, _disarmButton, _simPresenceButton,
            _reqPresenceButton, _killButton, _resetKillButton, _restartAdapterButton, _restartServerButton,
            _lockWindowButton, _lockAppButton, _unlockButton, _profileSafeButton,
            _profileBalancedButton, _profilePowerButton, _openWebButton, _copyButton, _clearButton,
            _cfgLoadButton, _cfgSaveButton, _cfgTestLlmButton, _tasksRefreshButton, _tasksRunButton,
            _tasksDeleteButton, _taskSaveButton, _schedulesRefreshButton, _schedulesRunButton, _schedulesDeleteButton,
            _scheduleSaveButton, _goalsRefreshButton, _goalsToggleAutoButton, _goalsDoneButton,
            _goalsRemoveButton, _goalAddButton, _auditRefreshButton, _auditCopyButton, _auditClearButton
        }.Where(button => button != null).Cast<Button>();
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
        });
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
