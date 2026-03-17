using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Text.RegularExpressions;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Avalonia.Threading;

namespace DesktopAgent.Tray;

internal partial class QuickChatWindow : Window
{
    private readonly WebApiClient _apiClient;
    private readonly Func<Task>? _runFirstSetup;
    private readonly Func<Task>? _checkUpdatesNow;
    private readonly Action? _applyUpdateNow;
    private readonly Action? _forceApplyUpdateNow;
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
    private TabItem? _timelineTab;
    private Border? _chatTabBadgePanel;
    private TextBlock? _chatTabBadgeText;
    private Border? _timelineTabBadgePanel;
    private TextBlock? _timelineTabBadgeText;
    private Border? _armedStateBadge;
    private TextBlock? _armedStateText;
    private TextBlock? _armedStateIcon;
    private Border? _presenceStateBadge;
    private TextBlock? _presenceStateText;
    private TextBlock? _presenceStateIcon;
    private Border? _llmStateBadge;
    private TextBlock? _llmStateText;
    private TextBlock? _llmStateIcon;
    private Border? _killStateBadge;
    private TextBlock? _killStateText;
    private TextBlock? _killStateIcon;
    private TextBlock? _healthText;
    private ComboBox? _timelineFilterCombo;
    private TextBox? _timelineSearchBox;
    private Button? _cancelRequestButton;
    private Border? _chatUpdateBadgePanel;
    private TextBlock? _chatUpdateBadgeText;
    private Button? _chatUpdateDetailsButton;
    private Button? _chatApplyUpdateButton;
    private ListBox? _timelineList;
    private Button? _timelineCopyButton;
    private Button? _timelineExportButton;
    private TextBlock? _timelineStatusText;
    private ListBox? _commandPaletteList;
    private Button? _useSuggestionButton;
    private Border? _busyPanel;
    private TextBlock? _busyText;
    private Border? _quickActionsPanel;
    private Button? _quickRunSuggestionButton;
    private Button? _quickDryRunButton;
    private Button? _quickExplainPlanButton;
    private Button? _quickEditPromptButton;
    private Button? _quickHelpButton;
    private Border? _planEditorPanel;
    private TextBlock? _planEditorTitleText;
    private TextBox? _planEditorBox;
    private TextBox? _planEditorHumanBox;
    private TextBlock? _planEditorStatusText;
    private Button? _planEditorLoadButton;
    private Button? _planEditorValidateButton;
    private Button? _planEditorDryRunButton;
    private Button? _planEditorExecuteButton;
    private Button? _planEditorRefreshHumanButton;
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
    private Button? _copyButton;
    private Button? _clearButton;

    private CheckBox? _cfgLlmEnabled;
    private ComboBox? _cfgLlmProvider;
    private ComboBox? _cfgLlmParsingMode;
    private ComboBox? _cfgAudioBackend;
    private TextBox? _cfgAudioDevice;
    private ComboBox? _cfgAudioDeviceSelector;
    private Button? _cfgAudioRefreshButton;
    private CheckBox? _cfgPrimaryDisplayOnly;
    private TextBox? _cfgLlmEndpoint;
    private TextBox? _cfgLlmModel;
    private TextBox? _cfgLlmTimeout;
    private TextBox? _cfgLlmMaxTokens;
    private CheckBox? _cfgLlmAllowRemote;
    private TextBox? _cfgMediaOutputDirectory;
    private Button? _cfgOpenDataFolderButton;
    private Button? _cfgLoadButton;
    private Button? _cfgSaveButton;
    private Button? _cfgTestLlmButton;
    private Button? _cfgRunFirstSetupButton;
    private Button? _cfgCheckUpdatesButton;
    private Button? _cfgApplyUpdateButton;
    private Button? _cfgForceApplyUpdateButton;
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
    private Button? _macroRecordToggleButton;
    private Button? _macroClearButton;
    private ListBox? _macroList;
    private TextBox? _macroStepEditorInput;
    private Button? _macroApplyEditButton;
    private Button? _macroRemoveStepButton;
    private Button? _macroMoveUpButton;
    private Button? _macroMoveDownButton;
    private TextBox? _macroAddWaitSecondsInput;
    private Button? _macroAddWaitButton;
    private TextBox? _macroNameInput;
    private TextBox? _macroDescriptionInput;
    private Button? _macroSaveTaskButton;
    private TextBlock? _macroStatusText;

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

    private Button? _mediaRefreshButton;
    private Button? _mediaOpenFolderButton;
    private Button? _mediaOpenButton;
    private Button? _mediaCopyPathButton;
    private Button? _mediaDeleteButton;
    private ListBox? _mediaList;
    private TextBlock? _mediaStatusText;
    private TextBlock? _mediaPreviewTitleText;
    private TextBox? _mediaPreviewInfoText;
    private Image? _mediaPreviewImage;

    private Button? _auditRefreshButton;
    private Button? _auditCopyButton;
    private Button? _auditClearButton;
    private TextBox? _auditBox;
    private Button? _diagRefreshButton;
    private Button? _diagCopyButton;
    private TextBlock? _diagStatusText;
    private TextBox? _diagBox;
    private TextBox? _wikiSearchBox;
    private Button? _wikiResetSearchButton;
    private Button? _wikiCopyButton;
    private TextBlock? _wikiStatusText;
    private TextBox? _wikiContentBox;

    private string? _pendingToken;
    private string? _aiSuggestedCommand;
    private string? _pendingMessage;
    private string? _lastPlanJson;
    private string _lastSentMessage = string.Empty;
    private string _lastReply = string.Empty;
    private DateTimeOffset _lastSentAtUtc = DateTimeOffset.MinValue;
    private DateTimeOffset _lastUtilityProbeAtUtc = DateTimeOffset.MinValue;
    private DateTimeOffset _busyStartedAtUtc = DateTimeOffset.MinValue;
    private double _busyAvgSeconds = 7;
    private bool _ffmpegAvailable;
    private CancellationTokenSource? _activeRequestCts;
    private CancellationTokenSource? _busyAnimationCts;
    private CancellationTokenSource? _streamingCts;
    private CancellationTokenSource? _paletteDebounceCts;
    private CancellationTokenSource? _persistSessionCts;
    private bool _busy;
    private bool _macroRecording;
    private bool _loadingSessionState;
    private bool _suppressPaletteSelectionChanged;
    private bool _suppressGoalAutoEvents;
    private string _configuredMediaOutputDirectory = "media";
    private string _lastStatusLine = string.Empty;
    private int _unreadChatEvents;
    private int _unreadErrorEvents;
    private int _unreadInfoEvents;
    private int _unreadTimelineEvents;
    private const int MaxTimelineLines = 140;
    private const int MaxRecentCommands = 20;
    private const int MinPaletteInputChars = 3;
    private readonly string _sessionStatePath = ResolveSessionStatePath();
    private readonly List<MacroStep> _macroSteps = new();
    private readonly List<MediaFileItem> _mediaItems = new();
    private Bitmap? _mediaPreviewBitmap;
    private bool? _lastOcrEnabled;
    private string _wikiFullText = string.Empty;
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
        Action? forceApplyUpdateNow = null,
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
        _forceApplyUpdateNow = forceApplyUpdateNow;
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
        _timelineTab = this.FindControl<TabItem>("TimelineTab");
        _chatTabBadgePanel = this.FindControl<Border>("ChatTabBadgePanel");
        _chatTabBadgeText = this.FindControl<TextBlock>("ChatTabBadgeText");
        _timelineTabBadgePanel = this.FindControl<Border>("TimelineTabBadgePanel");
        _timelineTabBadgeText = this.FindControl<TextBlock>("TimelineTabBadgeText");
        _armedStateBadge = this.FindControl<Border>("ArmedStateBadge");
        _armedStateText = this.FindControl<TextBlock>("ArmedStateText");
        _armedStateIcon = this.FindControl<TextBlock>("ArmedStateIcon");
        _presenceStateBadge = this.FindControl<Border>("PresenceStateBadge");
        _presenceStateText = this.FindControl<TextBlock>("PresenceStateText");
        _presenceStateIcon = this.FindControl<TextBlock>("PresenceStateIcon");
        _llmStateBadge = this.FindControl<Border>("LlmStateBadge");
        _llmStateText = this.FindControl<TextBlock>("LlmStateText");
        _llmStateIcon = this.FindControl<TextBlock>("LlmStateIcon");
        _killStateBadge = this.FindControl<Border>("KillStateBadge");
        _killStateText = this.FindControl<TextBlock>("KillStateText");
        _killStateIcon = this.FindControl<TextBlock>("KillStateIcon");
        _healthText = this.FindControl<TextBlock>("HealthText");
        _timelineFilterCombo = this.FindControl<ComboBox>("TimelineFilterCombo");
        _timelineSearchBox = this.FindControl<TextBox>("TimelineSearchBox");
        _cancelRequestButton = this.FindControl<Button>("CancelRequestButton");
        _chatUpdateBadgePanel = this.FindControl<Border>("ChatUpdateBadgePanel");
        _chatUpdateBadgeText = this.FindControl<TextBlock>("ChatUpdateBadgeText");
        _chatUpdateDetailsButton = this.FindControl<Button>("ChatUpdateDetailsButton");
        _chatApplyUpdateButton = this.FindControl<Button>("ChatApplyUpdateButton");
        _timelineList = this.FindControl<ListBox>("TimelineList");
        _timelineCopyButton = this.FindControl<Button>("TimelineCopyButton");
        _timelineExportButton = this.FindControl<Button>("TimelineExportButton");
        _timelineStatusText = this.FindControl<TextBlock>("TimelineStatusText");
        _commandPaletteList = this.FindControl<ListBox>("CommandPaletteList");
        _useSuggestionButton = this.FindControl<Button>("UseSuggestionButton");
        _busyPanel = this.FindControl<Border>("BusyPanel");
        _busyText = this.FindControl<TextBlock>("BusyText");
        _quickActionsPanel = this.FindControl<Border>("QuickActionsPanel");
        _quickRunSuggestionButton = this.FindControl<Button>("QuickRunSuggestionButton");
        _quickDryRunButton = this.FindControl<Button>("QuickDryRunButton");
        _quickExplainPlanButton = this.FindControl<Button>("QuickExplainPlanButton");
        _quickEditPromptButton = this.FindControl<Button>("QuickEditPromptButton");
        _quickHelpButton = this.FindControl<Button>("QuickHelpButton");
        _planEditorPanel = this.FindControl<Border>("PlanEditorPanel");
        _planEditorTitleText = this.FindControl<TextBlock>("PlanEditorTitleText");
        _planEditorBox = this.FindControl<TextBox>("PlanEditorBox");
        _planEditorHumanBox = this.FindControl<TextBox>("PlanEditorHumanBox");
        _planEditorStatusText = this.FindControl<TextBlock>("PlanEditorStatusText");
        _planEditorLoadButton = this.FindControl<Button>("PlanEditorLoadButton");
        _planEditorValidateButton = this.FindControl<Button>("PlanEditorValidateButton");
        _planEditorDryRunButton = this.FindControl<Button>("PlanEditorDryRunButton");
        _planEditorExecuteButton = this.FindControl<Button>("PlanEditorExecuteButton");
        _planEditorRefreshHumanButton = this.FindControl<Button>("PlanEditorRefreshHumanButton");
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
        _copyButton = this.FindControl<Button>("CopyButton");
        _clearButton = this.FindControl<Button>("ClearButton");

        _cfgLlmEnabled = this.FindControl<CheckBox>("CfgLlmEnabled");
        _cfgLlmProvider = this.FindControl<ComboBox>("CfgLlmProvider");
        _cfgLlmParsingMode = this.FindControl<ComboBox>("CfgLlmParsingMode");
        _cfgAudioBackend = this.FindControl<ComboBox>("CfgAudioBackend");
        _cfgAudioDevice = this.FindControl<TextBox>("CfgAudioDevice");
        _cfgAudioDeviceSelector = this.FindControl<ComboBox>("CfgAudioDeviceSelector");
        _cfgAudioRefreshButton = this.FindControl<Button>("CfgAudioRefreshButton");
        _cfgPrimaryDisplayOnly = this.FindControl<CheckBox>("CfgPrimaryDisplayOnly");
        _cfgLlmEndpoint = this.FindControl<TextBox>("CfgLlmEndpoint");
        _cfgLlmModel = this.FindControl<TextBox>("CfgLlmModel");
        _cfgLlmTimeout = this.FindControl<TextBox>("CfgLlmTimeout");
        _cfgLlmMaxTokens = this.FindControl<TextBox>("CfgLlmMaxTokens");
        _cfgLlmAllowRemote = this.FindControl<CheckBox>("CfgLlmAllowRemote");
        _cfgMediaOutputDirectory = this.FindControl<TextBox>("CfgMediaOutputDirectory");
        _cfgOpenDataFolderButton = this.FindControl<Button>("CfgOpenDataFolderButton");
        _cfgLoadButton = this.FindControl<Button>("CfgLoadButton");
        _cfgSaveButton = this.FindControl<Button>("CfgSaveButton");
        _cfgTestLlmButton = this.FindControl<Button>("CfgTestLlmButton");
        _cfgRunFirstSetupButton = this.FindControl<Button>("CfgRunFirstSetupButton");
        _cfgCheckUpdatesButton = this.FindControl<Button>("CfgCheckUpdatesButton");
        _cfgApplyUpdateButton = this.FindControl<Button>("CfgApplyUpdateButton");
        _cfgForceApplyUpdateButton = this.FindControl<Button>("CfgForceApplyUpdateButton");
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
        _macroRecordToggleButton = this.FindControl<Button>("MacroRecordToggleButton");
        _macroClearButton = this.FindControl<Button>("MacroClearButton");
        _macroList = this.FindControl<ListBox>("MacroList");
        _macroStepEditorInput = this.FindControl<TextBox>("MacroStepEditorInput");
        _macroApplyEditButton = this.FindControl<Button>("MacroApplyEditButton");
        _macroRemoveStepButton = this.FindControl<Button>("MacroRemoveStepButton");
        _macroMoveUpButton = this.FindControl<Button>("MacroMoveUpButton");
        _macroMoveDownButton = this.FindControl<Button>("MacroMoveDownButton");
        _macroAddWaitSecondsInput = this.FindControl<TextBox>("MacroAddWaitSecondsInput");
        _macroAddWaitButton = this.FindControl<Button>("MacroAddWaitButton");
        _macroNameInput = this.FindControl<TextBox>("MacroNameInput");
        _macroDescriptionInput = this.FindControl<TextBox>("MacroDescriptionInput");
        _macroSaveTaskButton = this.FindControl<Button>("MacroSaveTaskButton");
        _macroStatusText = this.FindControl<TextBlock>("MacroStatusText");

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

        _mediaRefreshButton = this.FindControl<Button>("MediaRefreshButton");
        _mediaOpenFolderButton = this.FindControl<Button>("MediaOpenFolderButton");
        _mediaOpenButton = this.FindControl<Button>("MediaOpenButton");
        _mediaCopyPathButton = this.FindControl<Button>("MediaCopyPathButton");
        _mediaDeleteButton = this.FindControl<Button>("MediaDeleteButton");
        _mediaList = this.FindControl<ListBox>("MediaList");
        _mediaStatusText = this.FindControl<TextBlock>("MediaStatusText");
        _mediaPreviewTitleText = this.FindControl<TextBlock>("MediaPreviewTitleText");
        _mediaPreviewInfoText = this.FindControl<TextBox>("MediaPreviewInfoText");
        _mediaPreviewImage = this.FindControl<Image>("MediaPreviewImage");

        _auditRefreshButton = this.FindControl<Button>("AuditRefreshButton");
        _auditCopyButton = this.FindControl<Button>("AuditCopyButton");
        _auditClearButton = this.FindControl<Button>("AuditClearButton");
        _auditBox = this.FindControl<TextBox>("AuditBox");
        _diagRefreshButton = this.FindControl<Button>("DiagRefreshButton");
        _diagCopyButton = this.FindControl<Button>("DiagCopyButton");
        _diagStatusText = this.FindControl<TextBlock>("DiagStatusText");
        _diagBox = this.FindControl<TextBox>("DiagBox");
        _wikiSearchBox = this.FindControl<TextBox>("WikiSearchBox");
        _wikiResetSearchButton = this.FindControl<Button>("WikiResetSearchButton");
        _wikiCopyButton = this.FindControl<Button>("WikiCopyButton");
        _wikiStatusText = this.FindControl<TextBlock>("WikiStatusText");
        _wikiContentBox = this.FindControl<TextBox>("WikiContentBox");

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
        if (_quickRunSuggestionButton != null)
        {
            _quickRunSuggestionButton.Click += async (_, _) => await UseAiSuggestionAsync();
        }
        if (_quickDryRunButton != null)
        {
            _quickDryRunButton.Click += async (_, _) => await RunDryRunLastCommandAsync();
        }
        if (_quickExplainPlanButton != null)
        {
            _quickExplainPlanButton.Click += (_, _) => ExplainCurrentPlan();
        }
        if (_quickEditPromptButton != null)
        {
            _quickEditPromptButton.Click += (_, _) => EditCurrentPrompt();
        }
        if (_quickHelpButton != null)
        {
            _quickHelpButton.Click += (_, _) => AppendSystem("Try commands like: arm, open notepad, run open edge and search weather, take snapshot.");
        }
        if (_planEditorLoadButton != null)
        {
            _planEditorLoadButton.Click += (_, _) => LoadCurrentPlanIntoEditor();
        }
        if (_planEditorValidateButton != null)
        {
            _planEditorValidateButton.Click += (_, _) => ValidateEditedPlan();
        }
        if (_planEditorDryRunButton != null)
        {
            _planEditorDryRunButton.Click += async (_, _) => await ExecuteEditedPlanAsync(dryRun: true);
        }
        if (_planEditorExecuteButton != null)
        {
            _planEditorExecuteButton.Click += async (_, _) => await ExecuteEditedPlanAsync(dryRun: false);
        }
        if (_planEditorRefreshHumanButton != null)
        {
            _planEditorRefreshHumanButton.Click += (_, _) => RefreshHumanPlanPreview();
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

                if (IsTimelineTabSelected())
                {
                    ResetUnreadTimelineEvents();
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
        if (_cfgForceApplyUpdateButton != null)
        {
            _cfgForceApplyUpdateButton.Click += (_, _) => ForceApplyUpdateFromConfig();
        }
        if (_cfgAudioRefreshButton != null)
        {
            _cfgAudioRefreshButton.Click += async (_, _) => await RefreshAudioInputsAsync();
        }
        if (_cfgOpenDataFolderButton != null)
        {
            _cfgOpenDataFolderButton.Click += (_, _) => OpenDataFolder();
        }
        if (_cfgAudioDeviceSelector != null)
        {
            _cfgAudioDeviceSelector.SelectionChanged += (_, _) =>
            {
                if (_cfgAudioDeviceSelector.SelectedItem is string selected &&
                    !string.IsNullOrWhiteSpace(selected) &&
                    !string.Equals(selected, "<auto>", StringComparison.OrdinalIgnoreCase))
                {
                    SetText(_cfgAudioDevice, selected);
                }
            };
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
        if (_macroRecordToggleButton != null)
        {
            _macroRecordToggleButton.Click += (_, _) => ToggleMacroRecording();
        }
        if (_macroClearButton != null)
        {
            _macroClearButton.Click += (_, _) => ClearMacroSteps();
        }
        if (_macroList != null)
        {
            _macroList.SelectionChanged += (_, _) => OnMacroSelectionChanged();
        }
        if (_macroApplyEditButton != null)
        {
            _macroApplyEditButton.Click += (_, _) => ApplyMacroStepEdit();
        }
        if (_macroRemoveStepButton != null)
        {
            _macroRemoveStepButton.Click += (_, _) => RemoveSelectedMacroStep();
        }
        if (_macroMoveUpButton != null)
        {
            _macroMoveUpButton.Click += (_, _) => MoveSelectedMacroStep(-1);
        }
        if (_macroMoveDownButton != null)
        {
            _macroMoveDownButton.Click += (_, _) => MoveSelectedMacroStep(1);
        }
        if (_macroAddWaitButton != null)
        {
            _macroAddWaitButton.Click += (_, _) => AddMacroWaitStep();
        }
        if (_macroSaveTaskButton != null)
        {
            _macroSaveTaskButton.Click += async (_, _) => await SaveMacroAsTaskAsync();
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
        if (_mediaRefreshButton != null)
        {
            _mediaRefreshButton.Click += async (_, _) => await LoadMediaAsync();
        }
        if (_mediaOpenFolderButton != null)
        {
            _mediaOpenFolderButton.Click += (_, _) => OpenMediaFolder();
        }
        if (_mediaOpenButton != null)
        {
            _mediaOpenButton.Click += (_, _) => OpenSelectedMedia();
        }
        if (_mediaCopyPathButton != null)
        {
            _mediaCopyPathButton.Click += async (_, _) => await CopySelectedMediaPathAsync();
        }
        if (_mediaDeleteButton != null)
        {
            _mediaDeleteButton.Click += async (_, _) => await DeleteSelectedMediaAsync();
        }
        if (_mediaList != null)
        {
            _mediaList.SelectionChanged += (_, _) => RefreshSelectedMediaPreview();
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

        if (_timelineCopyButton != null)
        {
            _timelineCopyButton.Click += async (_, _) => await CopyTextAsync(string.Join(Environment.NewLine, _timelineEntries), "Timeline copied.");
        }
        if (_timelineExportButton != null)
        {
            _timelineExportButton.Click += (_, _) => ExportTimeline();
        }
        if (_diagRefreshButton != null)
        {
            _diagRefreshButton.Click += async (_, _) => await RefreshDiagnosticsAsync();
        }
        if (_diagCopyButton != null)
        {
            _diagCopyButton.Click += async (_, _) => await CopyTextAsync(_diagBox?.Text, "Diagnostics copied.");
        }
        if (_wikiSearchBox != null)
        {
            _wikiSearchBox.PropertyChanged += (_, args) =>
            {
                if (args.Property.Name == nameof(TextBox.Text))
                {
                    ApplyWikiFilter();
                }
            };
        }
        if (_wikiResetSearchButton != null)
        {
            _wikiResetSearchButton.Click += (_, _) =>
            {
                SetText(_wikiSearchBox, string.Empty);
                ApplyWikiFilter();
            };
        }
        if (_wikiCopyButton != null)
        {
            _wikiCopyButton.Click += async (_, _) => await CopyTextAsync(_wikiContentBox?.Text, "Wiki copied.");
        }

        Opened += OnOpened;
        Activated += (_, _) =>
        {
            if (IsChatTabSelected())
            {
                ResetUnreadChatEvents();
            }
            if (IsTimelineTabSelected())
            {
                ResetUnreadTimelineEvents();
            }
        };
        Closed += (_, _) =>
        {
            _pollingCts.Cancel();
            _activeRequestCts?.Cancel();
            _activeRequestCts?.Dispose();
            _activeRequestCts = null;
            _busyAnimationCts?.Cancel();
            _busyAnimationCts?.Dispose();
            _busyAnimationCts = null;
            _streamingCts?.Cancel();
            _streamingCts?.Dispose();
            _streamingCts = null;
            _paletteDebounceCts?.Cancel();
            _paletteDebounceCts?.Dispose();
            _paletteDebounceCts = null;
            _persistSessionCts?.Cancel();
            _persistSessionCts?.Dispose();
            _persistSessionCts = null;
            PersistSessionState();
            DisposeMediaPreview();
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
        UpdateQuickActionState();
        UpdatePlanEditorState();
        await RefreshStatusAsync();
        await LoadConfigAsync();
        RefreshUpdateStatusInConfig();
        await LoadTasksAsync();
        await LoadSchedulesAsync();
        await LoadGoalsAsync();
        await LoadMediaAsync();
        await LoadAuditAsync();
        await RefreshDiagnosticsAsync();
        RefreshWikiContent();
        RefreshMacroRecorderUi();
        _ = PollStatusAsync(_pollingCts.Token);
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
        MaybeRecordMacroCommand(normalized);

        if (await TryHandleLocalTrayCommandAsync(normalized))
        {
            return;
        }

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
            await RenderResponseAsync(response, requestToken);
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

    private async Task<bool> TryHandleLocalTrayCommandAsync(string command)
    {
        static bool IsMatch(string text, params string[] candidates)
        {
            foreach (var candidate in candidates)
            {
                if (string.Equals(text, candidate, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        var normalized = command.Trim();
        if (!IsMatch(normalized,
                "force update now",
                "force update",
                "force apply update",
                "apply update force",
                "forza aggiornamento",
                "forza update"))
        {
            return false;
        }

        AppendUser(normalized);
        SetTimeline(new[] { "[..] Running local tray update command..." });
        SetBusy(true);

        try
        {
            AppendSystem("Local command: force update (check + apply).");
            if (_checkUpdatesNow != null)
            {
                AppendSystem("Checking updates...");
                await _checkUpdatesNow();
            }

            if (_forceApplyUpdateNow != null)
            {
                AppendSystem("Force applying downloaded update (kill blockers)...");
                _forceApplyUpdateNow();
                AppendSystem("Force apply requested.");
            }
            else
            {
                AppendSystem("Force apply update is not available.");
            }

            RefreshUpdateStatusInConfig();
            RefreshChatUpdateBadge();
        }
        catch (Exception ex)
        {
            AppendSystem($"Force update failed: {ex.Message}");
        }
        finally
        {
            SetBusy(false);
            await RefreshStatusAsync();
        }

        return true;
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
            await RenderResponseAsync(response, CancellationToken.None);
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

    private async Task RenderResponseAsync(WebChatResponse response, CancellationToken cancellationToken)
    {
        var reply = string.IsNullOrWhiteSpace(response.Reply) ? "<no reply>" : response.Reply;
        _lastReply = reply;
        await AppendAgentStreamingAsync(reply, cancellationToken);
        UpdateAiSuggestionState(reply);
        _lastPlanJson = string.IsNullOrWhiteSpace(response.PlanJson) ? _lastPlanJson : response.PlanJson;
        if (!string.IsNullOrWhiteSpace(response.PlanJson))
        {
            SetText(_planEditorBox, response.PlanJson!);
            RefreshHumanPlanPreview(response.PlanJson);
            SetText(_planEditorTitleText, "Plan preview (editable) - loaded from last response");
            SetText(_planEditorStatusText, "Plan loaded. You can validate, dry-run, or execute.");
        }

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

        UpdateQuickActionState();
        UpdatePlanEditorState();
    }

    private async Task RefreshStatusAsync()
    {
        try
        {
            var snapshot = await _apiClient.GetStatusAsync(CancellationToken.None);
            if (snapshot?.Adapter == null)
            {
                SetText(_statusText, "Adapter unavailable");
                SetConnectionBadgesUnknown();
                return;
            }

            var llmEnabled = snapshot.Llm?.Enabled == true;
            var llmAvailable = llmEnabled && snapshot.Llm?.Available == true;
            var killTripped = snapshot.KillSwitch?.Tripped == true;
            UpdateConnectionBadges(snapshot.Adapter.Armed, snapshot.Adapter.RequireUserPresence, llmEnabled, llmAvailable, killTripped);

            var statusLine = $"Adapter online · {DateTime.Now:HH:mm:ss}";
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
            SetConnectionBadgesUnknown();
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
        if (!IsTimelineTabSelected())
        {
            _unreadTimelineEvents = Math.Min(_unreadTimelineEvents + items.Count, 999);
            UpdateTimelineBadge();
        }
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
            if (_timelineStatusText != null)
            {
                var filterLabel = string.IsNullOrWhiteSpace(query) ? filter : $"{filter} + search";
                _timelineStatusText.Text = $"Showing {Math.Max(0, items.Count == 1 && items[0].Contains("No timeline.", StringComparison.OrdinalIgnoreCase) ? 0 : items.Count)} of {_timelineEntries.Count} ({filterLabel})";
            }
        });
    }

    private async Task ShowUpdateDetailsAsync()
    {
        var details = _getChatUpdateDetails?.Invoke()
            ?? new ChatUpdateDetails(false, "Updates", "No update information available.", false, false, string.Empty);

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

        var showBlockersButton = new Button
        {
            Content = "Show blockers",
            MinWidth = 110,
            IsVisible = details.HasBlockers && !string.IsNullOrWhiteSpace(details.Blockers)
        };
        var showingBlockers = false;
        showBlockersButton.Click += (_, _) =>
        {
            showingBlockers = !showingBlockers;
            contentBox.Text = showingBlockers
                ? details.Blockers
                : (string.IsNullOrWhiteSpace(details.Details) ? "No details available." : details.Details);
            showBlockersButton.Content = showingBlockers ? "Show details" : "Show blockers";
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
        if (showBlockersButton.IsVisible)
        {
            footer.Children.Add(showBlockersButton);
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
        UpdateQuickActionState();
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

    private void UpdateQuickActionState()
    {
        Dispatcher.UIThread.Post(() =>
        {
            var hasSuggestion = !string.IsNullOrWhiteSpace(_aiSuggestedCommand);
            var hasLastCommand = !string.IsNullOrWhiteSpace(_lastSentMessage);
            var hasPlan = !string.IsNullOrWhiteSpace(_lastPlanJson);

            if (_quickActionsPanel != null)
            {
                _quickActionsPanel.IsVisible = hasSuggestion || hasLastCommand || hasPlan;
            }

            if (_quickRunSuggestionButton != null)
            {
                _quickRunSuggestionButton.IsEnabled = !_busy && hasSuggestion;
            }

            if (_quickDryRunButton != null)
            {
                _quickDryRunButton.IsEnabled = !_busy && hasLastCommand;
            }

            if (_quickExplainPlanButton != null)
            {
                _quickExplainPlanButton.IsEnabled = hasPlan;
            }

            if (_quickEditPromptButton != null)
            {
                _quickEditPromptButton.IsEnabled = hasSuggestion || hasLastCommand;
            }
        });
    }

    private void UpdatePlanEditorState()
    {
        Dispatcher.UIThread.Post(() =>
        {
            var hasPlan = !string.IsNullOrWhiteSpace(_lastPlanJson);
            if (_planEditorPanel != null)
            {
                _planEditorPanel.IsVisible = hasPlan;
            }

            if (_planEditorLoadButton != null)
            {
                _planEditorLoadButton.IsEnabled = hasPlan;
            }

            if (_planEditorValidateButton != null)
            {
                _planEditorValidateButton.IsEnabled = !_busy;
            }

            if (_planEditorDryRunButton != null)
            {
                _planEditorDryRunButton.IsEnabled = !_busy && hasPlan;
            }

            if (_planEditorExecuteButton != null)
            {
                _planEditorExecuteButton.IsEnabled = !_busy && hasPlan;
            }

            if (_planEditorRefreshHumanButton != null)
            {
                _planEditorRefreshHumanButton.IsEnabled = !_busy && hasPlan;
            }
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
            var inputLength = input.Count(ch => !char.IsWhiteSpace(ch));
            if (inputLength < MinPaletteInputChars)
            {
                _suppressPaletteSelectionChanged = true;
                try
                {
                    _commandPaletteList.ItemsSource = Array.Empty<string>();
                    _commandPaletteList.SelectedIndex = -1;
                    _commandPaletteList.IsVisible = false;
                }
                finally
                {
                    _suppressPaletteSelectionChanged = false;
                }

                return;
            }

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

    private async Task RunDryRunLastCommandAsync()
    {
        if (_busy)
        {
            return;
        }

        var last = _lastSentMessage?.Trim();
        if (string.IsNullOrWhiteSpace(last))
        {
            AppendSystem("No previous command to dry-run.");
            return;
        }

        await SendMessageAsync($"dry-run {last}");
    }

    private void ExplainCurrentPlan()
    {
        if (string.IsNullOrWhiteSpace(_lastPlanJson))
        {
            AppendSystem("No plan available yet.");
            return;
        }

        LoadCurrentPlanIntoEditor();
        AppendSystem("Plan loaded in editor. You can validate, dry-run, or execute it.");
    }

    private void EditCurrentPrompt()
    {
        var prompt = _aiSuggestedCommand?.Trim();
        if (string.IsNullOrWhiteSpace(prompt))
        {
            prompt = _lastSentMessage?.Trim();
        }

        if (string.IsNullOrWhiteSpace(prompt))
        {
            return;
        }

        if (_inputBox != null)
        {
            _inputBox.Text = prompt;
            _inputBox.CaretIndex = _inputBox.Text?.Length ?? 0;
            _inputBox.Focus();
        }
    }

    private void LoadCurrentPlanIntoEditor()
    {
        if (string.IsNullOrWhiteSpace(_lastPlanJson))
        {
            SetText(_planEditorStatusText, "No plan to load.");
            return;
        }

        SetText(_planEditorBox, _lastPlanJson!);
        RefreshHumanPlanPreview(_lastPlanJson);
        SetText(_planEditorStatusText, "Plan loaded.");
        UpdatePlanEditorState();
    }

    private void ValidateEditedPlan()
    {
        var json = _planEditorBox?.Text?.Trim() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(json))
        {
            SetText(_planEditorStatusText, "Plan JSON is empty.");
            return;
        }

        if (!TryParsePlanJson(json, out var error))
        {
            SetText(_planEditorStatusText, $"Invalid plan: {error}");
            return;
        }

        _lastPlanJson = json;
        RefreshHumanPlanPreview(json);
        SetText(_planEditorStatusText, "Plan is valid.");
        UpdatePlanEditorState();
    }

    private async Task ExecuteEditedPlanAsync(bool dryRun)
    {
        if (_busy)
        {
            return;
        }

        var json = _planEditorBox?.Text?.Trim() ?? string.Empty;
        if (!TryParsePlanJson(json, out var error))
        {
            SetText(_planEditorStatusText, $"Invalid plan: {error}");
            return;
        }

        _lastPlanJson = json;
        RefreshHumanPlanPreview(json);
        SetText(_planEditorStatusText, dryRun ? "Running dry-run..." : "Executing plan...");

        SetBusy(true);
        AppendUser(dryRun ? "dry-run [edited plan]" : "run [edited plan]");
        SetTimeline(new[] { "[..] Waiting for response..." });
        try
        {
            var response = await Task.Run(
                () => _apiClient.ExecutePlanJsonAsync(json, dryRun, CancellationToken.None),
                CancellationToken.None);
            if (response == null)
            {
                AppendSystem("No response from plan execution.");
                return;
            }

            await RenderResponseAsync(ToChatResponse(response), CancellationToken.None);
            SetText(_planEditorStatusText, dryRun ? "Dry-run completed." : "Plan submitted.");
        }
        catch (Exception ex)
        {
            AppendSystem($"Error: {ex.Message}");
            SetText(_planEditorStatusText, $"Execution failed: {ex.Message}");
        }
        finally
        {
            SetBusy(false);
            await RefreshStatusAsync();
        }
    }

    private void RefreshHumanPlanPreview(string? rawJson = null)
    {
        var json = rawJson ?? _planEditorBox?.Text;
        var summary = BuildHumanPlanSummary(json, out var error);
        SetText(_planEditorHumanBox, summary);
        if (!string.IsNullOrWhiteSpace(error))
        {
            SetText(_planEditorStatusText, $"Human summary warning: {error}");
        }
    }

    private static string BuildHumanPlanSummary(string? json, out string error)
    {
        error = string.Empty;
        if (string.IsNullOrWhiteSpace(json))
        {
            return "No plan loaded.";
        }

        try
        {
            using var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;
            if (!TryGetJsonPropertyIgnoreCase(root, "steps", out var steps)
                || steps.ValueKind != JsonValueKind.Array
                || steps.GetArrayLength() == 0)
            {
                error = "Missing or empty steps.";
                return "No executable steps found.";
            }

            var lines = new List<string>();
            if (TryGetJsonPropertyIgnoreCase(root, "intent", out var intentElement))
            {
                var intent = JsonElementToCompactString(intentElement);
                if (!string.IsNullOrWhiteSpace(intent))
                {
                    lines.Add($"Intent: {intent}");
                    lines.Add(string.Empty);
                }
            }

            var index = 1;
            foreach (var step in steps.EnumerateArray())
            {
                var typeLabel = TryGetJsonPropertyIgnoreCase(step, "type", out var typeElement)
                    ? NormalizeActionTypeLabel(typeElement)
                    : "Unknown";
                var details = new List<string>();

                if (TryGetJsonPropertyIgnoreCase(step, "appidorpath", out var appElement))
                {
                    var app = JsonElementToCompactString(appElement);
                    if (!string.IsNullOrWhiteSpace(app))
                    {
                        details.Add($"app={app}");
                    }
                }

                if (TryGetJsonPropertyIgnoreCase(step, "target", out var targetElement))
                {
                    var target = JsonElementToCompactString(targetElement);
                    if (!string.IsNullOrWhiteSpace(target))
                    {
                        details.Add($"target={target}");
                    }
                }

                if (TryGetJsonPropertyIgnoreCase(step, "text", out var textElement))
                {
                    var text = JsonElementToCompactString(textElement);
                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        details.Add($"text={text}");
                    }
                }

                if (TryGetJsonPropertyIgnoreCase(step, "keys", out var keysElement))
                {
                    var keys = JsonElementToCompactString(keysElement);
                    if (!string.IsNullOrWhiteSpace(keys))
                    {
                        details.Add($"keys={keys}");
                    }
                }

                if (TryGetJsonPropertyIgnoreCase(step, "waitfor", out var waitElement))
                {
                    var wait = JsonElementToCompactString(waitElement);
                    if (!string.IsNullOrWhiteSpace(wait))
                    {
                        details.Add($"wait={wait}");
                    }
                }

                if (TryGetJsonPropertyIgnoreCase(step, "expectedappid", out var expectedAppElement))
                {
                    var expectedApp = JsonElementToCompactString(expectedAppElement);
                    if (!string.IsNullOrWhiteSpace(expectedApp))
                    {
                        details.Add($"expectedApp={expectedApp}");
                    }
                }

                if (TryGetJsonPropertyIgnoreCase(step, "expectedwindowid", out var expectedWindowElement))
                {
                    var expectedWindow = JsonElementToCompactString(expectedWindowElement);
                    if (!string.IsNullOrWhiteSpace(expectedWindow))
                    {
                        details.Add($"expectedWindow={expectedWindow}");
                    }
                }

                lines.Add(details.Count == 0
                    ? $"{index}. {typeLabel}"
                    : $"{index}. {typeLabel} ({string.Join(", ", details)})");
                index++;
            }

            lines.Add(string.Empty);
            lines.Add($"Total steps: {steps.GetArrayLength()}");
            return string.Join(Environment.NewLine, lines);
        }
        catch (Exception ex)
        {
            error = ex.Message;
            return "Unable to build human summary from plan JSON.";
        }
    }

    private static string NormalizeActionTypeLabel(JsonElement value)
    {
        if (value.ValueKind == JsonValueKind.Number && value.TryGetInt32(out var intValue))
        {
            if (Enum.IsDefined(typeof(DesktopAgent.Core.Models.ActionType), intValue))
            {
                return ((DesktopAgent.Core.Models.ActionType)intValue).ToString();
            }

            return $"Action#{intValue}";
        }

        var raw = JsonElementToCompactString(value);
        if (string.IsNullOrWhiteSpace(raw))
        {
            return "Unknown";
        }

        if (Enum.TryParse<DesktopAgent.Core.Models.ActionType>(raw, ignoreCase: true, out var parsed))
        {
            return parsed.ToString();
        }

        return raw;
    }

    private static string JsonElementToCompactString(JsonElement value)
    {
        return value.ValueKind switch
        {
            JsonValueKind.String => value.GetString()?.Trim() ?? string.Empty,
            JsonValueKind.Null => string.Empty,
            JsonValueKind.Undefined => string.Empty,
            JsonValueKind.Array => string.Join("+", value.EnumerateArray().Select(JsonElementToCompactString).Where(v => !string.IsNullOrWhiteSpace(v))),
            JsonValueKind.Object => value.GetRawText(),
            _ => value.GetRawText()
        };
    }

    private static bool TryParsePlanJson(string planJson, out string error)
    {
        error = string.Empty;
        try
        {
            using var doc = JsonDocument.Parse(planJson);
            if (!TryGetJsonPropertyIgnoreCase(doc.RootElement, "steps", out var steps)
                || steps.ValueKind != JsonValueKind.Array
                || steps.GetArrayLength() == 0)
            {
                error = "Missing or empty 'steps' array.";
                return false;
            }

            return true;
        }
        catch (Exception ex)
        {
            error = ex.Message;
            return false;
        }
    }

    private static bool TryGetJsonPropertyIgnoreCase(JsonElement element, string name, out JsonElement value)
    {
        value = default;
        if (element.ValueKind != JsonValueKind.Object)
        {
            return false;
        }

        foreach (var property in element.EnumerateObject())
        {
            if (string.Equals(property.Name, name, StringComparison.OrdinalIgnoreCase))
            {
                value = property.Value;
                return true;
            }
        }

        return false;
    }

    private static WebChatResponse ToChatResponse(WebIntentResponse response)
    {
        return new WebChatResponse(
            response.Reply,
            response.NeedsConfirmation,
            response.Token,
            response.ActionLabel,
            response.Steps,
            response.PlanJson,
            response.ModeLabel);
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
            SelectLlmParsingMode(config.LlmInterpretationMode ?? "primary");
            _configuredMediaOutputDirectory = string.IsNullOrWhiteSpace(config.MediaOutputDirectory)
                ? "media"
                : config.MediaOutputDirectory.Trim();
            SetText(_cfgMediaOutputDirectory, _configuredMediaOutputDirectory);
            SelectAudioBackend(config.ScreenRecordingAudioBackendPreference ?? "auto");
            SetText(_cfgAudioDevice, config.ScreenRecordingAudioDevice ?? string.Empty);
            if (_cfgPrimaryDisplayOnly != null)
            {
                _cfgPrimaryDisplayOnly.IsChecked = config.ScreenRecordingPrimaryDisplayOnly;
            }

            SelectProvider(config.Llm.Provider ?? "ollama");
            await RefreshAudioInputsAsync();
            AppendConfigStatus("Config loaded.");
            RefreshUpdateStatusInConfig();
            RefreshWikiContent();
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
                MediaOutputDirectory: _cfgMediaOutputDirectory?.Text?.Trim(),
                ScreenRecordingAudioBackendPreference: GetSelectedAudioBackend(),
                ScreenRecordingAudioDevice: _cfgAudioDevice?.Text?.Trim(),
                ScreenRecordingPrimaryDisplayOnly: _cfgPrimaryDisplayOnly?.IsChecked ?? false,
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
                AuditLlmIncludeRawText: null,
                LlmInterpretationMode: GetSelectedLlmParsingMode());

            var response = await _apiClient.SaveConfigAsync(update, CancellationToken.None);
            if (response == null)
            {
                AppendConfigStatus("Config save failed.");
                return;
            }

            if (response.MediaOutputDirectory != null)
            {
                _configuredMediaOutputDirectory = string.IsNullOrWhiteSpace(response.MediaOutputDirectory)
                    ? "media"
                    : response.MediaOutputDirectory.Trim();
                SetText(_cfgMediaOutputDirectory, _configuredMediaOutputDirectory);
                await LoadMediaAsync();
            }

            AppendConfigStatus("Config saved.");
            await TestLlmAsync();
        }
        catch (Exception ex)
        {
            AppendConfigStatus($"Config save failed: {ex.Message}");
        }
    }

    private async Task RefreshAudioInputsAsync()
    {
        if (_cfgAudioDeviceSelector == null)
        {
            return;
        }

        try
        {
            var ffmpegEnvPath = (Environment.GetEnvironmentVariable("DESKTOP_AGENT_FFMPEG_PATH") ?? string.Empty).Trim();
            var ffmpegPath = string.IsNullOrWhiteSpace(ffmpegEnvPath)
                ? await Task.Run(() => ResolveCommandPath("ffmpeg"), CancellationToken.None)
                : ffmpegEnvPath;
            if (string.IsNullOrWhiteSpace(ffmpegPath))
            {
                _cfgAudioDeviceSelector.ItemsSource = new[] { "<auto>" };
                _cfgAudioDeviceSelector.SelectedIndex = 0;
                AppendConfigStatus("Audio inputs: ffmpeg not found.");
                return;
            }

            List<string> devices;
            if (OperatingSystem.IsWindows())
            {
                var dshowList = await Task.Run(() => RunProcessCapture(ffmpegPath, "-hide_banner -f dshow -list_devices true -i dummy", 9000), CancellationToken.None);
                devices = ParseDirectShowAudioDevices(dshowList.StdErr)
                    .Where(static x => !string.IsNullOrWhiteSpace(x))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();
            }
            else if (OperatingSystem.IsMacOS())
            {
                var avList = await Task.Run(() => RunProcessCapture(ffmpegPath, "-hide_banner -f avfoundation -list_devices true -i \"\"", 9000), CancellationToken.None);
                devices = ParseAvFoundationAudioDevices(avList.StdErr)
                    .Where(static x => !string.IsNullOrWhiteSpace(x))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();
            }
            else if (OperatingSystem.IsLinux())
            {
                var pulseList = await Task.Run(() => RunProcessCapture("pactl", "list short sources", 6000), CancellationToken.None);
                devices = ParsePulseAudioSources(pulseList.StdOut)
                    .Where(static x => !string.IsNullOrWhiteSpace(x))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();
            }
            else
            {
                devices = new List<string>();
            }

            var items = new List<string> { "<auto>" };
            items.AddRange(devices);
            _cfgAudioDeviceSelector.ItemsSource = items;

            var current = (_cfgAudioDevice?.Text ?? string.Empty).Trim();
            if (!string.IsNullOrWhiteSpace(current))
            {
                var match = items.FindIndex(x => string.Equals(x, current, StringComparison.OrdinalIgnoreCase));
                _cfgAudioDeviceSelector.SelectedIndex = match >= 0 ? match : 0;
            }
            else
            {
                _cfgAudioDeviceSelector.SelectedIndex = 0;
            }

            AppendConfigStatus($"Audio inputs loaded: {devices.Count}");
        }
        catch (Exception ex)
        {
            AppendConfigStatus($"Audio inputs refresh failed: {ex.Message}");
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

    private async void ForceApplyUpdateFromConfig()
    {
        try
        {
            AppendConfigStatus("Force apply update requested...");
            if (_checkUpdatesNow != null)
            {
                await _checkUpdatesNow();
            }

            if (_forceApplyUpdateNow != null)
            {
                _forceApplyUpdateNow();
                AppendConfigStatus("Force apply requested.");
            }
            else
            {
                AppendConfigStatus("Force apply update is not available.");
            }

            RefreshUpdateStatusInConfig();
            RefreshChatUpdateBadge();
        }
        catch (Exception ex)
        {
            AppendConfigStatus($"Force apply update failed: {ex.Message}");
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

        if (_cfgForceApplyUpdateButton != null)
        {
            var canApply = _getChatUpdateBadge?.Invoke().CanApply ?? false;
            _cfgForceApplyUpdateButton.IsEnabled = canApply && !_busy;
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

    private void ToggleMacroRecording()
    {
        _macroRecording = !_macroRecording;
        RefreshMacroRecorderUi();
        var message = _macroRecording ? "Macro recording started." : "Macro recording stopped.";
        SetText(_macroStatusText, message);
        AppendSystem(message);
    }

    private void ClearMacroSteps()
    {
        _macroSteps.Clear();
        if (_macroList != null)
        {
            _macroList.SelectedIndex = -1;
        }

        RefreshMacroRecorderUi();
        SetText(_macroStatusText, "Macro cleared.");
    }

    private void OnMacroSelectionChanged()
    {
        var index = _macroList?.SelectedIndex ?? -1;
        if (index < 0 || index >= _macroSteps.Count)
        {
            SetText(_macroStepEditorInput, string.Empty);
            return;
        }

        var step = _macroSteps[index];
        SetText(_macroStepEditorInput, step.Value);
    }

    private void ApplyMacroStepEdit()
    {
        var index = _macroList?.SelectedIndex ?? -1;
        if (index < 0 || index >= _macroSteps.Count)
        {
            SetText(_macroStatusText, "Select a macro step.");
            return;
        }

        var edited = (_macroStepEditorInput?.Text ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(edited))
        {
            SetText(_macroStatusText, "Step cannot be empty.");
            return;
        }

        var current = _macroSteps[index];
        if (string.Equals(current.Kind, "wait", StringComparison.OrdinalIgnoreCase))
        {
            if (!int.TryParse(edited, NumberStyles.Integer, CultureInfo.InvariantCulture, out var seconds) || seconds <= 0)
            {
                SetText(_macroStatusText, "Wait step must be a positive number (seconds).");
                return;
            }

            _macroSteps[index] = new MacroStep("wait", seconds.ToString(CultureInfo.InvariantCulture));
        }
        else
        {
            _macroSteps[index] = new MacroStep("command", edited);
        }

        RefreshMacroRecorderUi(index);
        SetText(_macroStatusText, "Step updated.");
    }

    private void RemoveSelectedMacroStep()
    {
        var index = _macroList?.SelectedIndex ?? -1;
        if (index < 0 || index >= _macroSteps.Count)
        {
            SetText(_macroStatusText, "Select a macro step.");
            return;
        }

        _macroSteps.RemoveAt(index);
        RefreshMacroRecorderUi(Math.Clamp(index - 1, 0, _macroSteps.Count - 1));
        SetText(_macroStatusText, "Step removed.");
    }

    private void MoveSelectedMacroStep(int delta)
    {
        var index = _macroList?.SelectedIndex ?? -1;
        if (index < 0 || index >= _macroSteps.Count)
        {
            SetText(_macroStatusText, "Select a macro step.");
            return;
        }

        var target = index + delta;
        if (target < 0 || target >= _macroSteps.Count)
        {
            return;
        }

        (_macroSteps[index], _macroSteps[target]) = (_macroSteps[target], _macroSteps[index]);
        RefreshMacroRecorderUi(target);
    }

    private void AddMacroWaitStep()
    {
        var raw = (_macroAddWaitSecondsInput?.Text ?? string.Empty).Trim();
        if (!int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out var seconds) || seconds <= 0)
        {
            SetText(_macroStatusText, "Wait seconds must be a positive integer.");
            return;
        }

        _macroSteps.Add(new MacroStep("wait", seconds.ToString(CultureInfo.InvariantCulture)));
        RefreshMacroRecorderUi(_macroSteps.Count - 1);
        SetText(_macroStatusText, $"Wait step added ({seconds}s).");
    }

    private async Task SaveMacroAsTaskAsync()
    {
        var name = (_macroNameInput?.Text ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(name))
        {
            SetText(_macroStatusText, "Macro name is required.");
            return;
        }

        if (!TryBuildMacroIntent(out var intent))
        {
            SetText(_macroStatusText, "Macro has no steps.");
            return;
        }

        var description = (_macroDescriptionInput?.Text ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(description))
        {
            description = "Recorded macro";
        }

        try
        {
            var response = await _apiClient.SaveTaskAsync(
                new WebTaskUpsertRequest(name, intent, description, null),
                CancellationToken.None);

            SetText(_macroStatusText, response?.Message ?? "Macro saved as task.");
            SetText(_tasksStatusText, response?.Message ?? "Task saved.");
            SetText(_taskNameInput, name);
            SetText(_taskIntentInput, intent);
            SetText(_taskDescriptionInput, description);
            await LoadTasksAsync();
        }
        catch (Exception ex)
        {
            SetText(_macroStatusText, $"Macro save failed: {ex.Message}");
        }
    }

    private void MaybeRecordMacroCommand(string command)
    {
        if (!_macroRecording)
        {
            return;
        }

        var normalized = command.Trim();
        if (string.IsNullOrWhiteSpace(normalized))
        {
            return;
        }

        _macroSteps.Add(new MacroStep("command", normalized));
        RefreshMacroRecorderUi(_macroSteps.Count - 1);
    }

    private bool TryBuildMacroIntent(out string intent)
    {
        intent = string.Empty;
        if (_macroSteps.Count == 0)
        {
            return false;
        }

        var parts = new List<string>();
        foreach (var step in _macroSteps)
        {
            if (string.Equals(step.Kind, "wait", StringComparison.OrdinalIgnoreCase))
            {
                if (!int.TryParse(step.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var seconds) || seconds <= 0)
                {
                    continue;
                }

                parts.Add($"wait for {seconds} seconds");
                continue;
            }

            if (!string.IsNullOrWhiteSpace(step.Value))
            {
                parts.Add(step.Value.Trim());
            }
        }

        if (parts.Count == 0)
        {
            return false;
        }

        intent = string.Join(" and then ", parts);
        return true;
    }

    private void RefreshMacroRecorderUi(int selectedIndex = -1)
    {
        if (_macroRecordToggleButton != null)
        {
            _macroRecordToggleButton.Content = _macroRecording ? "Stop Recording" : "Start Recording";
        }

        if (_macroList != null)
        {
            _macroList.ItemsSource = _macroSteps
                .Select((step, index) => FormatMacroStep(index, step))
                .ToList();

            if (_macroSteps.Count == 0)
            {
                _macroList.SelectedIndex = -1;
            }
            else if (selectedIndex >= 0 && selectedIndex < _macroSteps.Count)
            {
                _macroList.SelectedIndex = selectedIndex;
            }
        }

        if (_macroNameInput != null && string.IsNullOrWhiteSpace(_macroNameInput.Text))
        {
            _macroNameInput.Text = $"macro-{DateTime.Now:HHmmss}";
        }
    }

    private Task LoadMediaAsync()
    {
        try
        {
            var root = ResolveMediaRoot();
            Directory.CreateDirectory(root);

            var supported = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                ".png", ".jpg", ".jpeg", ".bmp", ".gif", ".webp",
                ".mp4", ".mkv", ".mov", ".avi", ".webm"
            };

            _mediaItems.Clear();
            foreach (var file in Directory.EnumerateFiles(root))
            {
                var ext = Path.GetExtension(file);
                if (string.IsNullOrWhiteSpace(ext) || !supported.Contains(ext))
                {
                    continue;
                }

                FileInfo info;
                try
                {
                    info = new FileInfo(file);
                }
                catch
                {
                    continue;
                }

                var isImage = IsImageExtension(ext);
                var isVideo = IsVideoExtension(ext);
                _mediaItems.Add(new MediaFileItem(info.FullName, info.Name, info.Length, info.LastWriteTimeUtc, ext, isImage, isVideo));
            }

            _mediaItems.Sort((a, b) => b.LastWriteUtc.CompareTo(a.LastWriteUtc));

            if (_mediaList != null)
            {
                _mediaList.ItemsSource = _mediaItems.Select(FormatMedia).ToList();
            }

            SetText(_mediaStatusText, $"Media: {_mediaItems.Count} files");
            if (_mediaItems.Count == 0)
            {
                DisposeMediaPreview();
                SetText(_mediaPreviewTitleText, "No media found");
                SetText(_mediaPreviewInfoText, $"Folder: {root}");
            }
            else
            {
                if (_mediaList != null)
                {
                    _mediaList.SelectedIndex = 0;
                }
                RefreshSelectedMediaPreview();
            }
        }
        catch (Exception ex)
        {
            SetText(_mediaStatusText, $"Media load failed: {ex.Message}");
        }

        return Task.CompletedTask;
    }

    private string ResolveMediaRoot()
    {
        var root = ResolveDataRoot();
        var configured = (_configuredMediaOutputDirectory ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(configured))
        {
            configured = "media";
        }

        var mediaPath = Path.IsPathRooted(configured)
            ? configured
            : Path.Combine(root, configured);
        return Path.GetFullPath(mediaPath);
    }

    private void OpenMediaFolder()
    {
        try
        {
            var root = ResolveMediaRoot();
            Directory.CreateDirectory(root);
            Process.Start(new ProcessStartInfo
            {
                FileName = root,
                UseShellExecute = true
            });
        }
        catch (Exception ex)
        {
            SetText(_mediaStatusText, $"Open folder failed: {ex.Message}");
        }
    }

    private void OpenSelectedMedia()
    {
        var selected = SelectedMedia();
        if (selected == null)
        {
            SetText(_mediaStatusText, "Select a media file.");
            return;
        }

        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = selected.Path,
                UseShellExecute = true
            });
        }
        catch (Exception ex)
        {
            SetText(_mediaStatusText, $"Open media failed: {ex.Message}");
        }
    }

    private async Task CopySelectedMediaPathAsync()
    {
        var selected = SelectedMedia();
        if (selected == null)
        {
            SetText(_mediaStatusText, "Select a media file.");
            return;
        }

        await CopyTextAsync(selected.Path, "Media path copied.");
    }

    private async Task DeleteSelectedMediaAsync()
    {
        var selected = SelectedMedia();
        if (selected == null)
        {
            SetText(_mediaStatusText, "Select a media file.");
            return;
        }

        try
        {
            File.Delete(selected.Path);
            SetText(_mediaStatusText, $"Deleted: {selected.Name}");
            await LoadMediaAsync();
        }
        catch (Exception ex)
        {
            SetText(_mediaStatusText, $"Delete failed: {ex.Message}");
        }
    }

    private void RefreshSelectedMediaPreview()
    {
        var selected = SelectedMedia();
        if (selected == null)
        {
            DisposeMediaPreview();
            SetText(_mediaPreviewTitleText, "No media selected");
            SetText(_mediaPreviewInfoText, string.Empty);
            return;
        }

        SetText(_mediaPreviewTitleText, selected.Name);
        var kind = selected.IsImage ? "image" : selected.IsVideo ? "video" : "file";
        var info = $"{kind.ToUpperInvariant()}{Environment.NewLine}" +
                   $"Size: {FormatBytes(selected.SizeBytes)}{Environment.NewLine}" +
                   $"Modified: {selected.LastWriteUtc.ToLocalTime():yyyy-MM-dd HH:mm:ss}{Environment.NewLine}" +
                   $"Path: {selected.Path}";
        SetText(_mediaPreviewInfoText, info);

        if (!selected.IsImage || !File.Exists(selected.Path))
        {
            DisposeMediaPreview();
            return;
        }

        try
        {
            DisposeMediaPreview();
            _mediaPreviewBitmap = new Bitmap(selected.Path);
            if (_mediaPreviewImage != null)
            {
                _mediaPreviewImage.Source = _mediaPreviewBitmap;
            }
        }
        catch
        {
            DisposeMediaPreview();
        }
    }

    private void DisposeMediaPreview()
    {
        try
        {
            if (_mediaPreviewImage != null)
            {
                _mediaPreviewImage.Source = null;
            }

            _mediaPreviewBitmap?.Dispose();
            _mediaPreviewBitmap = null;
        }
        catch
        {
            // ignored
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

    private MediaFileItem? SelectedMedia()
    {
        var index = _mediaList?.SelectedIndex ?? -1;
        if (index < 0 || index >= _mediaItems.Count)
        {
            return null;
        }

        return _mediaItems[index];
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
        if (busy)
        {
            _busyStartedAtUtc = DateTimeOffset.UtcNow;
            _busyAnimationCts?.Cancel();
            _busyAnimationCts?.Dispose();
            _busyAnimationCts = new CancellationTokenSource();
            StartBusyAnimation(_busyAnimationCts.Token);
        }
        else
        {
            _busyAnimationCts?.Cancel();
            _busyAnimationCts?.Dispose();
            _busyAnimationCts = null;

            if (_busyStartedAtUtc != DateTimeOffset.MinValue)
            {
                var elapsed = (DateTimeOffset.UtcNow - _busyStartedAtUtc).TotalSeconds;
                if (elapsed is > 0.1 and < 120)
                {
                    _busyAvgSeconds = (_busyAvgSeconds * 0.7) + (elapsed * 0.3);
                }
            }
        }

        _busy = busy;
        Dispatcher.UIThread.Post(() =>
        {
            if (_busyPanel != null)
            {
                _busyPanel.IsVisible = busy;
            }
            if (_busyText != null)
            {
                _busyText.Text = busy
                    ? $"Thinking... 0.0s elapsed | ETA ~{Math.Max(1, (int)Math.Round(_busyAvgSeconds, MidpointRounding.AwayFromZero))}s"
                    : string.Empty;
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
        UpdateQuickActionState();
        UpdatePlanEditorState();
        RefreshChatUpdateBadge();
        RefreshUpdateStatusInConfig();
    }

    private void StartBusyAnimation(CancellationToken cancellationToken)
    {
        _ = Task.Run(async () =>
        {
            var frame = 0;
            while (!cancellationToken.IsCancellationRequested)
            {
                var elapsed = DateTimeOffset.UtcNow - _busyStartedAtUtc;
                var etaSeconds = Math.Max(0, _busyAvgSeconds - elapsed.TotalSeconds);
                var etaLabel = etaSeconds < 1
                    ? "<1s"
                    : $"{Math.Ceiling(etaSeconds):0}s";
                SetText(_busyText, $"Thinking... {elapsed.TotalSeconds:0.0}s elapsed | ETA ~{etaLabel}");
                frame++;
                try
                {
                    await Task.Delay(250, cancellationToken);
                }
                catch (OperationCanceledException)
                {
                    break;
                }
            }
        }, CancellationToken.None);
    }

    private IEnumerable<Button> AllButtons()
    {
        return new[]
        {
            _sendButton, _confirmButton, _cancelButton, _statusButton, _armButton, _disarmButton, _simPresenceButton,
            _useSuggestionButton, _chatUpdateDetailsButton, _chatApplyUpdateButton, _cancelRequestButton,
            _quickRunSuggestionButton, _quickDryRunButton, _quickExplainPlanButton, _quickEditPromptButton, _quickHelpButton,
            _planEditorLoadButton, _planEditorValidateButton, _planEditorDryRunButton, _planEditorExecuteButton, _planEditorRefreshHumanButton,
            _reqPresenceButton, _killButton, _resetKillButton, _restartAdapterButton, _restartServerButton,
            _lockWindowButton, _lockAppButton, _unlockButton, _profileSafeButton,
            _profileBalancedButton, _profilePowerButton, _copyButton, _clearButton,
            _cfgLoadButton, _cfgSaveButton, _cfgTestLlmButton, _cfgRunFirstSetupButton, _cfgCheckUpdatesButton, _cfgApplyUpdateButton, _cfgForceApplyUpdateButton, _cfgAudioRefreshButton, _cfgOpenDataFolderButton, _tasksRefreshButton, _tasksRunButton,
            _tasksDeleteButton, _taskSaveButton, _macroRecordToggleButton, _macroClearButton, _macroApplyEditButton, _macroRemoveStepButton,
            _macroMoveUpButton, _macroMoveDownButton, _macroAddWaitButton, _macroSaveTaskButton, _schedulesRefreshButton, _schedulesRunButton, _schedulesDeleteButton,
            _scheduleSaveButton, _goalsRefreshButton, _goalsToggleAutoButton, _goalsDoneButton,
            _goalsRemoveButton, _goalAddButton, _mediaRefreshButton, _mediaOpenFolderButton, _mediaOpenButton, _mediaCopyPathButton, _mediaDeleteButton,
            _auditRefreshButton, _auditCopyButton, _auditClearButton
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

    private bool IsTimelineTabSelected()
    {
        if (_mainTabs == null || _timelineTab == null)
        {
            return false;
        }

        return ReferenceEquals(_mainTabs.SelectedItem, _timelineTab);
    }

    private void ResetUnreadChatEvents()
    {
        _unreadChatEvents = 0;
        _unreadErrorEvents = 0;
        _unreadInfoEvents = 0;
        UpdateUnreadBadge();
    }

    private void ResetUnreadTimelineEvents()
    {
        _unreadTimelineEvents = 0;
        UpdateTimelineBadge();
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

    private void UpdateTimelineBadge()
    {
        Dispatcher.UIThread.Post(() =>
        {
            if (_timelineTabBadgePanel == null || _timelineTabBadgeText == null)
            {
                return;
            }

            var hasUnread = _unreadTimelineEvents > 0;
            _timelineTabBadgePanel.IsVisible = hasUnread;
            if (!hasUnread)
            {
                ToolTip.SetTip(_timelineTabBadgePanel, null);
                return;
            }

            _timelineTabBadgeText.Text = _unreadTimelineEvents > 99 ? "99+" : _unreadTimelineEvents.ToString(CultureInfo.InvariantCulture);
            ToolTip.SetTip(_timelineTabBadgePanel, $"New timeline events: {_unreadTimelineEvents}");
        });
    }

    private void ExportTimeline()
    {
        try
        {
            var root = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "DesktopAgent",
                "exports");
            Directory.CreateDirectory(root);
            var file = Path.Combine(root, $"timeline-{DateTimeOffset.Now:yyyyMMdd-HHmmss}.txt");
            File.WriteAllLines(file, _timelineEntries);
            SetText(_timelineStatusText, $"Timeline exported: {file}");
            AppendSystem($"Timeline exported: {file}");
        }
        catch (Exception ex)
        {
            SetText(_timelineStatusText, $"Timeline export failed: {ex.Message}");
        }
    }

    private async Task RefreshDiagnosticsAsync()
    {
        try
        {
            var lines = new List<string>
            {
                $"Timestamp: {DateTimeOffset.Now:yyyy-MM-dd HH:mm:ss zzz}",
                $"Process: {Environment.ProcessPath}",
                $"OS: {RuntimeInformation.OSDescription}",
                $"Framework: {RuntimeInformation.FrameworkDescription}",
                string.Empty
            };

            var status = await _apiClient.GetStatusAsync(CancellationToken.None);
            if (status?.Adapter != null)
            {
                lines.Add($"Adapter: armed={status.Adapter.Armed}, requirePresence={status.Adapter.RequireUserPresence}, message={status.Adapter.Message}");
            }
            else
            {
                lines.Add("Adapter: unavailable");
            }

            if (status?.Llm != null)
            {
                lines.Add($"LLM: enabled={status.Llm.Enabled}, available={status.Llm.Available}, provider={status.Llm.Provider}, endpoint={status.Llm.Endpoint}");
            }

            var ffmpegPath = await Task.Run(() => ResolveCommandPath("ffmpeg"), CancellationToken.None);
            var tesseractPath = await Task.Run(() => ResolveCommandPath("tesseract"), CancellationToken.None);
            var ffmpegEnvPath = (Environment.GetEnvironmentVariable("DESKTOP_AGENT_FFMPEG_PATH") ?? string.Empty).Trim();
            var ffmpegFound = !string.IsNullOrWhiteSpace(ffmpegPath);
            var tesseractFound = !string.IsNullOrWhiteSpace(tesseractPath);
            var hasWasapi = false;
            var hasDshow = false;
            var hasAvfoundation = false;
            var hasPipewire = false;
            var hasPulse = false;
            var hasAlsa = false;
            var dshowAudioDevices = new List<string>();

            lines.Add($"ffmpeg env override: {(string.IsNullOrWhiteSpace(ffmpegEnvPath) ? "<none>" : ffmpegEnvPath)}");
            lines.Add($"ffmpeg: {(ffmpegFound ? $"found ({ffmpegPath})" : "missing")}");
            lines.Add($"tesseract: {(tesseractFound ? $"found ({tesseractPath})" : "missing")}");

            if (ffmpegFound)
            {
                var devices = await Task.Run(() => RunProcessCapture(ffmpegPath!, "-hide_banner -devices", 7000), CancellationToken.None);
                var devicesText = string.Join(Environment.NewLine, new[] { devices.StdOut, devices.StdErr });
                hasWasapi = Regex.IsMatch(devicesText, @"\bwasapi\b", RegexOptions.IgnoreCase);
                hasDshow = Regex.IsMatch(devicesText, @"\bdshow\b", RegexOptions.IgnoreCase);
                hasAvfoundation = Regex.IsMatch(devicesText, @"\bavfoundation\b", RegexOptions.IgnoreCase);
                hasPipewire = Regex.IsMatch(devicesText, @"\bpipewire\b", RegexOptions.IgnoreCase);
                hasPulse = Regex.IsMatch(devicesText, @"\bpulse\b", RegexOptions.IgnoreCase);
                hasAlsa = Regex.IsMatch(devicesText, @"\balsa\b", RegexOptions.IgnoreCase);

                var backendLabels = new List<string>();
                if (hasWasapi) backendLabels.Add("WASAPI");
                if (hasDshow) backendLabels.Add("DirectShow");
                if (hasAvfoundation) backendLabels.Add("AVFoundation");
                if (hasPipewire) backendLabels.Add("PipeWire");
                if (hasPulse) backendLabels.Add("PulseAudio");
                if (hasAlsa) backendLabels.Add("ALSA");
                lines.Add($"ffmpeg audio backends: {(backendLabels.Count == 0 ? "none detected" : string.Join(", ", backendLabels))}");

                var version = await Task.Run(() => RunProcessCapture(ffmpegPath!, "-hide_banner -version", 6000), CancellationToken.None);
                var versionLine = version.StdOut
                    .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                    .FirstOrDefault() ?? "<unknown>";
                lines.Add($"ffmpeg version: {versionLine}");

                if (hasDshow)
                {
                    var dshowList = await Task.Run(() => RunProcessCapture(ffmpegPath!, "-hide_banner -f dshow -list_devices true -i dummy", 9000), CancellationToken.None);
                    dshowAudioDevices = ParseDirectShowAudioDevices(dshowList.StdErr).ToList();
                    lines.Add($"dshow audio devices: {(dshowAudioDevices.Count == 0 ? "none" : string.Join(" | ", dshowAudioDevices.Take(5)))}");
                }
            }

            lines.Add(string.Empty);

            var startupPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "DesktopAgent", "tray-startup.log");
            var velopackPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "DesktopAgent", "velopack.log");
            lines.Add($"startup log: {startupPath}");
            lines.Add($"velopack log: {velopackPath}");
            lines.Add(string.Empty);
            lines.Add("--- startup tail ---");
            lines.AddRange(ReadFileTail(startupPath, 40));
            lines.Add(string.Empty);
            lines.Add("--- velopack tail ---");
            lines.AddRange(ReadFileTail(velopackPath, 60));

            SetText(_diagBox, string.Join(Environment.NewLine, lines));

            var issues = new List<string>();
            var hasAnyAudioBackend = hasWasapi || hasDshow || hasAvfoundation || hasPipewire || hasPulse || hasAlsa;
            if (!ffmpegFound)
            {
                issues.Add("Install FFmpeg (Plugin setup).");
            }
            else if (!hasAnyAudioBackend)
            {
                issues.Add("FFmpeg build has no supported audio input backend.");
            }
            else if (OperatingSystem.IsWindows() && hasDshow && dshowAudioDevices.Count == 0)
            {
                issues.Add("No DirectShow audio devices found (check microphone privacy + default input device).");
            }
            if (status?.Adapter == null)
            {
                issues.Add("Adapter offline. Use Restart Adapter.");
            }
            if (issues.Count == 0)
            {
                SetText(_diagStatusText, "Environment check: OK");
            }
            else
            {
                SetText(_diagStatusText, $"Environment check: issues found ({string.Join(" | ", issues)})");
            }
        }
        catch (Exception ex)
        {
            SetText(_diagStatusText, $"Diagnostics failed: {ex.Message}");
        }
    }

    private void RefreshWikiContent()
    {
        _wikiFullText = BuildWikiText();
        ApplyWikiFilter();
    }

    private void ApplyWikiFilter()
    {
        var fullText = _wikiFullText;
        if (string.IsNullOrWhiteSpace(fullText))
        {
            fullText = BuildWikiText();
            _wikiFullText = fullText;
        }

        var rawQuery = (_wikiSearchBox?.Text ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(rawQuery))
        {
            SetText(_wikiContentBox, fullText);
            SetText(_wikiStatusText, "Wiki ready. Search to filter.");
            return;
        }

        var tokens = rawQuery
            .Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Where(static token => token.Length >= 2)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
        if (tokens.Length == 0)
        {
            SetText(_wikiContentBox, fullText);
            SetText(_wikiStatusText, "Type at least 2 chars to filter.");
            return;
        }

        var sections = fullText
            .Split($"{Environment.NewLine}{Environment.NewLine}", StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .ToList();
        var matches = sections
            .Where(section => tokens.All(token => section.Contains(token, StringComparison.OrdinalIgnoreCase)))
            .ToList();

        if (matches.Count == 0)
        {
            SetText(_wikiContentBox, "No wiki matches found.");
            SetText(_wikiStatusText, $"No results for '{rawQuery}'.");
            return;
        }

        SetText(_wikiContentBox, string.Join($"{Environment.NewLine}{Environment.NewLine}", matches));
        SetText(_wikiStatusText, $"Wiki: {matches.Count} section(s) matched '{rawQuery}'.");
    }

    private string BuildWikiText()
    {
        var lines = new List<string>
        {
            "DesktopAgent Wiki",
            "Last update: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture),
            "",
            "## Quick Start",
            "1) arm",
            "2) run <intent>  (or type natural language in chat)",
            "3) confirm when requested",
            "4) disarm when done",
            "",
            "## Most Useful Commands",
            "- status",
            "- arm / disarm",
            "- simulate presence / require presence",
            "- force update now  (local tray command: check + force apply)",
            "- run \"open notepad and then type hello\"",
            "- run \"take screenshot\"",
            "- run \"take screenshot for each screen\"",
            "- run \"take screenshot single-screen\"",
            "- run \"record screen for 10 seconds\"",
            "- run \"start recording screen without audio\"",
            "- run \"stop recording\"",
            "- run \"file list .\"",
            "- run \"file search report in docs\"",
            "- run \"cerca file bolletta in .\"",
            "- order intake",
            "- order preview / order clear",
            "- order fill <url>",
            "",
            "## Plan Preview (Chat)",
            "- Shows the JSON plan generated from your request.",
            "- Human tab shows a readable step-by-step summary.",
            "- Validate: checks plan structure before execution.",
            "- Dry-run Plan: simulate without acting on desktop.",
            "- Execute Plan: run edited plan (policy and safety still enforced).",
            "- Supports both 'steps' and 'Steps' JSON keys.",
            "",
            "## Safety Model",
            "- Adapter starts DISARMED by default.",
            "- Allowlist can block actions: if active window app is not allowed, actions are denied.",
            "- Sensitive actions require confirmation (recording, dangerous clicks/keys, etc.).",
            "- Kill switch stops execution immediately.",
            "- Every action is written to audit logs.",
            "",
            "## Troubleshooting",
            "- 'Blocked: Active window not in allowlist':",
            "  Open Config and update AllowedApps, or clear the list to allow all apps.",
            "- 'Path not allowed by filesystem allowlist':",
            "  Open Config and extend FilesystemAllowedRoots with the target path.",
            "- 'Adapter unavailable':",
            "  Ensure adapter process is running on the configured endpoint.",
            "- 'LLM unavailable':",
            "  Check endpoint/model in Config and use Test LLM.",
            "- 'audio unavailable' during recording:",
            "  Verify ffmpeg build has audio backends and pick device/backend in Config.",
            "",
            "## Snapshot Modes",
            "- take screenshot                    => default mode",
            "- take screenshot for each screen    => mode:per-screen",
            "- take screenshot single-screen      => mode:single",
            "",
            "## Profiles",
            "- safe: stricter policy + lower rate",
            "- balanced: default",
            "- power: more permissive (but critical actions can still require confirmation)",
            "",
            "## Useful Tabs",
            "- Config: LLM, LLM parsing mode (primary/fallback), media folder, audio backend/device, updates",
            "- Chat: quick actions + plan preview editor (load/validate/dry-run/execute)",
            "- Media: screenshots/recordings preview and open",
            "- Audit: action log",
            "- Diagnostics: environment checks (ffmpeg, adapter, logs)",
            "",
            "## File Search",
            "- Syntax: file search <query> [in <path>]",
            "- Also supports: search file <query> [in <path>] / cerca file <query> [in <path>]",
            "- Wildcards supported in query: * and ?",
            "- Results are limited and recursively searched under the selected path.",
            "",
            "## Order Intake (AI)",
            "- You can ask naturally, e.g. 'mi e arrivata una mail ordine, compila il form'.",
            "- AI routing can map this to an order intake flow, even without the 'order' keyword command.",
            "- Use `order intake` and then paste the email text to extract structured fields.",
            "- Use `order preview` to review extracted data and `order clear` to reset.",
            "- Use `order fill <url>` to discover form fields and run smart safe autofill candidates (never auto-submit)."
        };

        return string.Join(Environment.NewLine, lines);
    }

    private static IEnumerable<string> ParseDirectShowAudioDevices(string? text)
    {
        var dump = text ?? string.Empty;
        if (string.IsNullOrWhiteSpace(dump))
        {
            return Array.Empty<string>();
        }

        var devices = new List<string>();
        var directAudioDevices = new List<string>();
        var inAudioSection = false;
        foreach (var rawLine in dump.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
        {
            var line = rawLine.Trim();
            if (line.Contains("Alternative name", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (line.Contains("DirectShow audio devices", StringComparison.OrdinalIgnoreCase))
            {
                inAudioSection = true;
                continue;
            }

            if (line.Contains("DirectShow video devices", StringComparison.OrdinalIgnoreCase))
            {
                inAudioSection = false;
            }

            var match = Regex.Match(line, "\"([^\"]+)\"");
            if (!match.Success)
            {
                continue;
            }

            var name = match.Groups[1].Value.Trim();
            if (string.IsNullOrWhiteSpace(name))
            {
                continue;
            }

            if (line.Contains("(audio)", StringComparison.OrdinalIgnoreCase))
            {
                directAudioDevices.Add(name);
            }
            else if (inAudioSection)
            {
                devices.Add(name);
            }
        }

        return directAudioDevices
            .Concat(devices)
            .Distinct(StringComparer.OrdinalIgnoreCase);
    }

    private static IEnumerable<string> ParseAvFoundationAudioDevices(string? text)
    {
        var dump = text ?? string.Empty;
        if (string.IsNullOrWhiteSpace(dump))
        {
            return Array.Empty<string>();
        }

        var inAudioSection = false;
        var devices = new List<string>();
        foreach (var rawLine in dump.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
        {
            var line = rawLine.Trim();
            if (line.Contains("AVFoundation audio devices", StringComparison.OrdinalIgnoreCase))
            {
                inAudioSection = true;
                continue;
            }

            if (line.Contains("AVFoundation video devices", StringComparison.OrdinalIgnoreCase))
            {
                inAudioSection = false;
            }

            if (!inAudioSection)
            {
                continue;
            }

            var match = Regex.Match(line, @"\[\d+\]\s+(.+)$");
            if (!match.Success)
            {
                continue;
            }

            var name = match.Groups[1].Value.Trim();
            if (!string.IsNullOrWhiteSpace(name))
            {
                devices.Add(name);
            }
        }

        return devices.Distinct(StringComparer.OrdinalIgnoreCase);
    }

    private static IEnumerable<string> ParsePulseAudioSources(string? text)
    {
        var dump = text ?? string.Empty;
        if (string.IsNullOrWhiteSpace(dump))
        {
            return Array.Empty<string>();
        }

        var devices = new List<string>();
        foreach (var rawLine in dump.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
        {
            var columns = rawLine.Split('\t', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            if (columns.Length < 2)
            {
                continue;
            }

            var name = columns[1].Trim();
            if (!string.IsNullOrWhiteSpace(name))
            {
                devices.Add(name);
            }
        }

        return devices.Distinct(StringComparer.OrdinalIgnoreCase);
    }

    private static string? ResolveCommandPath(string command)
    {
        if (string.IsNullOrWhiteSpace(command))
        {
            return null;
        }

        var trimmed = command.Trim();
        if (Path.IsPathRooted(trimmed) && File.Exists(trimmed))
        {
            return trimmed;
        }

        try
        {
            var checker = OperatingSystem.IsWindows() ? "where" : "which";
            var result = RunProcessCapture(checker, trimmed, 3000);
            if (result.TimedOut || result.ExitCode != 0)
            {
                return ResolveKnownCommandPath(trimmed);
            }

            var path = result.StdOut
                .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(static line => line.Trim())
                .FirstOrDefault(static line => !string.IsNullOrWhiteSpace(line));
            if (!string.IsNullOrWhiteSpace(path))
            {
                return path;
            }

            return ResolveKnownCommandPath(trimmed);
        }
        catch
        {
            return ResolveKnownCommandPath(trimmed);
        }
    }

    private static string? ResolveKnownCommandPath(string command)
    {
        foreach (var candidate in GetKnownCommandCandidates(command))
        {
            if (!string.IsNullOrWhiteSpace(candidate) && File.Exists(candidate))
            {
                return candidate;
            }
        }

        return null;
    }

    private static IEnumerable<string> GetKnownCommandCandidates(string command)
    {
        var cmd = command.Trim().ToLowerInvariant();
        if (OperatingSystem.IsWindows())
        {
            var programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            var programFilesX86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);
            if (cmd is "tesseract" or "tesseract.exe")
            {
                if (!string.IsNullOrWhiteSpace(programFiles))
                {
                    yield return Path.Combine(programFiles, "Tesseract-OCR", "tesseract.exe");
                }

                if (!string.IsNullOrWhiteSpace(programFilesX86))
                {
                    yield return Path.Combine(programFilesX86, "Tesseract-OCR", "tesseract.exe");
                }
            }

            if (cmd is "ffmpeg" or "ffmpeg.exe")
            {
                var localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                if (!string.IsNullOrWhiteSpace(localAppData))
                {
                    yield return Path.Combine(localAppData, "Microsoft", "WinGet", "Links", "ffmpeg.exe");
                }
            }
        }
        else if (OperatingSystem.IsMacOS())
        {
            if (cmd == "tesseract")
            {
                yield return "/opt/homebrew/bin/tesseract";
                yield return "/usr/local/bin/tesseract";
                yield return "/usr/bin/tesseract";
            }
        }
        else
        {
            if (cmd == "tesseract")
            {
                yield return "/usr/bin/tesseract";
                yield return "/usr/local/bin/tesseract";
            }
        }
    }

    private static ProcessCaptureResult RunProcessCapture(string fileName, string arguments, int timeoutMs)
    {
        try
        {
            using var process = Process.Start(new ProcessStartInfo
            {
                FileName = fileName,
                Arguments = arguments,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden
            });
            if (process == null)
            {
                return new ProcessCaptureResult(-1, string.Empty, "Process start returned null.", TimedOut: false);
            }

            var exited = process.WaitForExit(timeoutMs);
            if (!exited)
            {
                try
                {
                    process.Kill(entireProcessTree: true);
                }
                catch
                {
                    // ignored
                }

                return new ProcessCaptureResult(-1, string.Empty, "Process timeout.", TimedOut: true);
            }

            var output = process.StandardOutput.ReadToEnd();
            var error = process.StandardError.ReadToEnd();
            return new ProcessCaptureResult(process.ExitCode, output ?? string.Empty, error ?? string.Empty, TimedOut: false);
        }
        catch (Exception ex)
        {
            return new ProcessCaptureResult(-1, string.Empty, ex.Message, TimedOut: false);
        }
    }

    private readonly record struct ProcessCaptureResult(int ExitCode, string StdOut, string StdErr, bool TimedOut);

    private static IEnumerable<string> ReadFileTail(string path, int maxLines)
    {
        try
        {
            if (!File.Exists(path))
            {
                return new[] { "<not found>" };
            }

            return File.ReadLines(path).TakeLast(Math.Max(1, maxLines));
        }
        catch (Exception ex)
        {
            return new[] { $"<read failed: {ex.Message}>" };
        }
    }

    private void AppendUser(string text) => AppendLine("YOU", text);
    private void AppendAgent(string text) => AppendLine("AGENT", text);
    private void AppendSystem(string text) => AppendLine("SYSTEM", text);

    private async Task AppendAgentStreamingAsync(string text, CancellationToken cancellationToken)
    {
        var cleanText = text.Replace("\r", string.Empty).Replace('\n', ' ').Trim();
        if (string.IsNullOrWhiteSpace(cleanText))
        {
            AppendAgent("<no reply>");
            return;
        }

        if (cleanText.Length < 90)
        {
            AppendAgent(cleanText);
            return;
        }

        _streamingCts?.Cancel();
        _streamingCts?.Dispose();
        _streamingCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        var token = _streamingCts.Token;
        var timestamp = DateTime.Now.ToString("HH:mm:ss", CultureInfo.InvariantCulture);
        var prefix = $"[{timestamp}] {"AGENT",-6} ";

        await Dispatcher.UIThread.InvokeAsync(() =>
        {
            _historyLines.Enqueue($"{prefix}...");
            TrimHistoryLines();
            RewriteHistoryTextBox();

            if (!IsChatTabSelected())
            {
                _unreadChatEvents = Math.Min(_unreadChatEvents + 1, 999);
                _unreadInfoEvents = Math.Min(_unreadInfoEvents + 1, 999);
                UpdateUnreadBadge();
            }
        });

        var chunkSize = cleanText.Length > 800 ? 28 : 18;
        for (var i = chunkSize; i < cleanText.Length; i += chunkSize)
        {
            token.ThrowIfCancellationRequested();
            var partial = cleanText[..i];
            await Dispatcher.UIThread.InvokeAsync(() =>
            {
                RewriteLastHistoryLine($"{prefix}{partial}");
            });
            await Task.Delay(16, token);
        }

        await Dispatcher.UIThread.InvokeAsync(() =>
        {
            RewriteLastHistoryLine($"{prefix}{cleanText}");
            SchedulePersistSessionState();
        });
    }

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
            TrimHistoryLines();
            RewriteHistoryTextBox();

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

    private void TrimHistoryLines()
    {
        while (_historyLines.Count > MaxHistoryLines)
        {
            _historyLines.Dequeue();
        }
    }

    private void RewriteLastHistoryLine(string line)
    {
        if (_historyLines.Count == 0)
        {
            _historyLines.Enqueue(line);
            TrimHistoryLines();
            RewriteHistoryTextBox();
            return;
        }

        var lines = _historyLines.ToList();
        lines[^1] = line;
        _historyLines.Clear();
        foreach (var entry in lines)
        {
            _historyLines.Enqueue(entry);
        }

        TrimHistoryLines();
        RewriteHistoryTextBox();
    }

    private void RewriteHistoryTextBox()
    {
        SetText(_historyBox, string.Join(Environment.NewLine, _historyLines));
        if (_historyBox != null)
        {
            _historyBox.CaretIndex = _historyBox.Text?.Length ?? 0;
        }
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

        if (Dispatcher.UIThread.CheckAccess())
        {
            textBox.Text = value;
            return;
        }

        Dispatcher.UIThread.Post(() =>
        {
            try
            {
                textBox.Text = value;
            }
            catch
            {
                // ignored
            }
        });
    }

    private static void SetText(TextBlock? textBlock, string value)
    {
        if (textBlock == null)
        {
            return;
        }

        if (Dispatcher.UIThread.CheckAccess())
        {
            textBlock.Text = value;
            return;
        }

        Dispatcher.UIThread.Post(() =>
        {
            try
            {
                textBlock.Text = value;
            }
            catch
            {
                // ignored
            }
        });
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

    private void SelectLlmParsingMode(string mode)
    {
        if (_cfgLlmParsingMode == null)
        {
            return;
        }

        var normalized = mode.Trim().ToLowerInvariant();
        for (var i = 0; i < _cfgLlmParsingMode.ItemCount; i++)
        {
            if (_cfgLlmParsingMode.Items[i] is ComboBoxItem item &&
                string.Equals(item.Content?.ToString(), normalized, StringComparison.OrdinalIgnoreCase))
            {
                _cfgLlmParsingMode.SelectedIndex = i;
                return;
            }
        }

        _cfgLlmParsingMode.SelectedIndex = 0;
    }

    private void SelectAudioBackend(string backend)
    {
        if (_cfgAudioBackend == null)
        {
            return;
        }

        var normalized = backend.Trim().ToLowerInvariant();
        for (var i = 0; i < _cfgAudioBackend.ItemCount; i++)
        {
            if (_cfgAudioBackend.Items[i] is ComboBoxItem item &&
                string.Equals(item.Content?.ToString(), normalized, StringComparison.OrdinalIgnoreCase))
            {
                _cfgAudioBackend.SelectedIndex = i;
                return;
            }
        }

        _cfgAudioBackend.SelectedIndex = 0;
    }

    private string GetSelectedProvider()
    {
        if (_cfgLlmProvider?.SelectedItem is ComboBoxItem combo && combo.Content != null)
        {
            return combo.Content.ToString() ?? "ollama";
        }

        return "ollama";
    }

    private string GetSelectedLlmParsingMode()
    {
        if (_cfgLlmParsingMode?.SelectedItem is ComboBoxItem combo && combo.Content != null)
        {
            var mode = combo.Content.ToString() ?? "primary";
            return mode.Trim().ToLowerInvariant();
        }

        return "primary";
    }

    private string GetSelectedAudioBackend()
    {
        if (_cfgAudioBackend?.SelectedItem is ComboBoxItem combo && combo.Content != null)
        {
            return combo.Content.ToString() ?? "auto";
        }

        return "auto";
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

    private static string FormatMacroStep(int index, MacroStep step)
    {
        var label = string.Equals(step.Kind, "wait", StringComparison.OrdinalIgnoreCase)
            ? $"wait {step.Value}s"
            : step.Value;
        return $"{index + 1}. {label}";
    }

    private static string FormatMedia(MediaFileItem item)
    {
        var kind = item.IsImage ? "IMG" : item.IsVideo ? "VID" : "FILE";
        var size = FormatBytes(item.SizeBytes);
        var stamp = item.LastWriteUtc.ToLocalTime().ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture);
        return $"[{kind}] {item.Name} | {size} | {stamp}";
    }

    private static bool IsImageExtension(string extension)
    {
        return extension.Equals(".png", StringComparison.OrdinalIgnoreCase)
            || extension.Equals(".jpg", StringComparison.OrdinalIgnoreCase)
            || extension.Equals(".jpeg", StringComparison.OrdinalIgnoreCase)
            || extension.Equals(".bmp", StringComparison.OrdinalIgnoreCase)
            || extension.Equals(".gif", StringComparison.OrdinalIgnoreCase)
            || extension.Equals(".webp", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsVideoExtension(string extension)
    {
        return extension.Equals(".mp4", StringComparison.OrdinalIgnoreCase)
            || extension.Equals(".mkv", StringComparison.OrdinalIgnoreCase)
            || extension.Equals(".mov", StringComparison.OrdinalIgnoreCase)
            || extension.Equals(".avi", StringComparison.OrdinalIgnoreCase)
            || extension.Equals(".webm", StringComparison.OrdinalIgnoreCase);
    }

    private static string FormatBytes(long bytes)
    {
        const double k = 1024d;
        if (bytes < k)
        {
            return $"{bytes} B";
        }

        if (bytes < k * k)
        {
            return $"{bytes / k:0.0} KB";
        }

        if (bytes < k * k * k)
        {
            return $"{bytes / (k * k):0.0} MB";
        }

        return $"{bytes / (k * k * k):0.0} GB";
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

    private void UpdateConnectionBadges(bool armed, bool requirePresence, bool llmEnabled, bool llmAvailable, bool killTripped)
    {
        SetBadgeState(_armedStateBadge, _armedStateText, _armedStateIcon, "🛡", armed ? "ARMED ON" : "ARMED OFF", armed ? "#173526" : "#2A3344", armed ? "#2F6D4A" : "#3C4A63");
        SetBadgeState(_presenceStateBadge, _presenceStateText, _presenceStateIcon, "🔐", requirePresence ? "PRESENCE REQ" : "PRESENCE OFF", requirePresence ? "#3D2E15" : "#2A3344", requirePresence ? "#7D5B2D" : "#3C4A63");

        if (!llmEnabled)
        {
            SetBadgeState(_llmStateBadge, _llmStateText, _llmStateIcon, "🧠", "LLM OFF", "#2A3344", "#3C4A63");
        }
        else if (llmAvailable)
        {
            SetBadgeState(_llmStateBadge, _llmStateText, _llmStateIcon, "🧠", "LLM READY", "#163827", "#2F6D4A");
        }
        else
        {
            SetBadgeState(_llmStateBadge, _llmStateText, _llmStateIcon, "🧠", "LLM DOWN", "#3B1F1F", "#7E3D3D");
        }

        SetBadgeState(_killStateBadge, _killStateText, _killStateIcon, killTripped ? "⛔" : "✅", killTripped ? "KILL ON" : "KILL OFF", killTripped ? "#4A1818" : "#173526", killTripped ? "#8A3A3A" : "#2F6D4A");
    }

    private void SetConnectionBadgesUnknown()
    {
        SetBadgeState(_armedStateBadge, _armedStateText, _armedStateIcon, "🛡", "ARMED ?", "#2A3344", "#3C4A63");
        SetBadgeState(_presenceStateBadge, _presenceStateText, _presenceStateIcon, "🔐", "PRESENCE ?", "#2A3344", "#3C4A63");
        SetBadgeState(_llmStateBadge, _llmStateText, _llmStateIcon, "🧠", "LLM ?", "#2A3344", "#3C4A63");
        SetBadgeState(_killStateBadge, _killStateText, _killStateIcon, "⛔", "KILL ?", "#2A3344", "#3C4A63");
    }

    private static void SetBadgeState(Border? border, TextBlock? textBlock, TextBlock? iconBlock, string icon, string text, string background, string borderColor)
    {
        if (border == null || textBlock == null || iconBlock == null)
        {
            return;
        }

        Dispatcher.UIThread.Post(() =>
        {
            iconBlock.Text = icon;
            textBlock.Text = text;
            border.Background = new SolidColorBrush(Color.Parse(background));
            border.BorderBrush = new SolidColorBrush(Color.Parse(borderColor));
            iconBlock.Foreground = new SolidColorBrush(Color.Parse(borderColor));
        });
    }

    private static bool CommandExists(string command)
    {
        return !string.IsNullOrWhiteSpace(ResolveCommandPath(command));
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

        var cts = new CancellationTokenSource();
        var previous = Interlocked.Exchange(ref _persistSessionCts, cts);
        previous?.Cancel();

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
            finally
            {
                var current = Interlocked.CompareExchange(ref _persistSessionCts, null, cts);
                if (ReferenceEquals(current, cts))
                {
                    cts.Dispose();
                }
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

    private void OpenDataFolder()
    {
        try
        {
            var dataRoot = ResolveDataRoot();
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

    private static string ResolveDataRoot()
    {
        return Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "DesktopAgent");
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

internal sealed record MacroStep(string Kind, string Value);
internal sealed record MediaFileItem(string Path, string Name, long SizeBytes, DateTime LastWriteUtc, string Extension, bool IsImage, bool IsVideo);
internal sealed record ChatUpdateBadge(bool Visible, bool CanApply, string Text);
internal sealed record ChatUpdateDetails(bool HasUpdate, string Title, string Details, bool CanApply, bool HasBlockers, string Blockers);
