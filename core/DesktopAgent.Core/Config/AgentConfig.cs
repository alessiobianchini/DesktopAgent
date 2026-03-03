namespace DesktopAgent.Core.Config;

public sealed class AgentConfig
{
    public bool ProfileModeEnabled { get; set; } = false;
    public string ActiveProfile { get; set; } = "balanced";
    public string AdapterEndpoint { get; set; } = "http://localhost:51877";
    public List<string> AllowedApps { get; set; } = new();
    public Dictionary<string, string> AppAliases { get; set; } = new(StringComparer.OrdinalIgnoreCase);
    public string AppIndexCachePath { get; set; } = "app-index.json";
    public int AppIndexCacheTtlMinutes { get; set; } = 1440;
    public List<string> NewFileKeyCombo { get; set; } = new() { "ctrl", "n" };
    public List<string> SaveKeyCombo { get; set; } = new() { "ctrl", "s" };
    public List<string> SaveAsKeyCombo { get; set; } = new() { "ctrl", "shift", "s" };
    public List<string> NewTabKeyCombo { get; set; } = new() { "ctrl", "t" };
    public List<string> CloseTabKeyCombo { get; set; } = new() { "ctrl", "w" };
    public List<string> CloseWindowKeyCombo { get; set; } = new() { "alt", "f4" };
    public List<string> CopyKeyCombo { get; set; } = new() { "ctrl", "c" };
    public List<string> PasteKeyCombo { get; set; } = new() { "ctrl", "v" };
    public List<string> UndoKeyCombo { get; set; } = new() { "ctrl", "z" };
    public List<string> RedoKeyCombo { get; set; } = new() { "ctrl", "y" };
    public List<string> SelectAllKeyCombo { get; set; } = new() { "ctrl", "a" };
    public List<string> BrowserBackKeyCombo { get; set; } = new() { "alt", "left" };
    public List<string> BrowserForwardKeyCombo { get; set; } = new() { "alt", "right" };
    public List<string> RefreshKeyCombo { get; set; } = new() { "ctrl", "r" };
    public List<string> FindInPageKeyCombo { get; set; } = new() { "ctrl", "f" };
    public List<string> DangerousKeyCombos { get; set; } = new() { "alt+f4", "ctrl+w", "cmd+w", "cmd+q" };
    public List<string> BlockedActionsKeywords { get; set; } = new();
    public bool RequireConfirmation { get; set; } = true;
    public int MaxActionsPerSecond { get; set; } = 3;
    public bool MonitoringEnabled { get; set; } = false;
    public bool OcrEnabled { get; set; } = false;
    public bool QuizSafeModeEnabled { get; set; } = true;
    public string AuditLogPath { get; set; } = "audit.log.jsonl";
    public bool PostCheckStrict { get; set; } = true;
    public PostCheckRules PostCheckRules { get; set; } = new();
    public int PostCheckTimeoutMs { get; set; } = 900;
    public int PostCheckPollMs { get; set; } = 120;
    public OcrConfig Ocr { get; set; } = new();
    public LlmFallbackConfig LlmFallback { get; set; } = new();
    public bool LlmFallbackEnabled { get; set; } = false;
    public bool AllowNonLoopbackLlmEndpoint { get; set; } = true;
    public bool AuditLlmInteractions { get; set; } = true;
    public bool AuditLlmIncludeRawText { get; set; } = false;
    public string AdapterRestartCommand { get; set; } = "";
    public string AdapterRestartWorkingDir { get; set; } = "";
    public int OpenAppSettleDelayMs { get; set; } = 700;
    public int FindRetryCount { get; set; } = 2;
    public int FindRetryDelayMs { get; set; } = 250;
    public bool ContextBindingEnabled { get; set; } = true;
    public bool ContextBindingRequireWindow { get; set; } = false;
    public bool AutoRecoveryEnabled { get; set; } = true;
    public int AutoRecoveryMaxAttempts { get; set; } = 1;
    public int AutoRecoveryWaitMs { get; set; } = 700;
    public bool GoalSchedulerEnabled { get; set; } = true;
    public int GoalSchedulerIntervalSeconds { get; set; } = 300;
    public int GoalSchedulerMaxPerTick { get; set; } = 1;
    public int ClipboardHistoryMaxItems { get; set; } = 50;
    public List<string> FilesystemAllowedRoots { get; set; } = new() { "." };
    public ProfilePresets Profiles { get; set; } = ProfilePresets.CreateDefault();
    public string TaskLibraryPath { get; set; } = "tasks.library.json";
    public string ScheduleLibraryPath { get; set; } = "schedules.library.json";
}

public sealed class ProfilePresets
{
    public ProfilePreset Safe { get; set; } = new();
    public ProfilePreset Balanced { get; set; } = new();
    public ProfilePreset Power { get; set; } = new();

    public static ProfilePresets CreateDefault()
    {
        return new ProfilePresets
        {
            Safe = new ProfilePreset
            {
                RequireConfirmation = true,
                MaxActionsPerSecond = 1,
                QuizSafeModeEnabled = true,
                PostCheckStrict = true,
                ContextBindingEnabled = true,
                ContextBindingRequireWindow = true
            },
            Balanced = new ProfilePreset
            {
                RequireConfirmation = true,
                MaxActionsPerSecond = 3,
                QuizSafeModeEnabled = true,
                PostCheckStrict = true,
                ContextBindingEnabled = true,
                ContextBindingRequireWindow = false
            },
            Power = new ProfilePreset
            {
                RequireConfirmation = false,
                MaxActionsPerSecond = 8,
                QuizSafeModeEnabled = true,
                PostCheckStrict = false,
                ContextBindingEnabled = true,
                ContextBindingRequireWindow = false
            }
        };
    }
}

public sealed class ProfilePreset
{
    public bool RequireConfirmation { get; set; }
    public int MaxActionsPerSecond { get; set; }
    public bool QuizSafeModeEnabled { get; set; }
    public bool PostCheckStrict { get; set; }
    public bool ContextBindingEnabled { get; set; }
    public bool ContextBindingRequireWindow { get; set; }
}

public sealed class PostCheckRules
{
    public string MenuItem { get; set; } = "window-change";
    public string Checkbox { get; set; } = "present";
    public string Button { get; set; } = "disappear-or-window";
}

public sealed class OcrConfig
{
    public string Engine { get; set; } = "tesseract";
    public string TesseractPath { get; set; } = "tesseract";
}

public sealed class LlmFallbackConfig
{
    public string Provider { get; set; } = "ollama";
    public string Endpoint { get; set; } = "http://localhost:11434/api/generate";
    public string Model { get; set; } = "llama3";
    public int TimeoutSeconds { get; set; } = 10;
    public int MaxTokens { get; set; } = 128;
}
