using DesktopAgent.Core.Abstractions;
using DesktopAgent.Core.Config;
using DesktopAgent.Core.Models;
using DesktopAgent.Proto;

namespace DesktopAgent.Core.Services;

public sealed class PolicyEngine : IPolicyEngine
{
    private readonly AgentConfig _config;

    public PolicyEngine(AgentConfig config)
    {
        _config = config;
    }

    public PolicyDecision Evaluate(PlanStep step, WindowRef? activeWindow)
    {
        if (IsFilesystemAction(step.Type) && !IsPathAllowed(step.Target))
        {
            return PolicyDecision.Deny("Filesystem target blocked by allowlist");
        }

        if (IsQuizContext(activeWindow) && _config.QuizSafeModeEnabled)
        {
            if (step.Type is ActionType.ReadText or ActionType.Find or ActionType.WaitForText)
            {
                return PolicyDecision.Allow();
            }

            return PolicyDecision.Deny("Quiz safe mode active: blocking interactive actions");
        }

        if (RequiresAllowlistCheck(step.Type))
        {
            if (!IsAllowedApp(activeWindow))
            {
                return PolicyDecision.Deny("Active window not in allowlist");
            }
        }

        if (IsDangerous(step))
        {
            if (_config.RequireConfirmation)
            {
                return PolicyDecision.RequireConfirmation("Action flagged as dangerous");
            }

            return PolicyDecision.Deny("Dangerous action blocked");
        }

        if (step.Type == ActionType.SetClipboard && _config.RequireConfirmation)
        {
            return PolicyDecision.RequireConfirmation("Clipboard modification requires confirmation");
        }

        if (step.Type == ActionType.ClipboardHistory && _config.RequireConfirmation)
        {
            return PolicyDecision.RequireConfirmation("Clipboard history access requires confirmation");
        }

        if (IsFilesystemWriteAction(step.Type) && _config.RequireConfirmation)
        {
            return PolicyDecision.RequireConfirmation("Filesystem write requires confirmation");
        }

        if (step.Type == ActionType.OpenUrl && _config.RequireConfirmation)
        {
            return PolicyDecision.RequireConfirmation("Opening URL requires confirmation");
        }

        if (step.Type == ActionType.LockScreen)
        {
            if (_config.RequireConfirmation)
            {
                return PolicyDecision.RequireConfirmation("Lock screen requires confirmation");
            }

            return PolicyDecision.Deny("Lock screen blocked when confirmation is disabled");
        }

        if (step.Type == ActionType.KeyCombo && step.Keys != null && step.Keys.Any(k => IsEnterKey(k)))
        {
            if (_config.RequireConfirmation)
            {
                return PolicyDecision.RequireConfirmation("Enter key requires confirmation");
            }
        }

        if (step.Type == ActionType.KeyCombo && step.Keys != null && IsDangerousKeyCombo(step.Keys))
        {
            if (_config.RequireConfirmation)
            {
                return PolicyDecision.RequireConfirmation("Potentially irreversible key combo requires confirmation");
            }

            return PolicyDecision.Deny("Potentially irreversible key combo blocked");
        }

        return PolicyDecision.Allow();
    }

    private bool IsAllowedApp(WindowRef? activeWindow)
    {
        if (_config.AllowedApps.Count == 0)
        {
            return true;
        }

        if (activeWindow == null)
        {
            return false;
        }

        return _config.AllowedApps.Any(app => activeWindow.AppId.Contains(app, StringComparison.OrdinalIgnoreCase)
                                              || activeWindow.Title.Contains(app, StringComparison.OrdinalIgnoreCase));
    }

    private bool RequiresAllowlistCheck(ActionType type)
    {
        return type is ActionType.Click
            or ActionType.DoubleClick
            or ActionType.RightClick
            or ActionType.Drag
            or ActionType.TypeText
            or ActionType.KeyCombo
            or ActionType.Invoke
            or ActionType.SetValue
            or ActionType.WaitForText;
    }

    private bool IsQuizContext(WindowRef? activeWindow)
    {
        if (activeWindow == null)
        {
            return false;
        }

        var title = (activeWindow.Title ?? string.Empty).ToLowerInvariant();
        var keywords = new[] { "quiz", "exam", "test", "assessment" };
        return keywords.Any(k => title.Contains(k));
    }

    private bool IsDangerous(PlanStep step)
    {
        var keywords = _config.BlockedActionsKeywords;
        if (keywords.Count == 0)
        {
            return false;
        }

        var targetText = string.Join(' ', new[] { step.Text ?? string.Empty, step.Target ?? string.Empty });
        if (step.Selector != null)
        {
            targetText = string.Join(" ", new[] { targetText, step.Selector.NameContains ?? string.Empty });
        }

        return keywords.Any(k => targetText.Contains(k, StringComparison.OrdinalIgnoreCase));
    }

    private static bool IsEnterKey(string key)
    {
        return key.Equals("enter", StringComparison.OrdinalIgnoreCase) || key.Equals("return", StringComparison.OrdinalIgnoreCase);
    }

    private bool IsDangerousKeyCombo(IEnumerable<string> keys)
    {
        if (_config.DangerousKeyCombos.Count == 0)
        {
            return false;
        }

        var normalized = keys.Select(NormalizeKey).Where(k => !string.IsNullOrWhiteSpace(k)).ToHashSet(StringComparer.OrdinalIgnoreCase);
        if (normalized.Count == 0)
        {
            return false;
        }

        foreach (var combo in _config.DangerousKeyCombos)
        {
            if (string.IsNullOrWhiteSpace(combo))
            {
                continue;
            }

            var parts = combo.Split(new[] { '+', ' ' }, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .Select(NormalizeKey)
                .Where(k => !string.IsNullOrWhiteSpace(k))
                .ToList();

            if (parts.Count == 0)
            {
                continue;
            }

            var allPresent = parts.All(part => normalized.Contains(part));
            if (allPresent)
            {
                return true;
            }
        }

        return false;
    }

    private static string NormalizeKey(string key)
    {
        var value = key.Trim().ToLowerInvariant();
        return value switch
        {
            "control" => "ctrl",
            "command" => "cmd",
            "option" => "alt",
            "windows" => "win",
            _ => value
        };
    }

    private static bool IsFilesystemAction(ActionType type)
    {
        return type is ActionType.FileWrite or ActionType.FileAppend or ActionType.FileRead or ActionType.FileList;
    }

    private static bool IsFilesystemWriteAction(ActionType type)
    {
        return type is ActionType.FileWrite or ActionType.FileAppend;
    }

    private bool IsPathAllowed(string? target)
    {
        var roots = _config.FilesystemAllowedRoots;
        if (roots.Count == 0)
        {
            return true;
        }

        if (string.IsNullOrWhiteSpace(target))
        {
            return false;
        }

        string fullPath;
        try
        {
            fullPath = Path.GetFullPath(target.Trim().Trim('"', '\''), AppContext.BaseDirectory);
        }
        catch
        {
            return false;
        }

        foreach (var root in roots)
        {
            if (string.IsNullOrWhiteSpace(root))
            {
                continue;
            }

            string rootPath;
            try
            {
                rootPath = Path.GetFullPath(root.Trim().Trim('"', '\''), AppContext.BaseDirectory);
            }
            catch
            {
                continue;
            }

            if (fullPath.StartsWith(rootPath, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }
}
