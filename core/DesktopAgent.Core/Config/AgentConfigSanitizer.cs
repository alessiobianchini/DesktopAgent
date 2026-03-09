using System.Runtime.InteropServices;

namespace DesktopAgent.Core.Config;

public static class AgentConfigSanitizer
{
    public static void Normalize(AgentConfig config)
    {
        if (config == null)
        {
            return;
        }

        config.AllowedApps = CleanTokens(config.AllowedApps);
        config.BlockedActionsKeywords = CleanTokens(config.BlockedActionsKeywords);
        config.DangerousKeyCombos = CleanTokens(config.DangerousKeyCombos);
        config.FilesystemAllowedRoots = CleanTokens(config.FilesystemAllowedRoots);
        config.ClipboardHistoryMaxItems = Math.Clamp(config.ClipboardHistoryMaxItems, 1, 500);
        config.AutoRecoveryMaxAttempts = Math.Clamp(config.AutoRecoveryMaxAttempts, 0, 3);
        config.AutoRecoveryWaitMs = Math.Clamp(config.AutoRecoveryWaitMs, 100, 5000);
        config.GoalSchedulerIntervalSeconds = Math.Clamp(config.GoalSchedulerIntervalSeconds, 10, 3600);
        config.GoalSchedulerMaxPerTick = Math.Clamp(config.GoalSchedulerMaxPerTick, 1, 10);
        config.PostCheckTimeoutMs = Math.Clamp(config.PostCheckTimeoutMs, 100, 5000);
        config.PostCheckPollMs = Math.Clamp(config.PostCheckPollMs, 20, 1000);
        config.ScreenRecordingAudioBackendPreference = NormalizeAudioBackendPreference(config.ScreenRecordingAudioBackendPreference);
        config.ScreenRecordingAudioDevice = (config.ScreenRecordingAudioDevice ?? string.Empty).Trim();
        config.MediaOutputDirectory = string.IsNullOrWhiteSpace(config.MediaOutputDirectory)
            ? "media"
            : config.MediaOutputDirectory.Trim();
        config.ActiveProfile = AgentProfileService.NormalizeProfile(config.ActiveProfile);
        config.Profiles ??= ProfilePresets.CreateDefault();

        config.NewFileKeyCombo = NormalizeKeySequence(config.NewFileKeyCombo, DefaultCombo("ctrl", "n", "cmd", "n"));
        config.SaveKeyCombo = NormalizeKeySequence(config.SaveKeyCombo, DefaultCombo("ctrl", "s", "cmd", "s"));
        config.SaveAsKeyCombo = NormalizeKeySequence(config.SaveAsKeyCombo, DefaultCombo("ctrl", "shift", "s", "cmd", "shift", "s"));
        config.NewTabKeyCombo = NormalizeKeySequence(config.NewTabKeyCombo, DefaultCombo("ctrl", "t", "cmd", "t"));
        config.CloseTabKeyCombo = NormalizeKeySequence(config.CloseTabKeyCombo, DefaultCombo("ctrl", "w", "cmd", "w"));
        config.CloseWindowKeyCombo = NormalizeKeySequence(config.CloseWindowKeyCombo, DefaultCombo("alt", "f4", "cmd", "q"));
        config.CopyKeyCombo = NormalizeKeySequence(config.CopyKeyCombo, DefaultCombo("ctrl", "c", "cmd", "c"));
        config.PasteKeyCombo = NormalizeKeySequence(config.PasteKeyCombo, DefaultCombo("ctrl", "v", "cmd", "v"));
        config.UndoKeyCombo = NormalizeKeySequence(config.UndoKeyCombo, DefaultCombo("ctrl", "z", "cmd", "z"));
        config.RedoKeyCombo = NormalizeKeySequence(config.RedoKeyCombo, DefaultCombo("ctrl", "y", "cmd", "shift", "z"));
        config.SelectAllKeyCombo = NormalizeKeySequence(config.SelectAllKeyCombo, DefaultCombo("ctrl", "a", "cmd", "a"));
        config.BrowserBackKeyCombo = NormalizeKeySequence(config.BrowserBackKeyCombo, DefaultCombo("alt", "left", "cmd", "["));
        config.BrowserForwardKeyCombo = NormalizeKeySequence(config.BrowserForwardKeyCombo, DefaultCombo("alt", "right", "cmd", "]"));
        config.RefreshKeyCombo = NormalizeKeySequence(config.RefreshKeyCombo, DefaultCombo("ctrl", "r", "cmd", "r"));
        config.FindInPageKeyCombo = NormalizeKeySequence(config.FindInPageKeyCombo, DefaultCombo("ctrl", "f", "cmd", "f"));

        if (config.ProfileModeEnabled)
        {
            AgentProfileService.ApplyActiveProfile(config);
        }
    }

    private static List<string> CleanTokens(List<string>? items)
    {
        if (items == null)
        {
            return new List<string>();
        }

        return items
            .Select(item => item?.Trim())
            .Where(item => !string.IsNullOrWhiteSpace(item))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList()!;
    }

    private static List<string> NormalizeKeySequence(List<string>? keys, IReadOnlyList<string> fallback)
    {
        var cleaned = CleanTokens(keys);
        if (cleaned.Count == 0)
        {
            return fallback.ToList();
        }

        var collapsed = CollapseRepeatedPattern(cleaned);
        return collapsed.Count == 0 ? fallback.ToList() : collapsed;
    }

    private static List<string> CollapseRepeatedPattern(List<string> keys)
    {
        if (keys.Count <= 1)
        {
            return keys;
        }

        for (var period = 1; period <= keys.Count / 2; period++)
        {
            if (keys.Count % period != 0)
            {
                continue;
            }

            var repeats = keys.Count / period;
            if (repeats < 2)
            {
                continue;
            }

            var matches = true;
            for (var i = period; i < keys.Count; i++)
            {
                if (!string.Equals(keys[i], keys[i % period], StringComparison.OrdinalIgnoreCase))
                {
                    matches = false;
                    break;
                }
            }

            if (matches)
            {
                return keys.Take(period).ToList();
            }
        }

        return keys;
    }

    private static IReadOnlyList<string> DefaultCombo(
        string windowsFirst,
        string windowsSecond,
        string macFirst,
        string macSecond)
    {
        return RuntimeInformation.IsOSPlatform(OSPlatform.OSX)
            ? new[] { macFirst, macSecond }
            : new[] { windowsFirst, windowsSecond };
    }

    private static IReadOnlyList<string> DefaultCombo(
        string windowsFirst,
        string windowsSecond,
        string windowsThird,
        string macFirst,
        string macSecond,
        string macThird)
    {
        return RuntimeInformation.IsOSPlatform(OSPlatform.OSX)
            ? new[] { macFirst, macSecond, macThird }
            : new[] { windowsFirst, windowsSecond, windowsThird };
    }

    private static IReadOnlyList<string> DefaultCombo(
        string windowsFirst,
        string windowsSecond,
        string macFirst,
        string macSecond,
        string macThird)
    {
        return RuntimeInformation.IsOSPlatform(OSPlatform.OSX)
            ? new[] { macFirst, macSecond, macThird }
            : new[] { windowsFirst, windowsSecond };
    }

    private static string NormalizeAudioBackendPreference(string? value)
    {
        var normalized = (value ?? string.Empty).Trim().ToLowerInvariant();
        return normalized switch
        {
            "wasapi" => "wasapi",
            "dshow" => "dshow",
            "avfoundation" => "avfoundation",
            "pipewire" => "pipewire",
            "pulse" => "pulse",
            "alsa" => "alsa",
            _ => "auto"
        };
    }
}
