namespace DesktopAgent.Core.Config;

public static class AgentProfileService
{
    public static string NormalizeProfile(string? profile)
    {
        return profile?.Trim().ToLowerInvariant() switch
        {
            "safe" => "safe",
            "power" => "power",
            _ => "balanced"
        };
    }

    public static void ApplyActiveProfile(AgentConfig config)
    {
        var preset = ResolvePreset(config);
        config.RequireConfirmation = preset.RequireConfirmation;
        config.MaxActionsPerSecond = Math.Clamp(preset.MaxActionsPerSecond, 1, 60);
        config.QuizSafeModeEnabled = preset.QuizSafeModeEnabled;
        config.PostCheckStrict = preset.PostCheckStrict;
        config.ContextBindingEnabled = preset.ContextBindingEnabled;
        config.ContextBindingRequireWindow = preset.ContextBindingRequireWindow;
        config.ActiveProfile = NormalizeProfile(config.ActiveProfile);
    }

    public static ProfilePreset ResolvePreset(AgentConfig config)
    {
        var profile = NormalizeProfile(config.ActiveProfile);
        return profile switch
        {
            "safe" => config.Profiles.Safe ?? new ProfilePreset(),
            "power" => config.Profiles.Power ?? new ProfilePreset(),
            _ => config.Profiles.Balanced ?? new ProfilePreset()
        };
    }
}
