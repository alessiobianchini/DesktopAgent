using DesktopAgent.Core.Abstractions;
using DesktopAgent.Core.Config;
using DesktopAgent.Core.Models;
using DesktopAgent.Core.Services;
using Xunit;

namespace DesktopAgent.Tests;

public sealed class AppResolverTests
{
    [Fact]
    public void TryResolveApp_PrefersMoreSpecificChromeMatch()
    {
        var catalog = new StubCatalog(new List<AppEntry>
        {
            new("Chrome Remote Desktop", @"C:\Apps\Chrome Remote Desktop.lnk"),
            new("Google Chrome", @"C:\Program Files\Google\Chrome\Application\chrome.exe"),
            new("Microsoft Edge", @"C:\Program Files\Microsoft\Edge\Application\msedge.exe")
        });

        var resolver = new AppResolver(new AgentConfig(), catalog);

        var ok = resolver.TryResolveApp("chrome", out var resolved);
        Assert.True(ok);
        Assert.Contains("chrome.exe", resolved, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void TryResolveApp_UsesBuiltInBrowserAliasEvenIfOnlyRemoteDesktopExists()
    {
        var catalog = new StubCatalog(new List<AppEntry>
        {
            new("Chrome Remote Desktop", @"C:\Apps\Chrome Remote Desktop.lnk")
        });

        var resolver = new AppResolver(new AgentConfig(), catalog);

        var ok = resolver.TryResolveApp("chrome", out var resolved);
        Assert.True(ok);
        Assert.Equal("chrome", resolved);
    }

    [Fact]
    public void TryResolveApp_UsesBuiltInBrowserAliasForCommandLikeInput()
    {
        var catalog = new StubCatalog(new List<AppEntry>
        {
            new("Chrome Remote Desktop", @"C:\Apps\Chrome Remote Desktop.lnk"),
            new("Google Chrome", @"C:\Program Files\Google\Chrome\Application\chrome.exe")
        });

        var resolver = new AppResolver(new AgentConfig(), catalog);

        var ok = resolver.TryResolveApp("chrome and go to meteoam", out var resolved);
        Assert.True(ok);
        Assert.Contains("chrome.exe", resolved, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void TryResolveApp_UsesConfiguredAlias_ForNotepadPlusPlusSpelling()
    {
        var catalog = new StubCatalog(new List<AppEntry>
        {
            new("Notepad++", @"C:\Program Files\Notepad++\notepad++.exe")
        });

        var config = new AgentConfig();
        config.AppAliases["notepad plus plus"] = "notepad++";
        var resolver = new AppResolver(config, catalog);

        var ok = resolver.TryResolveApp("notepad plus plus", out var resolved);
        Assert.True(ok);
        Assert.Contains("notepad++.exe", resolved, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void TryResolveApp_UsesBuiltInAlias_ForNotepadPlusPlusCompactForm()
    {
        var catalog = new StubCatalog(new List<AppEntry>
        {
            new("Notepad++", @"C:\Program Files\Notepad++\notepad++.exe")
        });

        var resolver = new AppResolver(new AgentConfig(), catalog);

        var ok = resolver.TryResolveApp("notepadplusplus", out var resolved);
        Assert.True(ok);
        Assert.Contains("notepad++.exe", resolved, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void TryResolveApp_UsesBuiltInAlias_ForTeams()
    {
        var catalog = new StubCatalog(new List<AppEntry>
        {
            new("ms-teams", @"C:\Users\User\AppData\Local\Microsoft\WindowsApps\ms-teams.exe")
        });

        var resolver = new AppResolver(new AgentConfig(), catalog);

        var ok = resolver.TryResolveApp("teams", out var resolved);
        Assert.True(ok);
        Assert.Contains("ms-teams.exe", resolved, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void TryResolveApp_UsesBuiltInAlias_ForMicrosoftTeamsWithoutCatalogMatch()
    {
        var catalog = new StubCatalog(Array.Empty<AppEntry>());
        var resolver = new AppResolver(new AgentConfig(), catalog);

        var ok = resolver.TryResolveApp("microsoft teams", out var resolved);
        Assert.True(ok);
        Assert.True(
            string.Equals(resolved, "ms-teams", StringComparison.OrdinalIgnoreCase)
            || resolved.Contains("ms-teams.exe", StringComparison.OrdinalIgnoreCase),
            $"Unexpected resolution: {resolved}");
    }

    [Fact]
    public void Suggest_MarksAppsAllowed_WhenAllowlistIsEmpty()
    {
        var catalog = new StubCatalog(new List<AppEntry>
        {
            new("Notepad", @"C:\Windows\notepad.exe"),
            new("Calculator", @"C:\Windows\System32\calc.exe")
        });
        var resolver = new AppResolver(new AgentConfig(), catalog);

        var suggestions = resolver.Suggest(string.Empty, 10);

        Assert.NotEmpty(suggestions);
        Assert.All(suggestions, s => Assert.True(s.IsAllowed));
    }

    private sealed class StubCatalog : IAppCatalog
    {
        private readonly IReadOnlyList<AppEntry> _entries;

        public StubCatalog(IReadOnlyList<AppEntry> entries)
        {
            _entries = entries;
        }

        public IReadOnlyList<AppEntry> GetAll() => _entries;
    }
}
