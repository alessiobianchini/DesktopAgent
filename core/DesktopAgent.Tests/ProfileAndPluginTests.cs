using DesktopAgent.Core.Abstractions;
using DesktopAgent.Core.Config;
using DesktopAgent.Core.Models;
using DesktopAgent.Core.Services;
using DesktopAgent.Proto;
using Xunit;

namespace DesktopAgent.Tests;

public sealed class ProfileAndPluginTests
{
    [Fact]
    public void ApplyActiveProfile_UsesSafePreset()
    {
        var config = new AgentConfig
        {
            ProfileModeEnabled = true,
            ActiveProfile = "safe",
            Profiles = ProfilePresets.CreateDefault()
        };

        AgentProfileService.ApplyActiveProfile(config);

        Assert.True(config.RequireConfirmation);
        Assert.Equal(1, config.MaxActionsPerSecond);
        Assert.True(config.ContextBindingRequireWindow);
    }

    [Fact]
    public void PolicyEngine_BlocksFileOutsideAllowedRoots()
    {
        var config = new AgentConfig
        {
            FilesystemAllowedRoots = new List<string> { ".\\safe-root" },
            RequireConfirmation = false
        };
        var engine = new PolicyEngine(config);
        var step = new PlanStep
        {
            Type = ActionType.FileWrite,
            Target = "C:\\Windows\\System32\\drivers\\etc\\hosts",
            Text = "x"
        };

        var decision = engine.Evaluate(step, new WindowRef());
        Assert.False(decision.Allowed);
        Assert.Contains("allowlist", decision.Reason, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Interpreter_ParsesPluginIntents()
    {
        var interpreter = new RuleBasedIntentInterpreter(new StubAppResolver(), new AgentConfig());

        var writePlan = interpreter.Interpret("file write notes.txt hello");
        Assert.Single(writePlan.Steps);
        Assert.Equal(ActionType.FileWrite, writePlan.Steps[0].Type);
        Assert.Equal("notes.txt", writePlan.Steps[0].Target);
        Assert.Equal("hello", writePlan.Steps[0].Text);

        var urlPlan = interpreter.Interpret("open url https://example.com");
        Assert.Single(urlPlan.Steps);
        Assert.Equal(ActionType.OpenUrl, urlPlan.Steps[0].Type);
    }

    [Fact]
    public void Interpreter_ParsesWebSearchIntents()
    {
        var interpreter = new RuleBasedIntentInterpreter(new StubAppResolver(), new AgentConfig());

        var browserPlan = interpreter.Interpret("search desktop agent on chrome");
        Assert.Equal(2, browserPlan.Steps.Count);
        Assert.Equal(ActionType.OpenApp, browserPlan.Steps[0].Type);
        Assert.Equal("chrome", browserPlan.Steps[0].AppIdOrPath);
        Assert.Equal(ActionType.OpenUrl, browserPlan.Steps[1].Type);
        Assert.Equal("chrome", browserPlan.Steps[1].AppIdOrPath);
        Assert.Equal("https://www.google.com/search?q=desktop%20agent", browserPlan.Steps[1].Target);

        var italianPlan = interpreter.Interpret("cerca meteo milano");
        Assert.Single(italianPlan.Steps);
        Assert.Equal(ActionType.OpenUrl, italianPlan.Steps[0].Type);
        Assert.Equal("https://www.google.com/search?q=meteo%20milano", italianPlan.Steps[0].Target);

        var gotoBrowserPlan = interpreter.Interpret("go to github.com on chrome");
        Assert.Equal(2, gotoBrowserPlan.Steps.Count);
        Assert.Equal(ActionType.OpenApp, gotoBrowserPlan.Steps[0].Type);
        Assert.Equal(ActionType.OpenUrl, gotoBrowserPlan.Steps[1].Type);
        Assert.Equal("chrome", gotoBrowserPlan.Steps[1].AppIdOrPath);
        Assert.Equal("https://github.com/", gotoBrowserPlan.Steps[1].Target);

        var italianGoto = interpreter.Interpret("vai a openai.com");
        Assert.Single(italianGoto.Steps);
        Assert.Equal(ActionType.OpenUrl, italianGoto.Steps[0].Type);
        Assert.Equal("https://openai.com/", italianGoto.Steps[0].Target);
    }

    [Fact]
    public void Interpreter_ParsesBrowserToolkitIntents()
    {
        var interpreter = new RuleBasedIntentInterpreter(new StubAppResolver(), new AgentConfig());

        var refreshPlan = interpreter.Interpret("refresh");
        Assert.Single(refreshPlan.Steps);
        Assert.Equal(ActionType.KeyCombo, refreshPlan.Steps[0].Type);
        Assert.Equal(new[] { "ctrl", "r" }, refreshPlan.Steps[0].Keys);

        var backPlan = interpreter.Interpret("go back");
        Assert.Single(backPlan.Steps);
        Assert.Equal(ActionType.KeyCombo, backPlan.Steps[0].Type);
        Assert.Equal(new[] { "alt", "left" }, backPlan.Steps[0].Keys);

        var findInPagePlan = interpreter.Interpret("find in page desktop agent");
        Assert.Equal(2, findInPagePlan.Steps.Count);
        Assert.Equal(ActionType.KeyCombo, findInPagePlan.Steps[0].Type);
        Assert.Equal(ActionType.TypeText, findInPagePlan.Steps[1].Type);
        Assert.Equal("desktop agent", findInPagePlan.Steps[1].Text);
    }

    [Fact]
    public void Interpreter_UsesSameBrowserInChainedOpenAndSearch()
    {
        var interpreter = new RuleBasedIntentInterpreter(new StubAppResolver(), new AgentConfig());

        var plan = interpreter.Interpret("open edge and search meteo gubbio");
        Assert.Equal(2, plan.Steps.Count);
        Assert.Equal(ActionType.OpenApp, plan.Steps[0].Type);
        Assert.Equal("edge", plan.Steps[0].AppIdOrPath);

        Assert.Equal(ActionType.OpenUrl, plan.Steps[1].Type);
        Assert.Equal("edge", plan.Steps[1].AppIdOrPath);
        Assert.Equal("https://www.google.com/search?q=meteo%20gubbio", plan.Steps[1].Target);
    }

    [Fact]
    public void Interpreter_ParsesMouseJiggleDuration()
    {
        var interpreter = new RuleBasedIntentInterpreter(new StubAppResolver(), new AgentConfig());

        var english = interpreter.Interpret("jiggle mouse for 5 minutes");
        Assert.Single(english.Steps);
        Assert.Equal(ActionType.MouseJiggle, english.Steps[0].Type);
        Assert.Equal(TimeSpan.FromMinutes(5), english.Steps[0].WaitFor);

        var italian = interpreter.Interpret("muovi randomicamente il mouse per 30 secondi");
        Assert.Single(italian.Steps);
        Assert.Equal(ActionType.MouseJiggle, italian.Steps[0].Type);
        Assert.Equal(TimeSpan.FromSeconds(30), italian.Steps[0].WaitFor);

        var typo = interpreter.Interpret("move mouse for 2 munutes");
        Assert.Single(typo.Steps);
        Assert.Equal(ActionType.MouseJiggle, typo.Steps[0].Type);
        Assert.Equal(TimeSpan.FromMinutes(2), typo.Steps[0].WaitFor);

        var prefixed = interpreter.Interpret("You: move mouse for 2 munutes");
        Assert.Single(prefixed.Steps);
        Assert.Equal(ActionType.MouseJiggle, prefixed.Steps[0].Type);
        Assert.Equal(TimeSpan.FromMinutes(2), prefixed.Steps[0].WaitFor);
    }

    [Fact]
    public void Interpreter_ConvertsGoToTextIntoSearchAndKeepsOpenedBrowser()
    {
        var interpreter = new RuleBasedIntentInterpreter(new StubAppResolver(), new AgentConfig());

        var plan = interpreter.Interpret("open chrome and go to meteoam");
        Assert.Equal(2, plan.Steps.Count);
        Assert.Equal(ActionType.OpenApp, plan.Steps[0].Type);
        Assert.Equal("chrome", plan.Steps[0].AppIdOrPath);

        Assert.Equal(ActionType.OpenUrl, plan.Steps[1].Type);
        Assert.Equal("chrome", plan.Steps[1].AppIdOrPath);
        Assert.Equal("https://www.google.com/search?q=meteoam", plan.Steps[1].Target);
    }

    private sealed class StubAppResolver : IAppResolver
    {
        public bool TryResolveApp(string input, out string resolved)
        {
            resolved = input;
            return false;
        }

        public IReadOnlyList<AppMatch> Suggest(string input, int maxResults) => Array.Empty<AppMatch>();
    }
}
