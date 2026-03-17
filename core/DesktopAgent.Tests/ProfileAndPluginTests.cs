using DesktopAgent.Core.Abstractions;
using DesktopAgent.Core.Config;
using DesktopAgent.Core.Models;
using DesktopAgent.Core.Services;
using DesktopAgent.Proto;
using System.Runtime.InteropServices;
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

        var fileSearchPlan = interpreter.Interpret("file search report in docs");
        Assert.Single(fileSearchPlan.Steps);
        Assert.Equal(ActionType.FileList, fileSearchPlan.Steps[0].Type);
        Assert.Equal("docs", fileSearchPlan.Steps[0].Target);
        Assert.Equal("report", fileSearchPlan.Steps[0].Text);

        var italianFileSearchPlan = interpreter.Interpret("cerca file bolletta in .");
        Assert.Single(italianFileSearchPlan.Steps);
        Assert.Equal(ActionType.FileList, italianFileSearchPlan.Steps[0].Type);
        Assert.Equal(".", italianFileSearchPlan.Steps[0].Target);
        Assert.Equal("bolletta", italianFileSearchPlan.Steps[0].Text);
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
        var isMac = RuntimeInformation.IsOSPlatform(OSPlatform.OSX);

        var refreshPlan = interpreter.Interpret("refresh");
        Assert.Single(refreshPlan.Steps);
        Assert.Equal(ActionType.KeyCombo, refreshPlan.Steps[0].Type);
        Assert.Equal(isMac ? new[] { "cmd", "r" } : new[] { "ctrl", "r" }, refreshPlan.Steps[0].Keys);

        var backPlan = interpreter.Interpret("go back");
        Assert.Single(backPlan.Steps);
        Assert.Equal(ActionType.KeyCombo, backPlan.Steps[0].Type);
        Assert.Equal(isMac ? new[] { "cmd", "[" } : new[] { "alt", "left" }, backPlan.Steps[0].Keys);

        var findInPagePlan = interpreter.Interpret("find in page desktop agent");
        Assert.Equal(2, findInPagePlan.Steps.Count);
        Assert.Equal(ActionType.KeyCombo, findInPagePlan.Steps[0].Type);
        Assert.Equal(isMac ? new[] { "cmd", "f" } : new[] { "ctrl", "f" }, findInPagePlan.Steps[0].Keys);
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
    public void Interpreter_ParsesSnapshotAndScreenRecording()
    {
        var interpreter = new RuleBasedIntentInterpreter(new StubAppResolver(), new AgentConfig());

        var snapshot = interpreter.Interpret("take screenshot");
        Assert.Single(snapshot.Steps);
        Assert.Equal(ActionType.CaptureScreen, snapshot.Steps[0].Type);

        var perScreen = interpreter.Interpret("take screenshot for each screen");
        Assert.Single(perScreen.Steps);
        Assert.Equal(ActionType.CaptureScreen, perScreen.Steps[0].Type);
        Assert.Equal("mode:per-screen", perScreen.Steps[0].Text);

        var italianPerScreen = interpreter.Interpret("fai una foto per schermo");
        Assert.Single(italianPerScreen.Steps);
        Assert.Equal(ActionType.CaptureScreen, italianPerScreen.Steps[0].Type);
        Assert.Equal("mode:per-screen", italianPerScreen.Steps[0].Text);

        var singleScreen = interpreter.Interpret("take snapshot single-screen");
        Assert.Single(singleScreen.Steps);
        Assert.Equal(ActionType.CaptureScreen, singleScreen.Steps[0].Type);
        Assert.Equal("mode:single", singleScreen.Steps[0].Text);

        var italianSingleScreen = interpreter.Interpret("fai screenshot su schermo principale");
        Assert.Single(italianSingleScreen.Steps);
        Assert.Equal(ActionType.CaptureScreen, italianSingleScreen.Steps[0].Type);
        Assert.Equal("mode:single", italianSingleScreen.Steps[0].Text);

        var conversationalItalian = interpreter.Interpret("puoi fare una cattura schermo?");
        Assert.Single(conversationalItalian.Steps);
        Assert.Equal(ActionType.CaptureScreen, conversationalItalian.Steps[0].Type);

        var chainedSnapshotAndOpen = interpreter.Interpret("puoi fare una cattura schermo e aprire notepad?");
        Assert.Equal(2, chainedSnapshotAndOpen.Steps.Count);
        Assert.Equal(ActionType.CaptureScreen, chainedSnapshotAndOpen.Steps[0].Type);
        Assert.Equal(ActionType.OpenApp, chainedSnapshotAndOpen.Steps[1].Type);
        Assert.Equal("notepad", chainedSnapshotAndOpen.Steps[1].AppIdOrPath);

        var chainedSnapshotThenOpen = interpreter.Interpret("fai uno screenshot, poi apri notepad");
        Assert.Equal(2, chainedSnapshotThenOpen.Steps.Count);
        Assert.Equal(ActionType.CaptureScreen, chainedSnapshotThenOpen.Steps[0].Type);
        Assert.Equal(ActionType.OpenApp, chainedSnapshotThenOpen.Steps[1].Type);
        Assert.Equal("notepad", chainedSnapshotThenOpen.Steps[1].AppIdOrPath);

        var recordWithAudio = interpreter.Interpret("record screen and audio for 2 minutes");
        Assert.Single(recordWithAudio.Steps);
        Assert.Equal(ActionType.RecordScreen, recordWithAudio.Steps[0].Type);
        Assert.Equal(TimeSpan.FromMinutes(2), recordWithAudio.Steps[0].WaitFor);
        Assert.Equal("audio:on", recordWithAudio.Steps[0].Text);

        var recordWithoutAudio = interpreter.Interpret("registra schermo senza audio per 45 secondi");
        Assert.Single(recordWithoutAudio.Steps);
        Assert.Equal(ActionType.RecordScreen, recordWithoutAudio.Steps[0].Type);
        Assert.Equal(TimeSpan.FromSeconds(45), recordWithoutAudio.Steps[0].WaitFor);
        Assert.Equal("audio:off", recordWithoutAudio.Steps[0].Text);

        var startRecording = interpreter.Interpret("start recording screen without audio");
        Assert.Single(startRecording.Steps);
        Assert.Equal(ActionType.StartScreenRecording, startRecording.Steps[0].Type);
        Assert.Equal("audio:off", startRecording.Steps[0].Text);

        var stopRecording = interpreter.Interpret("stop recording");
        Assert.Single(stopRecording.Steps);
        Assert.Equal(ActionType.StopScreenRecording, stopRecording.Steps[0].Type);
    }

    [Fact]
    public void PolicyEngine_RequiresConfirmation_ForScreenRecording()
    {
        var engine = new PolicyEngine(new AgentConfig { RequireConfirmation = true });
        var step = new PlanStep { Type = ActionType.RecordScreen, WaitFor = TimeSpan.FromSeconds(10) };

        var decision = engine.Evaluate(step, new WindowRef());
        Assert.True(decision.Allowed);
        Assert.True(decision.RequiresConfirmation);

        var startStep = new PlanStep { Type = ActionType.StartScreenRecording };
        var startDecision = engine.Evaluate(startStep, new WindowRef());
        Assert.True(startDecision.Allowed);
        Assert.True(startDecision.RequiresConfirmation);
    }

    [Fact]
    public void PolicyEngine_StillRequiresConfirmation_ForScreenRecording_WhenGlobalConfirmationDisabled()
    {
        var engine = new PolicyEngine(new AgentConfig { RequireConfirmation = false });

        var recordDecision = engine.Evaluate(new PlanStep { Type = ActionType.RecordScreen }, new WindowRef());
        Assert.True(recordDecision.Allowed);
        Assert.True(recordDecision.RequiresConfirmation);

        var startDecision = engine.Evaluate(new PlanStep { Type = ActionType.StartScreenRecording }, new WindowRef());
        Assert.True(startDecision.Allowed);
        Assert.True(startDecision.RequiresConfirmation);
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

    [Fact]
    public void Interpreter_ParsesConversationalOpenCommands()
    {
        var interpreter = new RuleBasedIntentInterpreter(new StubAppResolver(), new AgentConfig());

        var italian = interpreter.Interpret("potresti aprirmi notepad?");
        Assert.Single(italian.Steps);
        Assert.Equal(ActionType.OpenApp, italian.Steps[0].Type);
        Assert.Equal("notepad", italian.Steps[0].AppIdOrPath);

        var english = interpreter.Interpret("can you open for me notepad");
        Assert.Single(english.Steps);
        Assert.Equal(ActionType.OpenApp, english.Steps[0].Type);
        Assert.Equal("notepad", english.Steps[0].AppIdOrPath);

        var chatPrefixed = interpreter.Interpret("You: can you open for me notepad?");
        Assert.Single(chatPrefixed.Steps);
        Assert.Equal(ActionType.OpenApp, chatPrefixed.Steps[0].Type);
        Assert.Equal("notepad", chatPrefixed.Steps[0].AppIdOrPath);
    }

    [Fact]
    public void Interpreter_ParsesConversationalChainedCommands()
    {
        var interpreter = new RuleBasedIntentInterpreter(new StubAppResolver(), new AgentConfig());

        var italian = interpreter.Interpret("mi apri notepad e poi scrivi ciao");
        Assert.Equal(2, italian.Steps.Count);
        Assert.Equal(ActionType.OpenApp, italian.Steps[0].Type);
        Assert.Equal("notepad", italian.Steps[0].AppIdOrPath);
        Assert.Equal(ActionType.TypeText, italian.Steps[1].Type);
        Assert.Equal("ciao", italian.Steps[1].Text);
        Assert.Equal("notepad", italian.Steps[1].ExpectedAppId);

        var english = interpreter.Interpret("can you open for me notepad and then write hello");
        Assert.Equal(2, english.Steps.Count);
        Assert.Equal(ActionType.OpenApp, english.Steps[0].Type);
        Assert.Equal("notepad", english.Steps[0].AppIdOrPath);
        Assert.Equal(ActionType.TypeText, english.Steps[1].Type);
        Assert.Equal("hello", english.Steps[1].Text);
        Assert.Equal("notepad", english.Steps[1].ExpectedAppId);
    }

    [Fact]
    public void Interpreter_KeepsBrowserContext_InConversationalChains()
    {
        var interpreter = new RuleBasedIntentInterpreter(new StubAppResolver(), new AgentConfig());

        var onBrowser = interpreter.Interpret("open edge and search meteo gubbio on browser");
        Assert.Equal(2, onBrowser.Steps.Count);
        Assert.Equal(ActionType.OpenApp, onBrowser.Steps[0].Type);
        Assert.Equal("edge", onBrowser.Steps[0].AppIdOrPath);
        Assert.Equal(ActionType.OpenUrl, onBrowser.Steps[1].Type);
        Assert.Equal("edge", onBrowser.Steps[1].AppIdOrPath);
        Assert.Equal("https://www.google.com/search?q=meteo%20gubbio", onBrowser.Steps[1].Target);

        var there = interpreter.Interpret("open edge and then search meteo gubbio there");
        Assert.Equal(2, there.Steps.Count);
        Assert.Equal(ActionType.OpenApp, there.Steps[0].Type);
        Assert.Equal("edge", there.Steps[0].AppIdOrPath);
        Assert.Equal(ActionType.OpenUrl, there.Steps[1].Type);
        Assert.Equal("edge", there.Steps[1].AppIdOrPath);
        Assert.Equal("https://www.google.com/search?q=meteo%20gubbio", there.Steps[1].Target);
    }

    [Fact]
    public void Interpreter_ParsesChatLikeTaskCommands_WithSendConfirmationHook()
    {
        var interpreter = new RuleBasedIntentInterpreter(new StubAppResolver(), new AgentConfig());

        var plan = interpreter.Interpret("apri teams, cerca mario rossi e inviagli ciao");
        Assert.Equal(4, plan.Steps.Count);

        Assert.Equal(ActionType.OpenApp, plan.Steps[0].Type);
        Assert.Equal("teams", plan.Steps[0].AppIdOrPath);

        Assert.Equal(ActionType.Find, plan.Steps[1].Type);
        Assert.Equal("mario rossi", plan.Steps[1].Selector?.NameContains);
        Assert.Equal("teams", plan.Steps[1].ExpectedAppId);

        Assert.Equal(ActionType.TypeText, plan.Steps[2].Type);
        Assert.Equal("ciao", plan.Steps[2].Text);
        Assert.Equal("teams", plan.Steps[2].ExpectedAppId);

        Assert.Equal(ActionType.Click, plan.Steps[3].Type);
        Assert.Equal("send", plan.Steps[3].Selector?.NameContains);
        Assert.Equal("send", plan.Steps[3].Target);
        Assert.Equal("teams", plan.Steps[3].ExpectedAppId);
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
