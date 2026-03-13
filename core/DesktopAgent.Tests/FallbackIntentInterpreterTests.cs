using DesktopAgent.Core.Abstractions;
using DesktopAgent.Core.Config;
using DesktopAgent.Core.Models;
using DesktopAgent.Core.Services;
using Xunit;

namespace DesktopAgent.Tests;

public sealed class FallbackIntentInterpreterTests
{
    [Fact]
    public void Interpret_UsesLlmRewriteFirst_WhenEnabled()
    {
        var config = new AgentConfig { LlmFallbackEnabled = true };
        var ruleBased = new RuleBasedIntentInterpreter(new StubAppResolver(), config);
        var interpreter = new FallbackIntentInterpreter(ruleBased, new StubRewriter("open notepad"), new StubAuditLog(), config);

        var plan = interpreter.Interpret("open calculator");

        Assert.Single(plan.Steps);
        Assert.Equal(ActionType.OpenApp, plan.Steps[0].Type);
        Assert.Equal("notepad", plan.Steps[0].AppIdOrPath);
        Assert.StartsWith("Rewritten intent:", plan.Steps[0].Note);
    }

    [Fact]
    public void Interpret_FallsBackToRuleBased_WhenLlmReturnsNothing()
    {
        var config = new AgentConfig { LlmFallbackEnabled = true };
        var ruleBased = new RuleBasedIntentInterpreter(new StubAppResolver(), config);
        var interpreter = new FallbackIntentInterpreter(ruleBased, new StubRewriter(null), new StubAuditLog(), config);

        var plan = interpreter.Interpret("open calculator");

        Assert.Single(plan.Steps);
        Assert.Equal(ActionType.OpenApp, plan.Steps[0].Type);
        Assert.Equal("calculator", plan.Steps[0].AppIdOrPath);
    }

    [Fact]
    public void Interpret_FallsBackToRuleBased_WhenRewrittenPlanIsUnusable()
    {
        var config = new AgentConfig { LlmFallbackEnabled = true };
        var ruleBased = new RuleBasedIntentInterpreter(new StubAppResolver(), config);
        var interpreter = new FallbackIntentInterpreter(ruleBased, new StubRewriter("totally unknown command"), new StubAuditLog(), config);

        var plan = interpreter.Interpret("open calculator");

        Assert.Single(plan.Steps);
        Assert.Equal(ActionType.OpenApp, plan.Steps[0].Type);
        Assert.Equal("calculator", plan.Steps[0].AppIdOrPath);
    }

    [Fact]
    public void Interpret_PreservesMouseDuration_FromOriginalIntent_WhenRewriteDropsIt()
    {
        var config = new AgentConfig { LlmFallbackEnabled = true };
        var ruleBased = new RuleBasedIntentInterpreter(new StubAppResolver(), config);
        var interpreter = new FallbackIntentInterpreter(ruleBased, new StubRewriter("move mouse"), new StubAuditLog(), config);

        var plan = interpreter.Interpret("move mouse for 2 munutes");

        Assert.Single(plan.Steps);
        Assert.Equal(ActionType.MouseJiggle, plan.Steps[0].Type);
        Assert.Equal(TimeSpan.FromMinutes(2), plan.Steps[0].WaitFor);
    }

    [Fact]
    public void Interpret_UsesRewrittenPlan_WhenActionableEvenWithUnknownNoise()
    {
        var config = new AgentConfig { LlmFallbackEnabled = true };
        var ruleBased = new RuleBasedIntentInterpreter(new StubAppResolver(), config);
        var interpreter = new FallbackIntentInterpreter(ruleBased, new StubRewriter("open teams and then blah blah"), new StubAuditLog(), config);

        var plan = interpreter.Interpret("apri teams e poi fai qualcosa");

        Assert.Single(plan.Steps);
        Assert.Equal(ActionType.OpenApp, plan.Steps[0].Type);
        Assert.Equal("teams", plan.Steps[0].AppIdOrPath);
    }

    [Fact]
    public void Interpret_PreservesSnapshotPlan_WhenLlmRewriteChangesIntent()
    {
        var config = new AgentConfig { LlmFallbackEnabled = true };
        var ruleBased = new RuleBasedIntentInterpreter(new StubAppResolver(), config);
        var interpreter = new FallbackIntentInterpreter(ruleBased, new StubRewriter("find snapshot"), new StubAuditLog(), config);

        var plan = interpreter.Interpret("take snapshot");

        Assert.Single(plan.Steps);
        Assert.Equal(ActionType.CaptureScreen, plan.Steps[0].Type);
    }

    [Fact]
    public void Interpret_PreservesSnapshotChain_WhenLlmRewriteDropsSnapshot()
    {
        var config = new AgentConfig { LlmFallbackEnabled = true };
        var ruleBased = new RuleBasedIntentInterpreter(new StubAppResolver(), config);
        var interpreter = new FallbackIntentInterpreter(ruleBased, new StubRewriter("open chrome and then open notepad"), new StubAuditLog(), config);

        var plan = interpreter.Interpret("puoi fare una cattura schermo e aprire notepad?");

        Assert.Equal(2, plan.Steps.Count);
        Assert.Equal(ActionType.CaptureScreen, plan.Steps[0].Type);
        Assert.Equal(ActionType.OpenApp, plan.Steps[1].Type);
        Assert.Equal("notepad", plan.Steps[1].AppIdOrPath);
    }

    [Fact]
    public void Interpret_PreservesScreenRecording_WhenLlmRewriteDropsRecording()
    {
        var config = new AgentConfig { LlmFallbackEnabled = true };
        var ruleBased = new RuleBasedIntentInterpreter(new StubAppResolver(), config);
        var interpreter = new FallbackIntentInterpreter(ruleBased, new StubRewriter("open chrome"), new StubAuditLog(), config);

        var plan = interpreter.Interpret("record screen and audio for 2 minutes");

        Assert.Single(plan.Steps);
        Assert.Equal(ActionType.RecordScreen, plan.Steps[0].Type);
        Assert.Equal(TimeSpan.FromMinutes(2), plan.Steps[0].WaitFor);
        Assert.Equal("audio:on", plan.Steps[0].Text);
    }

    [Fact]
    public void Interpret_UsesRuleBased_WhenFallbackModeAndRuleBasedIsRecognized()
    {
        var config = new AgentConfig { LlmFallbackEnabled = true, LlmInterpretationMode = "fallback" };
        var ruleBased = new RuleBasedIntentInterpreter(new StubAppResolver(), config);
        var interpreter = new FallbackIntentInterpreter(ruleBased, new StubRewriter("open notepad"), new StubAuditLog(), config);

        var plan = interpreter.Interpret("open calculator");

        Assert.Single(plan.Steps);
        Assert.Equal(ActionType.OpenApp, plan.Steps[0].Type);
        Assert.Equal("calculator", plan.Steps[0].AppIdOrPath);
    }

    [Fact]
    public void Interpret_UsesLlm_WhenFallbackModeAndRuleBasedIsUnrecognized()
    {
        var config = new AgentConfig { LlmFallbackEnabled = true, LlmInterpretationMode = "fallback" };
        var ruleBased = new RuleBasedIntentInterpreter(new StubAppResolver(), config);
        var interpreter = new FallbackIntentInterpreter(ruleBased, new StubRewriter("open calculator"), new StubAuditLog(), config);

        var plan = interpreter.Interpret("do the thing for me");

        Assert.Single(plan.Steps);
        Assert.Equal(ActionType.OpenApp, plan.Steps[0].Type);
        Assert.Equal("calculator", plan.Steps[0].AppIdOrPath);
    }

    [Fact]
    public void Interpret_AddsLowConfidenceMarker_WhenRewriteConfidenceIsBelowThreshold()
    {
        var config = new AgentConfig { LlmFallbackEnabled = true, LlmFallback = new LlmFallbackConfig { MinConfidence = 0.8 } };
        var ruleBased = new RuleBasedIntentInterpreter(new StubAppResolver(), config);
        var interpreter = new FallbackIntentInterpreter(
            ruleBased,
            new StubRewriter("open notepad", confidence: 0.42),
            new StubAuditLog(),
            config);

        var plan = interpreter.Interpret("open notepaad");

        Assert.Single(plan.Steps);
        Assert.Equal(ActionType.OpenApp, plan.Steps[0].Type);
        Assert.Contains("llm-low-confidence:0.42", plan.Steps[0].Note, StringComparison.OrdinalIgnoreCase);
    }

    private sealed class StubRewriter : ILlmIntentRewriter
    {
        private readonly string? _value;
        private readonly double _confidence;
        private readonly bool _needsClarification;
        private readonly string? _clarification;

        public StubRewriter(string? value, double confidence = 0.9, bool needsClarification = false, string? clarification = null)
        {
            _value = value;
            _confidence = confidence;
            _needsClarification = needsClarification;
            _clarification = clarification;
        }

        public LlmRewriteResult? Rewrite(string input)
        {
            return string.IsNullOrWhiteSpace(_value)
                ? null
                : new LlmRewriteResult(_value!, _confidence, _needsClarification, _clarification, _value);
        }
    }

    private sealed class StubAuditLog : IAuditLog
    {
        public Task WriteAsync(AuditEvent auditEvent, CancellationToken cancellationToken) => Task.CompletedTask;
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
