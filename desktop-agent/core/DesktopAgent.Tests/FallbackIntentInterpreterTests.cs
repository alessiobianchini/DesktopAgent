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

    private sealed class StubRewriter : ILlmIntentRewriter
    {
        private readonly string? _value;

        public StubRewriter(string? value)
        {
            _value = value;
        }

        public string? Rewrite(string input) => _value;
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
