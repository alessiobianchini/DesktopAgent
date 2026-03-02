using DesktopAgent.Core.Config;
using DesktopAgent.Core.Models;
using DesktopAgent.Core.Services;
using DesktopAgent.Proto;
using Xunit;

namespace DesktopAgent.Tests;

public sealed class PolicyEngineTests
{
    [Fact]
    public void BlocksSubmitWithoutConfirmation()
    {
        var config = new AgentConfig
        {
            RequireConfirmation = true,
            BlockedActionsKeywords = new List<string> { "submit" }
        };
        var engine = new PolicyEngine(config);

        var step = new PlanStep
        {
            Type = ActionType.Click,
            Selector = new Selector { NameContains = "Submit" }
        };

        var decision = engine.Evaluate(step, new WindowRef { Title = "Form", AppId = "notepad" });
        Assert.True(decision.Allowed);
        Assert.True(decision.RequiresConfirmation);
    }

    [Fact]
    public void EnforcesAllowlist()
    {
        var config = new AgentConfig
        {
            AllowedApps = new List<string> { "notepad" }
        };
        var engine = new PolicyEngine(config);

        var step = new PlanStep { Type = ActionType.TypeText, Text = "hello" };
        var decision = engine.Evaluate(step, new WindowRef { Title = "Calculator", AppId = "calc" });
        Assert.False(decision.Allowed);
    }
}

public sealed class RateLimiterTests
{
    [Fact]
    public void RateLimiterBlocksAfterLimit()
    {
        var limiter = new SlidingWindowRateLimiter(3);
        Assert.True(limiter.TryAcquire());
        Assert.True(limiter.TryAcquire());
        Assert.True(limiter.TryAcquire());
        Assert.False(limiter.TryAcquire());
    }
}
