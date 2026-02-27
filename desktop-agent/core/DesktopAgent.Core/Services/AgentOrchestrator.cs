using DesktopAgent.Core.Abstractions;
using DesktopAgent.Core.Models;
using Microsoft.Extensions.Logging;

namespace DesktopAgent.Core.Services;

public sealed class AgentOrchestrator
{
    private readonly IPlanner _planner;
    private readonly IExecutor _executor;
    private readonly ILogger<AgentOrchestrator> _logger;

    public AgentOrchestrator(IPlanner planner, IExecutor executor, ILogger<AgentOrchestrator> logger)
    {
        _planner = planner;
        _executor = executor;
        _logger = logger;
    }

    public Task<ExecutionResult> ExecuteIntentAsync(string intent, bool dryRun, CancellationToken cancellationToken)
    {
        var plan = _planner.PlanFromIntent(intent);
        _logger.LogInformation("Planned {Count} steps for intent: {Intent}", plan.Steps.Count, intent);
        return _executor.ExecutePlanAsync(plan, dryRun, cancellationToken);
    }
}
