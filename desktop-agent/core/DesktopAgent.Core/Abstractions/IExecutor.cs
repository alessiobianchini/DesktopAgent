using DesktopAgent.Core.Models;

namespace DesktopAgent.Core.Abstractions;

public interface IExecutor
{
    Task<ExecutionResult> ExecutePlanAsync(ActionPlan plan, bool dryRun, CancellationToken cancellationToken);
}
