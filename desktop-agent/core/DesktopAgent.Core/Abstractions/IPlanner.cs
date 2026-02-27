using DesktopAgent.Core.Models;

namespace DesktopAgent.Core.Abstractions;

public interface IPlanner
{
    ActionPlan PlanFromIntent(string intent);
}
