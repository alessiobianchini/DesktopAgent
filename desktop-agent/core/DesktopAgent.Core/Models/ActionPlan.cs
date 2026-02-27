namespace DesktopAgent.Core.Models;

public sealed class ActionPlan
{
    public string? Intent { get; set; }
    public List<PlanStep> Steps { get; set; } = new();
}
