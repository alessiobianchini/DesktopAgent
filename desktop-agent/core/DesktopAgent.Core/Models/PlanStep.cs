using DesktopAgent.Proto;

namespace DesktopAgent.Core.Models;

public sealed class PlanStep
{
    public ActionType Type { get; set; }
    public Selector? Selector { get; set; }
    public string? ExpectedAppId { get; set; }
    public string? ExpectedWindowId { get; set; }
    public string? Text { get; set; }
    public string? Target { get; set; }
    public string? AppIdOrPath { get; set; }
    public List<string>? Keys { get; set; }
    public Rect? Point { get; set; }
    public string? ElementId { get; set; }
    public TimeSpan? WaitFor { get; set; }
    public string? Note { get; set; }
}
