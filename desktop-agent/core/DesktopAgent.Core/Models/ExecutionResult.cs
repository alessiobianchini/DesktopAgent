namespace DesktopAgent.Core.Models;

public sealed class ExecutionResult
{
    public bool Success { get; set; }
    public string Message { get; set; } = string.Empty;
    public List<StepResult> Steps { get; set; } = new();
}

public sealed class StepResult
{
    public int Index { get; set; }
    public ActionType Type { get; set; }
    public bool Success { get; set; }
    public string Message { get; set; } = string.Empty;
    public object? Data { get; set; }
}
