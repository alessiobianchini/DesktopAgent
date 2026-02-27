namespace DesktopAgent.Core.Models;

public sealed class PolicyDecision
{
    public bool Allowed { get; set; }
    public bool RequiresConfirmation { get; set; }
    public string Reason { get; set; } = string.Empty;

    public static PolicyDecision Allow() => new() { Allowed = true };
    public static PolicyDecision Deny(string reason) => new() { Allowed = false, Reason = reason };
    public static PolicyDecision RequireConfirmation(string reason) => new() { Allowed = true, RequiresConfirmation = true, Reason = reason };
}
