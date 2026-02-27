namespace DesktopAgent.Core.Abstractions;

public interface IKillSwitch
{
    bool IsTripped { get; }
    void Trip(string reason);
    void Reset();
    string? Reason { get; }
}
