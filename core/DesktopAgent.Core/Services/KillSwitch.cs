using DesktopAgent.Core.Abstractions;

namespace DesktopAgent.Core.Services;

public sealed class KillSwitch : IKillSwitch
{
    private readonly object _sync = new();
    private bool _isTripped;
    private string? _reason;

    public bool IsTripped
    {
        get
        {
            lock (_sync)
            {
                return _isTripped;
            }
        }
    }

    public string? Reason
    {
        get
        {
            lock (_sync)
            {
                return _reason;
            }
        }
    }

    public void Trip(string reason)
    {
        lock (_sync)
        {
            _isTripped = true;
            _reason = string.IsNullOrWhiteSpace(reason) ? "Manual kill" : reason;
        }
    }

    public void Reset()
    {
        lock (_sync)
        {
            _isTripped = false;
            _reason = null;
        }
    }
}
