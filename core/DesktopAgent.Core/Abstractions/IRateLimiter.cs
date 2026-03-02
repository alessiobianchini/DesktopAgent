namespace DesktopAgent.Core.Abstractions;

public interface IRateLimiter
{
    bool TryAcquire();
}
