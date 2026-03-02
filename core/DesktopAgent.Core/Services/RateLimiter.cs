using DesktopAgent.Core.Abstractions;

namespace DesktopAgent.Core.Services;

public sealed class SlidingWindowRateLimiter : IRateLimiter
{
    private readonly Func<int> _maxActionsProvider;
    private readonly Queue<DateTimeOffset> _timestamps = new();
    private readonly object _lock = new();

    public SlidingWindowRateLimiter(int maxActionsPerSecond)
        : this(() => maxActionsPerSecond)
    {
    }

    public SlidingWindowRateLimiter(Func<int> maxActionsProvider)
    {
        _maxActionsProvider = maxActionsProvider ?? throw new ArgumentNullException(nameof(maxActionsProvider));
    }

    public bool TryAcquire()
    {
        var now = DateTimeOffset.UtcNow;
        var maxActionsPerSecond = Math.Max(1, _maxActionsProvider());
        lock (_lock)
        {
            while (_timestamps.Count > 0 && (now - _timestamps.Peek()).TotalSeconds >= 1)
            {
                _timestamps.Dequeue();
            }

            if (_timestamps.Count >= maxActionsPerSecond)
            {
                return false;
            }

            _timestamps.Enqueue(now);
            return true;
        }
    }
}
