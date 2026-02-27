using System.Collections.Concurrent;
using FlaUI.Core.AutomationElements;
using FlaUI.UIA3;

namespace DesktopAgent.Adapter.Windows;

public sealed class AdapterState : IDisposable
{
    private readonly ConcurrentDictionary<string, AutomationElement> _elements = new();
    private readonly ConcurrentDictionary<string, AutomationElement> _windows = new();

    public UIA3Automation Automation { get; } = new();

    public bool Armed { get; private set; }
    public bool RequireUserPresence { get; private set; }

    public string RememberWindow(AutomationElement element)
    {
        var id = Guid.NewGuid().ToString("n");
        _windows[id] = element;
        _elements[id] = element;
        return id;
    }

    public string RememberElement(AutomationElement element)
    {
        var id = Guid.NewGuid().ToString("n");
        _elements[id] = element;
        return id;
    }

    public AutomationElement? GetWindow(string id)
    {
        return _windows.TryGetValue(id, out var element) ? element : null;
    }

    public AutomationElement? GetElement(string id)
    {
        return _elements.TryGetValue(id, out var element) ? element : null;
    }

    public void Arm(bool requireUserPresence)
    {
        Armed = true;
        RequireUserPresence = requireUserPresence;
    }

    public void Disarm()
    {
        Armed = false;
        RequireUserPresence = false;
    }

    public void Dispose()
    {
        Automation.Dispose();
    }
}
