using DesktopAgent.Core.Abstractions;
using DesktopAgent.Core.Models;

namespace DesktopAgent.Core.Services;

public sealed class SimplePlanner : IPlanner
{
    private readonly IIntentInterpreter _interpreter;

    public SimplePlanner(IIntentInterpreter interpreter)
    {
        _interpreter = interpreter;
    }

    public ActionPlan PlanFromIntent(string intent)
    {
        return _interpreter.Interpret(intent);
    }
}
