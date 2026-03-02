using DesktopAgent.Core.Models;

namespace DesktopAgent.Core.Abstractions;

public interface IIntentInterpreter
{
    ActionPlan Interpret(string intent);
}
