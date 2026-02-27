using DesktopAgent.Core.Models;
using DesktopAgent.Proto;

namespace DesktopAgent.Core.Abstractions;

public interface IPolicyEngine
{
    PolicyDecision Evaluate(PlanStep step, WindowRef? activeWindow);
}
