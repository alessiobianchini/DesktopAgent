using DesktopAgent.Core.Models;

namespace DesktopAgent.Core.Abstractions;

public interface IAppResolver
{
    bool TryResolveApp(string input, out string resolved);
    IReadOnlyList<AppMatch> Suggest(string input, int maxResults);
}
