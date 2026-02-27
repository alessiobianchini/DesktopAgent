using DesktopAgent.Core.Models;

namespace DesktopAgent.Core.Abstractions;

public interface IContextProvider
{
    Task<ContextSnapshot> GetSnapshotAsync(CancellationToken cancellationToken);
    Task<FindResult> FindByTextAsync(string text, CancellationToken cancellationToken);
}
