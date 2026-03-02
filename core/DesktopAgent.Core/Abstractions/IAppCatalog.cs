using DesktopAgent.Core.Models;

namespace DesktopAgent.Core.Abstractions;

public interface IAppCatalog
{
    IReadOnlyList<AppEntry> GetAll();
}
