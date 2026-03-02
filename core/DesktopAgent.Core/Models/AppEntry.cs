namespace DesktopAgent.Core.Models;

public sealed record AppEntry(string Name, string Path);

public sealed record AppMatch(AppEntry Entry, double Score, bool IsAllowed);
