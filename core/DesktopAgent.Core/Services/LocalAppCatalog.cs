using System.Runtime.InteropServices;
using System.Text.Json;
using DesktopAgent.Core.Abstractions;
using DesktopAgent.Core.Config;
using DesktopAgent.Core.Models;

namespace DesktopAgent.Core.Services;

public sealed class LocalAppCatalog : IAppCatalog
{
    private readonly AgentConfig _config;
    private readonly object _lock = new();
    private DateTimeOffset _lastRefresh = DateTimeOffset.MinValue;
    private List<AppEntry> _cache = new();

    public LocalAppCatalog(AgentConfig config)
    {
        _config = config;
    }

    public IReadOnlyList<AppEntry> GetAll()
    {
        lock (_lock)
        {
            if ((DateTimeOffset.UtcNow - _lastRefresh) < TimeSpan.FromMinutes(5) && _cache.Count > 0)
            {
                return _cache;
            }

            var diskCache = LoadPersistentCache();
            if (diskCache is { Count: > 0 })
            {
                _cache = diskCache;
                _lastRefresh = DateTimeOffset.UtcNow;
                return _cache;
            }

            _cache = BuildIndex();
            _lastRefresh = DateTimeOffset.UtcNow;
            SavePersistentCache(_cache);
            return _cache;
        }
    }

    private List<AppEntry>? LoadPersistentCache()
    {
        var path = GetCachePath();
        if (path == null || !File.Exists(path))
        {
            return null;
        }

        try
        {
            var json = File.ReadAllText(path);
            var cache = JsonSerializer.Deserialize<AppIndexCache>(json);
            if (cache == null || cache.Apps == null || cache.Apps.Count == 0)
            {
                return null;
            }

            if (!string.Equals(cache.Os, GetOsKey(), StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            if (_config.AppIndexCacheTtlMinutes > 0)
            {
                var age = DateTimeOffset.UtcNow - cache.CreatedAt;
                if (age.TotalMinutes > _config.AppIndexCacheTtlMinutes)
                {
                    return null;
                }
            }

            // A tiny Windows cache is often incomplete (e.g., only Edge) after first bootstrap.
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows) && cache.Apps.Count < 10)
            {
                return null;
            }

            // Invalidate old cache format that missed WindowsApps execution aliases.
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)
                && !cache.Apps.Any(app => app.Path.Contains(@"\WindowsApps\", StringComparison.OrdinalIgnoreCase)))
            {
                return null;
            }

            return cache.Apps;
        }
        catch
        {
            return null;
        }
    }

    private void SavePersistentCache(List<AppEntry> entries)
    {
        var path = GetCachePath();
        if (path == null)
        {
            return;
        }

        try
        {
            var directory = Path.GetDirectoryName(path);
            if (!string.IsNullOrWhiteSpace(directory))
            {
                Directory.CreateDirectory(directory);
            }

            var cache = new AppIndexCache
            {
                Os = GetOsKey(),
                CreatedAt = DateTimeOffset.UtcNow,
                Apps = entries
            };

            var json = JsonSerializer.Serialize(cache, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(path, json);
        }
        catch
        {
        }
    }

    private string? GetCachePath()
    {
        if (string.IsNullOrWhiteSpace(_config.AppIndexCachePath))
        {
            return null;
        }

        if (_config.AppIndexCacheTtlMinutes <= 0)
        {
            return null;
        }

        return Path.GetFullPath(_config.AppIndexCachePath, AppContext.BaseDirectory);
    }

    private static string GetOsKey()
    {
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            return "windows";
        }

        if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
        {
            return "macos";
        }

        if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
        {
            return "linux";
        }

        return "unknown";
    }

    private static List<AppEntry> BuildIndex()
    {
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            return BuildWindows();
        }

        if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
        {
            return BuildMac();
        }

        return BuildLinux();
    }

    private static List<AppEntry> BuildWindows()
    {
        var results = new List<AppEntry>();
        var folders = new[]
        {
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.CommonStartMenu), "Programs"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.StartMenu), "Programs")
        };

        foreach (var folder in folders)
        {
            if (!Directory.Exists(folder))
            {
                continue;
            }

            foreach (var file in Directory.EnumerateFiles(folder, "*.*", SearchOption.AllDirectories))
            {
                var ext = Path.GetExtension(file);
                if (!ext.Equals(".lnk", StringComparison.OrdinalIgnoreCase)
                    && !ext.Equals(".exe", StringComparison.OrdinalIgnoreCase)
                    && !ext.Equals(".appref-ms", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var name = Path.GetFileNameWithoutExtension(file);
                if (string.IsNullOrWhiteSpace(name))
                {
                    continue;
                }

                results.Add(new AppEntry(name, file));
            }
        }

        // Include executable aliases (WindowsApps) and common per-user installs.
        AddWindowsExecutables(results, Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Microsoft",
            "WindowsApps"), searchDepth: 0);
        AddWindowsExecutables(results, Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Programs"), searchDepth: 2);

        return Deduplicate(results);
    }

    private static void AddWindowsExecutables(List<AppEntry> results, string root, int searchDepth)
    {
        if (!Directory.Exists(root))
        {
            return;
        }

        foreach (var exe in EnumerateFilesLimited(root, "*.exe", searchDepth))
        {
            var name = Path.GetFileNameWithoutExtension(exe);
            if (string.IsNullOrWhiteSpace(name))
            {
                continue;
            }

            results.Add(new AppEntry(name, exe));
        }
    }

    private static IEnumerable<string> EnumerateFilesLimited(string root, string pattern, int maxDepth)
    {
        var queue = new Queue<(string Dir, int Depth)>();
        queue.Enqueue((root, 0));

        while (queue.Count > 0)
        {
            var (dir, depth) = queue.Dequeue();
            IEnumerable<string> files;
            try
            {
                files = Directory.EnumerateFiles(dir, pattern, SearchOption.TopDirectoryOnly);
            }
            catch
            {
                continue;
            }

            foreach (var file in files)
            {
                yield return file;
            }

            if (depth >= maxDepth)
            {
                continue;
            }

            IEnumerable<string> subdirs;
            try
            {
                subdirs = Directory.EnumerateDirectories(dir);
            }
            catch
            {
                continue;
            }

            foreach (var subdir in subdirs)
            {
                queue.Enqueue((subdir, depth + 1));
            }
        }
    }

    private static List<AppEntry> BuildMac()
    {
        var results = new List<AppEntry>();
        var home = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
        var roots = new[]
        {
            "/Applications",
            "/System/Applications",
            Path.Combine(home, "Applications")
        };

        foreach (var root in roots)
        {
            if (!Directory.Exists(root))
            {
                continue;
            }

            foreach (var app in EnumerateAppBundles(root, maxDepth: 2))
            {
                var name = Path.GetFileNameWithoutExtension(app);
                if (string.IsNullOrWhiteSpace(name))
                {
                    continue;
                }
                results.Add(new AppEntry(name, app));
            }
        }

        return Deduplicate(results);
    }

    private static IEnumerable<string> EnumerateAppBundles(string root, int maxDepth)
    {
        var queue = new Queue<(string Path, int Depth)>();
        queue.Enqueue((root, 0));

        while (queue.Count > 0)
        {
            var (path, depth) = queue.Dequeue();
            if (depth > maxDepth)
            {
                continue;
            }

            IEnumerable<string> subdirs;
            try
            {
                subdirs = Directory.EnumerateDirectories(path);
            }
            catch
            {
                continue;
            }

            foreach (var dir in subdirs)
            {
                if (dir.EndsWith(".app", StringComparison.OrdinalIgnoreCase))
                {
                    yield return dir;
                    continue;
                }

                if (depth + 1 <= maxDepth)
                {
                    queue.Enqueue((dir, depth + 1));
                }
            }
        }
    }

    private static List<AppEntry> BuildLinux()
    {
        var results = new List<AppEntry>();
        var home = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
        var dirs = new[]
        {
            "/usr/share/applications",
            "/usr/local/share/applications",
            Path.Combine(home, ".local", "share", "applications")
        };

        foreach (var dir in dirs)
        {
            if (!Directory.Exists(dir))
            {
                continue;
            }

            foreach (var file in Directory.EnumerateFiles(dir, "*.desktop", SearchOption.TopDirectoryOnly))
            {
                var entry = ParseDesktopFile(file);
                if (entry != null)
                {
                    results.Add(entry);
                }
            }
        }

        return Deduplicate(results);
    }

    private static AppEntry? ParseDesktopFile(string path)
    {
        try
        {
            var inEntry = false;
            string? name = null;
            string? exec = null;
            var noDisplay = false;

            foreach (var raw in File.ReadLines(path))
            {
                var line = raw.Trim();
                if (line.StartsWith("[") && line.EndsWith("]"))
                {
                    inEntry = line.Equals("[Desktop Entry]", StringComparison.OrdinalIgnoreCase);
                    continue;
                }

                if (!inEntry || line.StartsWith("#", StringComparison.Ordinal))
                {
                    continue;
                }

                if (line.StartsWith("NoDisplay=", StringComparison.OrdinalIgnoreCase))
                {
                    noDisplay = line.EndsWith("true", StringComparison.OrdinalIgnoreCase);
                }

                if (line.StartsWith("Hidden=", StringComparison.OrdinalIgnoreCase))
                {
                    noDisplay = noDisplay || line.EndsWith("true", StringComparison.OrdinalIgnoreCase);
                }

                if (line.StartsWith("Name=", StringComparison.OrdinalIgnoreCase) && name == null)
                {
                    name = line.Substring("Name=".Length).Trim();
                }

                if (line.StartsWith("Exec=", StringComparison.OrdinalIgnoreCase) && exec == null)
                {
                    exec = line.Substring("Exec=".Length).Trim();
                }
            }

            if (noDisplay || string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(exec))
            {
                return null;
            }

            var command = CleanDesktopExec(exec);
            if (string.IsNullOrWhiteSpace(command))
            {
                return null;
            }

            return new AppEntry(name, command);
        }
        catch
        {
            return null;
        }
    }

    private static string? CleanDesktopExec(string exec)
    {
        var parts = exec.Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Where(part => !part.StartsWith("%", StringComparison.Ordinal))
            .ToList();

        if (parts.Count == 0)
        {
            return null;
        }

        if (parts[0].Equals("env", StringComparison.OrdinalIgnoreCase))
        {
            var idx = parts.FindIndex(part => !part.Contains('=') && !part.Equals("env", StringComparison.OrdinalIgnoreCase));
            if (idx >= 0)
            {
                parts = parts.Skip(idx).ToList();
            }
        }

        return parts.Count > 0 ? parts[0] : null;
    }

    private static List<AppEntry> Deduplicate(List<AppEntry> entries)
    {
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var results = new List<AppEntry>();

        foreach (var entry in entries)
        {
            var key = $"{entry.Name}|{entry.Path}";
            if (seen.Add(key))
            {
                results.Add(entry);
            }
        }

        return results;
    }

    private sealed class AppIndexCache
    {
        public string Os { get; set; } = "";
        public DateTimeOffset CreatedAt { get; set; } = DateTimeOffset.UtcNow;
        public List<AppEntry> Apps { get; set; } = new();
    }
}
