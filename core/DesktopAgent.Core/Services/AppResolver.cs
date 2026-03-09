using System.Text.RegularExpressions;
using System.Runtime.InteropServices;
using DesktopAgent.Core.Abstractions;
using DesktopAgent.Core.Config;
using DesktopAgent.Core.Models;

namespace DesktopAgent.Core.Services;

public sealed class AppResolver : IAppResolver
{
    private const double ConfidenceThreshold = 0.72;
    private static readonly string[] BrowserAliasKeys = { "google chrome", "microsoft edge", "chrome", "edge", "firefox", "brave", "opera", "safari" };
    private static readonly HashSet<string> BrowserCommandTokens = new(StringComparer.OrdinalIgnoreCase)
    {
        "and", "then", "e", "poi", "quindi",
        "go", "to", "navigate", "search", "find", "open",
        "vai", "naviga", "cerca", "trova", "apri",
        "url", "website", "site", "web", "browser",
        "on", "in", "su", "nel", "nella"
    };
    private static readonly IReadOnlyDictionary<string, string> BuiltInAliases = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
    {
        ["chrome"] = "chrome",
        ["google chrome"] = "chrome",
        ["edge"] = "msedge",
        ["microsoft edge"] = "msedge",
        ["firefox"] = "firefox",
        ["brave"] = "brave",
        ["opera"] = "opera",
        ["safari"] = "safari",
        ["notepad++"] = "notepad++",
        ["notepad plus plus"] = "notepad++",
        ["notepadplusplus"] = "notepad++",
        ["npp"] = "notepad++",
        ["teams"] = "ms-teams",
        ["microsoft teams"] = "ms-teams",
        ["ms teams"] = "ms-teams",
        ["msteams"] = "ms-teams",
        ["new teams"] = "ms-teams",
        ["vscode"] = "code",
        ["vs code"] = "code",
        ["visual studio code"] = "code"
    };

    private readonly AgentConfig _config;
    private readonly IAppCatalog _catalog;
    public AppResolver(AgentConfig config, IAppCatalog catalog)
    {
        _config = config;
        _catalog = catalog;
    }

    public bool TryResolveApp(string input, out string resolved)
    {
        resolved = input ?? string.Empty;
        if (string.IsNullOrWhiteSpace(resolved))
        {
            return false;
        }

        if (LooksLikePath(resolved))
        {
            return true;
        }

        var normalized = Normalize(resolved);
        if (string.IsNullOrWhiteSpace(normalized))
        {
            return false;
        }

        if (_config.AppAliases.TryGetValue(normalized, out var alias))
        {
            if (TryResolveExecutablePath(alias, out var executablePath))
            {
                resolved = executablePath;
                return true;
            }

            resolved = alias;
            return true;
        }

        if (BuiltInAliases.TryGetValue(normalized, out var builtIn))
        {
            if (TryResolveExecutablePath(builtIn, out var executablePath))
            {
                resolved = executablePath;
                return true;
            }

            resolved = builtIn;
            return true;
        }

        if (TryResolveBuiltInAliasFromCommandLikeInput(normalized, out builtIn))
        {
            if (TryResolveExecutablePath(builtIn, out var executablePath))
            {
                resolved = executablePath;
                return true;
            }

            resolved = builtIn;
            return true;
        }

        var matches = BuildMatches(normalized);
        if (matches.Count == 0)
        {
            return false;
        }

        var allowed = RankMatches(matches.Where(m => m.IsAllowed), normalized).FirstOrDefault();
        if (allowed != null && allowed.Score >= ConfidenceThreshold)
        {
            resolved = allowed.Entry.Path;
            return true;
        }

        var best = RankMatches(matches, normalized).First();
        if (best.Score >= ConfidenceThreshold)
        {
            resolved = best.Entry.Path;
            return true;
        }

        return false;
    }

    public IReadOnlyList<AppMatch> Suggest(string input, int maxResults)
    {
        maxResults = Math.Clamp(maxResults, 1, 50);
        var normalized = Normalize(input ?? string.Empty);
        var apps = _catalog.GetAll();

        if (string.IsNullOrWhiteSpace(normalized))
        {
            return apps
                .OrderBy(a => a.Name, StringComparer.OrdinalIgnoreCase)
                .Take(maxResults)
                .Select(a => new AppMatch(a, 1.0, IsAllowed(a)))
                .ToList();
        }

        var matches = BuildMatches(normalized);
        return RankMatches(matches, normalized)
            .Where(m => m.Score >= 0.35)
            .Take(maxResults)
            .ToList();
    }

    private List<AppMatch> BuildMatches(string normalizedQuery)
    {
        var apps = _catalog.GetAll();
        var results = new List<AppMatch>(apps.Count);
        foreach (var app in apps)
        {
            var score = ScoreMatch(normalizedQuery, app);
            if (score > 0)
            {
                results.Add(new AppMatch(app, score, IsAllowed(app)));
            }
        }
        return results;
    }

    private bool IsAllowed(AppEntry entry)
    {
        if (_config.AllowedApps.Count == 0)
        {
            return true;
        }

        return _config.AllowedApps.Any(app => entry.Name.Contains(app, StringComparison.OrdinalIgnoreCase)
                                              || entry.Path.Contains(app, StringComparison.OrdinalIgnoreCase));
    }

    private static double ScoreMatch(string normalizedQuery, AppEntry entry)
    {
        var nameScore = ScoreMatch(normalizedQuery, entry.Name);
        var pathScore = ScoreMatch(normalizedQuery, GetFileNameWithoutExtensionCrossPlatform(entry.Path));
        return Math.Max(nameScore, pathScore);
    }

    private static double ScoreMatch(string normalizedQuery, string candidate)
    {
        var normalizedCandidate = Normalize(candidate);
        if (string.IsNullOrWhiteSpace(normalizedCandidate))
        {
            return 0.0;
        }

        if (normalizedCandidate.Equals(normalizedQuery, StringComparison.OrdinalIgnoreCase))
        {
            return 1.0;
        }

        if (normalizedCandidate.StartsWith(normalizedQuery, StringComparison.OrdinalIgnoreCase))
        {
            return 0.9;
        }

        if (normalizedCandidate.Contains(normalizedQuery, StringComparison.OrdinalIgnoreCase))
        {
            return 0.8;
        }

        var queryTokens = Tokenize(normalizedQuery);
        var candidateTokens = Tokenize(normalizedCandidate);
        if (queryTokens.Count == 0 || candidateTokens.Count == 0)
        {
            return 0.0;
        }

        var intersection = queryTokens.Intersect(candidateTokens, StringComparer.OrdinalIgnoreCase).Count();
        var overlap = intersection / (double)Math.Max(queryTokens.Count, candidateTokens.Count);
        var allTokensPresent = intersection == queryTokens.Count;
        var tokenScore = allTokensPresent ? 0.85 : overlap;

        var acronym = GetInitialism(normalizedCandidate);
        if (normalizedQuery.Length <= 4 && acronym.StartsWith(normalizedQuery, StringComparison.OrdinalIgnoreCase))
        {
            tokenScore = Math.Max(tokenScore, 0.75);
        }

        return tokenScore;
    }

    private static IOrderedEnumerable<AppMatch> RankMatches(IEnumerable<AppMatch> matches, string normalizedQuery)
    {
        return matches
            .OrderByDescending(m => m.Score)
            .ThenBy(m => GetTokenDistance(normalizedQuery, m.Entry))
            .ThenBy(m => GetNameLength(m.Entry))
            .ThenBy(m => m.Entry.Name, StringComparer.OrdinalIgnoreCase);
    }

    private static int GetTokenDistance(string normalizedQuery, AppEntry entry)
    {
        var queryTokens = Tokenize(normalizedQuery);
        if (queryTokens.Count == 0)
        {
            return int.MaxValue;
        }

        var nameTokens = Tokenize(Normalize(entry.Name));
        var pathTokens = Tokenize(Normalize(GetFileNameWithoutExtensionCrossPlatform(entry.Path)));
        var candidateTokens = nameTokens.Count >= pathTokens.Count ? nameTokens : pathTokens;
        if (candidateTokens.Count == 0)
        {
            return int.MaxValue;
        }

        var missing = queryTokens.Count(token => !candidateTokens.Contains(token));
        var extras = Math.Max(0, candidateTokens.Count - queryTokens.Count);
        return missing * 10 + extras;
    }

    private static int GetNameLength(AppEntry entry)
    {
        var normalized = Normalize(entry.Name);
        return string.IsNullOrWhiteSpace(normalized) ? int.MaxValue : normalized.Length;
    }

    private bool TryResolveExecutablePath(string executableToken, out string executablePath)
    {
        executablePath = string.Empty;
        var expectedFileName = executableToken.EndsWith(".exe", StringComparison.OrdinalIgnoreCase)
            ? executableToken
            : $"{executableToken}.exe";

        var apps = _catalog.GetAll();

        var exactExecutable = apps.FirstOrDefault(entry =>
            string.Equals(GetFileNameCrossPlatform(entry.Path), expectedFileName, StringComparison.OrdinalIgnoreCase));
        if (exactExecutable != null)
        {
            executablePath = exactExecutable.Path;
            return true;
        }

        var normalizedExecutable = Normalize(Path.GetFileNameWithoutExtension(expectedFileName));
        var exactByName = apps.FirstOrDefault(entry =>
            string.Equals(Normalize(GetFileNameWithoutExtensionCrossPlatform(entry.Path)), normalizedExecutable, StringComparison.OrdinalIgnoreCase));
        if (exactByName != null)
        {
            executablePath = exactByName.Path;
            return true;
        }

        if (TryResolveExecutableFromPath(expectedFileName, out var fromPath))
        {
            executablePath = fromPath;
            return true;
        }

        return false;
    }

    private static bool TryResolveExecutableFromPath(string expectedFileName, out string executablePath)
    {
        executablePath = string.Empty;
        var pathValue = Environment.GetEnvironmentVariable("PATH");
        if (string.IsNullOrWhiteSpace(pathValue))
        {
            return false;
        }

        var separators = RuntimeInformation.IsOSPlatform(OSPlatform.Windows) ? ';' : ':';
        var paths = pathValue.Split(separators, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        foreach (var path in paths)
        {
            try
            {
                var candidate = Path.Combine(path.Trim('"'), expectedFileName);
                if (File.Exists(candidate))
                {
                    executablePath = candidate;
                    return true;
                }
            }
            catch
            {
                // Ignore malformed PATH entries.
            }
        }

        return false;
    }

    private static bool TryResolveBuiltInAliasFromCommandLikeInput(string normalized, out string alias)
    {
        alias = string.Empty;
        foreach (var key in BrowserAliasKeys)
        {
            if (!normalized.StartsWith($"{key} ", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            var tail = normalized[(key.Length + 1)..];
            var tailTokens = Tokenize(tail);
            if (tailTokens.Count == 0)
            {
                continue;
            }

            if (!tailTokens.Any(token => BrowserCommandTokens.Contains(token)))
            {
                continue;
            }

            if (BuiltInAliases.TryGetValue(key, out var resolved))
            {
                alias = resolved;
                return true;
            }
        }

        return false;
    }

    private static string Normalize(string input)
    {
        var value = (input ?? string.Empty).Trim().Trim('"', '\'');
        value = Regex.Replace(value, "\\s+", " ");
        var lower = value.ToLowerInvariant();
        lower = lower.Replace("per favore", string.Empty)
                     .Replace("perfavore", string.Empty)
                     .Replace("please", string.Empty)
                     .Replace("application", string.Empty)
                     .Replace("applicazione", string.Empty)
                     .Replace("programma", string.Empty)
                     .Replace("program", string.Empty)
                     .Replace("app", string.Empty)
                     .Replace("il ", string.Empty)
                     .Replace("la ", string.Empty)
                     .Replace("lo ", string.Empty);
        lower = Regex.Replace(lower, "[^a-z0-9 ]", " ");
        return Regex.Replace(lower, "\\s+", " ").Trim();
    }

    private static HashSet<string> Tokenize(string input)
    {
        return input.Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
    }

    private static string GetInitialism(string input)
    {
        var parts = input.Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        if (parts.Length == 0)
        {
            return string.Empty;
        }

        var chars = parts.Select(p => p[0]);
        return new string(chars.ToArray());
    }

    private static bool LooksLikePath(string input)
    {
        if (input.Contains(Path.DirectorySeparatorChar)
            || input.Contains(Path.AltDirectorySeparatorChar)
            || input.Contains(":\\", StringComparison.OrdinalIgnoreCase)
            || input.StartsWith("/", StringComparison.OrdinalIgnoreCase)
            || input.StartsWith("~/", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        return input.EndsWith(".exe", StringComparison.OrdinalIgnoreCase)
            || input.EndsWith(".lnk", StringComparison.OrdinalIgnoreCase)
            || input.EndsWith(".appref-ms", StringComparison.OrdinalIgnoreCase)
            || input.EndsWith(".app", StringComparison.OrdinalIgnoreCase);
    }

    private static string GetFileNameCrossPlatform(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return string.Empty;
        }

        var normalized = path.Replace('\\', '/');
        return Path.GetFileName(normalized);
    }

    private static string GetFileNameWithoutExtensionCrossPlatform(string path)
    {
        var fileName = GetFileNameCrossPlatform(path);
        return Path.GetFileNameWithoutExtension(fileName);
    }
}
