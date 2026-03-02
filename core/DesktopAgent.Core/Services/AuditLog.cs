using System.Text.Json;
using DesktopAgent.Core.Abstractions;
using DesktopAgent.Core.Config;
using DesktopAgent.Core.Models;

namespace DesktopAgent.Core.Services;

public sealed class JsonlAuditLog : IAuditLog
{
    private readonly string _fallbackPath;
    private string _activePath;
    private readonly JsonSerializerOptions _options = new() { WriteIndented = false };

    public JsonlAuditLog(AgentConfig config)
    {
        var configured = string.IsNullOrWhiteSpace(config.AuditLogPath)
            ? "audit.log.jsonl"
            : Environment.ExpandEnvironmentVariables(config.AuditLogPath.Trim());

        var candidatePath = Path.IsPathRooted(configured)
            ? configured
            : Path.GetFullPath(configured);

        var fileName = Path.GetFileName(candidatePath);
        if (string.IsNullOrWhiteSpace(fileName))
        {
            fileName = "audit.log.jsonl";
        }

        _fallbackPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "DesktopAgent",
            "logs",
            fileName);

        _activePath = candidatePath;
    }

    public Task WriteAsync(AuditEvent auditEvent, CancellationToken cancellationToken)
    {
        var line = JsonSerializer.Serialize(auditEvent, _options);
        return WriteWithFallbackAsync(line + Environment.NewLine, cancellationToken);
    }

    private async Task WriteWithFallbackAsync(string content, CancellationToken cancellationToken)
    {
        try
        {
            await WriteToPathAsync(_activePath, content, cancellationToken);
        }
        catch (UnauthorizedAccessException)
        {
            _activePath = _fallbackPath;
            await WriteToPathAsync(_activePath, content, cancellationToken);
        }
    }

    private static Task WriteToPathAsync(string path, string content, CancellationToken cancellationToken)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path)) ?? ".");
        return File.AppendAllTextAsync(path, content, cancellationToken);
    }
}
