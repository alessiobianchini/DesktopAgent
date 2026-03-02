using System.Text.Json;
using DesktopAgent.Core.Abstractions;
using DesktopAgent.Core.Config;
using DesktopAgent.Core.Models;

namespace DesktopAgent.Core.Services;

public sealed class JsonlAuditLog : IAuditLog
{
    private readonly string _path;
    private readonly JsonSerializerOptions _options = new() { WriteIndented = false };

    public JsonlAuditLog(AgentConfig config)
    {
        _path = config.AuditLogPath;
    }

    public Task WriteAsync(AuditEvent auditEvent, CancellationToken cancellationToken)
    {
        var line = JsonSerializer.Serialize(auditEvent, _options);
        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(_path)) ?? ".");
        return File.AppendAllTextAsync(_path, line + Environment.NewLine, cancellationToken);
    }
}
