using DesktopAgent.Core.Models;

namespace DesktopAgent.Core.Abstractions;

public interface IAuditLog
{
    Task WriteAsync(AuditEvent auditEvent, CancellationToken cancellationToken);
}
