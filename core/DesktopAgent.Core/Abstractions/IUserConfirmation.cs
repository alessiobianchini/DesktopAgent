namespace DesktopAgent.Core.Abstractions;

public interface IUserConfirmation
{
    Task<bool> ConfirmAsync(string message, CancellationToken cancellationToken);
}
