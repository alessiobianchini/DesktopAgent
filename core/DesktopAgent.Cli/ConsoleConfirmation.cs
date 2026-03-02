using DesktopAgent.Core.Abstractions;

namespace DesktopAgent.Cli;

public sealed class ConsoleConfirmation : IUserConfirmation
{
    public Task<bool> ConfirmAsync(string message, CancellationToken cancellationToken)
    {
        Console.Write($"Confirmation required: {message}. Proceed? (y/N): ");
        var input = Console.ReadLine();
        var confirmed = string.Equals(input, "y", StringComparison.OrdinalIgnoreCase)
                        || string.Equals(input, "yes", StringComparison.OrdinalIgnoreCase);
        return Task.FromResult(confirmed);
    }
}
