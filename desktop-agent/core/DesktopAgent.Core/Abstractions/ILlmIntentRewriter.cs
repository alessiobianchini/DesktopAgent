namespace DesktopAgent.Core.Abstractions;

public interface ILlmIntentRewriter
{
    string? Rewrite(string input);
}
