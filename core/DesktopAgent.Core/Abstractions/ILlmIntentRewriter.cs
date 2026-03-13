namespace DesktopAgent.Core.Abstractions;

public sealed record LlmRewriteResult(
    string Command,
    double Confidence = 0.5,
    bool NeedsClarification = false,
    string? ClarificationQuestion = null,
    string? RawOutput = null);

public interface ILlmIntentRewriter
{
    LlmRewriteResult? Rewrite(string input);
}
