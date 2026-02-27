using DesktopAgent.Core.Abstractions;
using DesktopAgent.Core.Config;
using DesktopAgent.Core.Models;

namespace DesktopAgent.Core.Services;

public sealed class FallbackIntentInterpreter : IIntentInterpreter
{
    private readonly RuleBasedIntentInterpreter _ruleBased;
    private readonly ILlmIntentRewriter _rewriter;
    private readonly IAuditLog _auditLog;
    private readonly AgentConfig _config;

    public FallbackIntentInterpreter(RuleBasedIntentInterpreter ruleBased, ILlmIntentRewriter rewriter, IAuditLog auditLog, AgentConfig config)
    {
        _ruleBased = ruleBased;
        _rewriter = rewriter;
        _auditLog = auditLog;
        _config = config;
    }

    public ActionPlan Interpret(string intent)
    {
        if (!_config.LlmFallbackEnabled)
        {
            return _ruleBased.Interpret(intent);
        }

        var rewritten = _rewriter.Rewrite(intent);
        if (!string.IsNullOrWhiteSpace(rewritten))
        {
            var rewrittenPlan = _ruleBased.Interpret(rewritten);
            if (!ShouldFallback(rewrittenPlan))
            {
                PreserveMouseJiggleDuration(intent, rewrittenPlan);
                rewrittenPlan.Intent = intent;
                if (rewrittenPlan.Steps.Count > 0 && string.IsNullOrWhiteSpace(rewrittenPlan.Steps[0].Note))
                {
                    rewrittenPlan.Steps[0].Note = $"Rewritten intent: {rewritten}";
                }

                WriteAudit("llm_rewrite_applied", "LLM rewrite applied", new
                {
                    input = ToAuditText(intent),
                    rewritten = ToAuditText(rewritten),
                    stepCount = rewrittenPlan.Steps.Count
                });
                return rewrittenPlan;
            }

            WriteAudit("llm_fallback_rule_based", "LLM rewrite unusable; using rule-based plan", new
            {
                input = ToAuditText(intent),
                rewritten = ToAuditText(rewritten)
            });
            return _ruleBased.Interpret(intent);
        }

        WriteAudit("llm_fallback_rule_based", "LLM rewrite unavailable; using rule-based plan", new
        {
            input = ToAuditText(intent)
        });
        return _ruleBased.Interpret(intent);
    }

    private static bool ShouldFallback(ActionPlan plan)
    {
        if (plan.Steps.Count == 0)
        {
            return true;
        }

        if (plan.Steps.Count == 1 && plan.Steps[0].Type == ActionType.ReadText)
        {
            return true;
        }

        if (plan.Steps.Any(step => step.Type == ActionType.ReadText
                                   && !string.IsNullOrWhiteSpace(step.Note)
                                   && step.Note.StartsWith("Unrecognized", StringComparison.OrdinalIgnoreCase)))
        {
            return true;
        }

        return false;
    }

    private void PreserveMouseJiggleDuration(string originalIntent, ActionPlan rewrittenPlan)
    {
        var rewrittenMouse = rewrittenPlan.Steps.FirstOrDefault(step => step.Type == ActionType.MouseJiggle);
        if (rewrittenMouse == null)
        {
            return;
        }

        var rewrittenDuration = rewrittenMouse.WaitFor ?? TimeSpan.FromSeconds(30);
        if (rewrittenDuration != TimeSpan.FromSeconds(30))
        {
            return;
        }

        var originalPlan = _ruleBased.Interpret(originalIntent);
        var originalMouse = originalPlan.Steps.FirstOrDefault(step => step.Type == ActionType.MouseJiggle);
        if (originalMouse?.WaitFor is not TimeSpan originalDuration)
        {
            return;
        }

        if (originalDuration <= TimeSpan.Zero || originalDuration == TimeSpan.FromSeconds(30))
        {
            return;
        }

        rewrittenMouse.WaitFor = originalDuration;
    }

    private string ToAuditText(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw))
        {
            return string.Empty;
        }

        return _config.AuditLlmIncludeRawText ? raw : "[redacted]";
    }

    private void WriteAudit(string eventType, string message, object data)
    {
        if (!_config.AuditLlmInteractions)
        {
            return;
        }

        try
        {
            _auditLog.WriteAsync(new AuditEvent
            {
                Timestamp = DateTimeOffset.UtcNow,
                EventType = eventType,
                Message = message,
                Data = data
            }, CancellationToken.None).GetAwaiter().GetResult();
        }
        catch
        {
            // Best effort: never fail intent interpretation because audit logging failed.
        }
    }
}
