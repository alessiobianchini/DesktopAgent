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
        var ruleBasedPlan = _ruleBased.Interpret(intent);
        if (!_config.LlmFallbackEnabled)
        {
            return ruleBasedPlan;
        }

        // Keep deterministic, safe parser behavior for explicit snapshot commands,
        // including chained actions like "take screenshot and open notepad".
        if (ShouldPreserveDeterministicRuleBasedPlan(ruleBasedPlan))
        {
            return ruleBasedPlan;
        }

        var mode = (_config.LlmInterpretationMode ?? "primary").Trim().ToLowerInvariant();
        if (mode == "fallback" && !ShouldFallback(ruleBasedPlan))
        {
            return ruleBasedPlan;
        }

        var rewritten = _rewriter.Rewrite(intent);
        if (!string.IsNullOrWhiteSpace(rewritten))
        {
            var rewrittenPlan = PruneNonActionableNoise(_ruleBased.Interpret(rewritten));
            if (!ShouldFallback(rewrittenPlan))
            {
                PreserveMouseJiggleDuration(intent, rewrittenPlan);
                PreserveScreenRecordingDuration(intent, rewrittenPlan);
                rewrittenPlan.Intent = intent;
                if (rewrittenPlan.Steps.Count > 0 && string.IsNullOrWhiteSpace(rewrittenPlan.Steps[0].Note))
                {
                    rewrittenPlan.Steps[0].Note = $"Rewritten intent: {rewritten}";
                }

                WriteAudit("llm_rewrite_applied", "LLM rewrite applied", new
                {
                    input = ToAuditText(intent),
                    rewritten = ToAuditText(rewritten),
                    translatedCommand = rewritten,
                    stepCount = rewrittenPlan.Steps.Count
                });
                return rewrittenPlan;
            }

            WriteAudit("llm_fallback_rule_based", "LLM rewrite unusable; using rule-based plan", new
            {
                input = ToAuditText(intent),
                rewritten = ToAuditText(rewritten)
            });
            return ruleBasedPlan;
        }

        WriteAudit("llm_fallback_rule_based", "LLM rewrite unavailable; using rule-based plan", new
        {
            input = ToAuditText(intent)
        });
        return ruleBasedPlan;
    }

    private static bool ShouldFallback(ActionPlan plan)
    {
        if (plan.Steps.Count == 0)
        {
            return true;
        }

        var actionable = plan.Steps.Count(step => step.Type != ActionType.ReadText);
        if (actionable == 0)
        {
            return true;
        }

        return false;
    }

    private static bool ShouldPreserveDeterministicRuleBasedPlan(ActionPlan plan)
    {
        return plan.Steps.Any(step => step.Type == ActionType.CaptureScreen);
    }

    private static ActionPlan PruneNonActionableNoise(ActionPlan plan)
    {
        if (plan.Steps.Count == 0)
        {
            return plan;
        }

        var actionable = plan.Steps.Count(step => step.Type != ActionType.ReadText);
        if (actionable == 0)
        {
            return plan;
        }

        plan.Steps = plan.Steps
            .Where(step => step.Type != ActionType.ReadText
                           || string.IsNullOrWhiteSpace(step.Note)
                           || !step.Note.StartsWith("Unrecognized", StringComparison.OrdinalIgnoreCase))
            .ToList();

        return plan;
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

    private void PreserveScreenRecordingDuration(string originalIntent, ActionPlan rewrittenPlan)
    {
        var rewrittenRecord = rewrittenPlan.Steps.FirstOrDefault(step => step.Type == ActionType.RecordScreen);
        if (rewrittenRecord == null)
        {
            return;
        }

        var rewrittenDuration = rewrittenRecord.WaitFor ?? TimeSpan.FromSeconds(30);
        if (rewrittenDuration != TimeSpan.FromSeconds(30))
        {
            return;
        }

        var originalPlan = _ruleBased.Interpret(originalIntent);
        var originalRecord = originalPlan.Steps.FirstOrDefault(step => step.Type == ActionType.RecordScreen);
        if (originalRecord?.WaitFor is not TimeSpan originalDuration)
        {
            return;
        }

        if (originalDuration <= TimeSpan.Zero || originalDuration == TimeSpan.FromSeconds(30))
        {
            return;
        }

        rewrittenRecord.WaitFor = originalDuration;
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
