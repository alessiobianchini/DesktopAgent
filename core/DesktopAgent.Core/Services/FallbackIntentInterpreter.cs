using DesktopAgent.Core.Abstractions;
using DesktopAgent.Core.Config;
using DesktopAgent.Core.Models;
using System.Text.RegularExpressions;

namespace DesktopAgent.Core.Services;

public sealed class FallbackIntentInterpreter : IIntentInterpreter
{
    private const string LowConfidenceMarker = "llm-low-confidence";
    private const string ClarificationMarker = "llm-needs-clarification";
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
        if (rewritten != null && !string.IsNullOrWhiteSpace(rewritten.Command))
        {
            if (!IsRewriteCommandSchemaValid(rewritten.Command))
            {
                WriteAudit("llm_fallback_rule_based", "LLM rewrite rejected by schema validator", new
                {
                    input = ToAuditText(intent),
                    rewritten = ToAuditText(rewritten.Command),
                    confidence = rewritten.Confidence
                });
                return ruleBasedPlan;
            }

            var rewrittenPlan = PruneNonActionableNoise(_ruleBased.Interpret(rewritten.Command));
            if (!ShouldFallback(rewrittenPlan))
            {
                if (ShouldPreserveDeterministicRuleBasedPlan(ruleBasedPlan)
                    && !HasRequiredMediaActions(rewrittenPlan))
                {
                    WriteAudit("llm_fallback_rule_based", "LLM rewrite dropped required media action; keeping rule-based plan", new
                    {
                        input = ToAuditText(intent),
                        rewritten = ToAuditText(rewritten.Command),
                        confidence = rewritten.Confidence
                    });
                    return ruleBasedPlan;
                }

                PreserveMouseJiggleDuration(intent, rewrittenPlan);
                PreserveScreenRecordingDuration(intent, rewrittenPlan);
                rewrittenPlan.Intent = intent;
                if (rewrittenPlan.Steps.Count > 0)
                {
                    var noteParts = new List<string> { $"Rewritten intent: {rewritten.Command}" };
                    if (rewritten.NeedsClarification)
                    {
                        noteParts.Add($"{ClarificationMarker}");
                        if (!string.IsNullOrWhiteSpace(rewritten.ClarificationQuestion))
                        {
                            noteParts.Add($"Clarification: {rewritten.ClarificationQuestion}");
                        }
                    }

                    var minConfidence = Math.Clamp(_config.LlmFallback.MinConfidence, 0.0, 1.0);
                    if (rewritten.Confidence < minConfidence)
                    {
                        noteParts.Add($"{LowConfidenceMarker}:{rewritten.Confidence:F2}");
                    }

                    rewrittenPlan.Steps[0].Note = string.Join(" | ", noteParts);
                }

                WriteAudit("llm_rewrite_applied", "LLM rewrite applied", new
                {
                    input = ToAuditText(intent),
                    rewritten = ToAuditText(rewritten.Command),
                    translatedCommand = rewritten.Command,
                    confidence = rewritten.Confidence,
                    needsClarification = rewritten.NeedsClarification,
                    clarificationQuestion = rewritten.ClarificationQuestion,
                    stepCount = rewrittenPlan.Steps.Count,
                    lowConfidence = rewritten.Confidence < Math.Clamp(_config.LlmFallback.MinConfidence, 0.0, 1.0)
                });
                return rewrittenPlan;
            }

            WriteAudit("llm_fallback_rule_based", "LLM rewrite unusable; using rule-based plan", new
            {
                input = ToAuditText(intent),
                rewritten = ToAuditText(rewritten.Command),
                confidence = rewritten.Confidence
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
        return plan.Steps.Any(step => step.Type is ActionType.CaptureScreen
            or ActionType.RecordScreen
            or ActionType.StartScreenRecording
            or ActionType.StopScreenRecording);
    }

    private static bool HasRequiredMediaActions(ActionPlan plan)
    {
        return plan.Steps.Any(step => step.Type is ActionType.CaptureScreen
            or ActionType.RecordScreen
            or ActionType.StartScreenRecording
            or ActionType.StopScreenRecording);
    }

    private static bool IsRewriteCommandSchemaValid(string command)
    {
        if (string.IsNullOrWhiteSpace(command))
        {
            return false;
        }

        var normalized = command.Trim().ToLowerInvariant();
        if (normalized.Contains('\n') || normalized.Contains('\r'))
        {
            return false;
        }

        var commandStart = Regex.IsMatch(
            normalized,
            "^(open|find|click|double click|right click|drag|type|press|save|save as|new tab|close tab|close window|minimize window|maximize window|restore window|switch window|focus|scroll|page up|page down|home|end|wait until|copy|paste|undo|redo|select all|open url|search|browser back|browser forward|refresh|find in page|file|notify|clipboard history|volume|brightness|lock screen|create new file|move mouse|jiggle mouse|record screen|start recording|stop recording|take screenshot|snapshot|apri|avvia|lancia|esegui|cerca|clicca|scrivi|digita|premi|registra|cattura)",
            RegexOptions.IgnoreCase);

        return commandStart;
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
