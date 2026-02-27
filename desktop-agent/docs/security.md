# Security and Safety Guardrails

DesktopAgent is designed to be safety-first and to discourage misuse. It is not intended for cheating or auto-submitting quiz/exam answers.

## Guardrails Implemented
- **Disarmed by default**: Adapters refuse all actions until explicitly armed.
- **Allowlist**: Actions are blocked unless the active window matches the configured allowlist.
- **Dangerous action detection**: Keywords like "submit", "send", "pay", "purchase", "delete", and "confirm" trigger confirmation or blocking.
- **Quiz-safe mode**: If the active window title contains `quiz`, `exam`, `test`, or `assessment`, interactive actions are blocked when `QuizSafeModeEnabled=true`.
- **Rate limiting**: Default limit is 3 actions/second to reduce runaway automation.
- **Audit logging**: All actions and decisions are logged as JSONL.
- **Kill switch**: Stops execution immediately.
- **Post-condition checks**: After actions, the executor verifies target presence or window changes. In strict mode (default), indeterminate checks fail; in lenient mode they pass with a warning.
  - Config: set `PostCheckStrict=false` in `appsettings.json` or use `--lenient-postcheck`.
  - Role rules: `PostCheckRules` controls expectations per role (e.g., buttons should disappear, checkboxes should remain, menu items should change windows).
  - CLI flags: `--postcheck-menuitem`, `--postcheck-checkbox`, `--postcheck-button`.

## Non-Cheating Policy
The policy engine blocks auto-submission actions by default. In quiz/exam contexts, the system falls back to "explain only" mode (read-only actions) unless the user explicitly disables `QuizSafeModeEnabled` in config.

## Risk Notes
- Accessibility APIs can access sensitive UI content. Use allowlist and explicit confirmations.
- OCR-based actions are inherently less precise. Confirm before executing OCR-based clicks.
- Clipboard modifications can leak data; confirmation is required by default.

## Recommendations
- Keep the allowlist narrow.
- Require confirmations for any irreversible actions.
- Use dry-run mode to preview plans before executing.
- Use `tutor` mode for read-only assistance when in quiz/exam contexts.
- Use `--lenient-postcheck` for flaky UI trees where verification is unreliable.
