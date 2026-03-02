# DesktopAgent Architecture

## Overview
DesktopAgent uses a cross-platform core (.NET 9) and OS-specific adapters that expose a gRPC service. The core prefers UI-tree / accessibility data when available, and falls back to vision (screenshot + OCR + coordinate actions) when it is not.

```
+------------------------+        gRPC        +-------------------------------+
|  DesktopAgent.Core     | <----------------> |  OS Adapter (Win/Mac/Linux)   |
|  - Planner             |                    |  - UI Tree (UIA/AX/AT-SPI2)   |
|  - Policy Engine       |                    |  - Input Injection            |
|  - Executor            |                    |  - Screenshot Capture         |
|  - OCR (optional)      |                    +-------------------------------+
|  - Audit Log           |
+------------------------+
```

## Option B Execution Flow
1. `ContextProvider` requests UI tree snapshot (`GetUiTree`).
2. If UI tree is unavailable, `ContextProvider` requests screenshot (`CaptureScreen`) and optional OCR.
3. `Planner` produces abstract steps (Find, Click, TypeText, OpenApp, KeyCombo, WaitFor, ReadText).
4. `PolicyEngine` enforces allowlist, confirmations, rate limiting, and quiz-safe mode.
5. `Executor` executes steps via adapter APIs and verifies post-conditions when possible.
6. `AuditLog` writes JSONL entries for every step and decision.

## Safety Layers
- Disarmed by default (adapter side).
- Allowlist of allowed app IDs/titles.
- Confirmation required for dangerous actions (submit/send/pay/delete/enter).
- Rate limit (default 3 actions/second).
- Quiz/exam detection prevents clicks and submissions unless config overrides.
- Kill switch stops execution immediately.

## Selector Model
`Selector` is a simple DSL used for UI-tree matching:
- role
- nameContains
- automationId
- className
- ancestorNameContains
- index
- boundsHint

Adapters implement best-effort matching on these fields.

## Data Contracts
All inter-process communication uses `proto/desktop_adapter.proto`. Handles are abstracted as string IDs in `WindowRef` and `ElementRef` to avoid OS-handle leakage into core logic.
