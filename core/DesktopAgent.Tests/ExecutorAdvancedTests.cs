using DesktopAgent.Core.Abstractions;
using DesktopAgent.Core.Config;
using DesktopAgent.Core.Models;
using DesktopAgent.Core.Services;
using DesktopAgent.Proto;
using Microsoft.Extensions.Logging.Abstractions;
using Xunit;

namespace DesktopAgent.Tests;

public sealed class ExecutorAdvancedTests
{
    [Fact]
    public async Task ClickPostCheck_WaitsUntilTargetDisappears()
    {
        var client = new StubDesktopClient();
        client.EnqueueActiveWindow(new WindowRef { Id = "w-app", AppId = "app", Title = "App" });
        var findCall = 0;
        client.OnFind = selector =>
        {
            if (!string.Equals(selector.NameContains, "Run", StringComparison.OrdinalIgnoreCase))
            {
                return Array.Empty<ElementRef>();
            }

            findCall++;
            if (findCall < 3)
            {
                return new[]
                {
                    new ElementRef
                    {
                        Id = "run",
                        Name = "Run",
                        Role = "Button",
                        Bounds = new Rect { X = 10, Y = 10, Width = 80, Height = 24 }
                    }
                };
            }

            return Array.Empty<ElementRef>();
        };

        var config = new AgentConfig
        {
            OcrEnabled = false,
            ContextBindingEnabled = false,
            PostCheckStrict = true,
            FindRetryCount = 0,
            FindRetryDelayMs = 0,
            PostCheckTimeoutMs = 500,
            PostCheckPollMs = 40
        };
        var executor = BuildExecutor(client, config);

        var plan = new ActionPlan
        {
            Steps =
            {
                new PlanStep { Type = ActionType.Click, Selector = new Selector { NameContains = "Run" } }
            }
        };

        var result = await executor.ExecutePlanAsync(plan, dryRun: false, CancellationToken.None);

        Assert.True(result.Success);
        Assert.Single(result.Steps);
        Assert.True(result.Steps[0].Success);
        Assert.Contains("disappeared", result.Steps[0].Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ClickPostCheck_FailsWhenCheckboxMissingAfterTimeout()
    {
        var client = new StubDesktopClient();
        client.EnqueueActiveWindow(new WindowRef { Id = "w-app", AppId = "app", Title = "App" });
        var findCall = 0;
        client.OnFind = selector =>
        {
            if (!string.Equals(selector.NameContains, "Enable setting", StringComparison.OrdinalIgnoreCase))
            {
                return Array.Empty<ElementRef>();
            }

            findCall++;
            if (findCall == 1)
            {
                return new[]
                {
                    new ElementRef
                    {
                        Id = "check",
                        Name = "Enable setting",
                        Role = "Checkbox",
                        Bounds = new Rect { X = 20, Y = 20, Width = 100, Height = 24 }
                    }
                };
            }

            return Array.Empty<ElementRef>();
        };

        var config = new AgentConfig
        {
            OcrEnabled = false,
            ContextBindingEnabled = false,
            PostCheckStrict = true,
            FindRetryCount = 0,
            FindRetryDelayMs = 0,
            PostCheckTimeoutMs = 180,
            PostCheckPollMs = 40
        };
        var executor = BuildExecutor(client, config);

        var plan = new ActionPlan
        {
            Steps =
            {
                new PlanStep { Type = ActionType.Click, Selector = new Selector { NameContains = "Enable setting" } }
            }
        };

        var result = await executor.ExecutePlanAsync(plan, dryRun: false, CancellationToken.None);

        Assert.False(result.Success);
        Assert.Single(result.Steps);
        Assert.False(result.Steps[0].Success);
        Assert.Contains("checkbox missing", result.Steps[0].Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task OpenApp_SettleWaitExitsEarlyWhenTargetAppIsForeground()
    {
        var client = new StubDesktopClient();
        client.EnqueueActiveWindow(new WindowRef { Id = "w-shell", AppId = "explorer", Title = "Desktop" }); // pre-step read
        client.EnqueueActiveWindow(new WindowRef { Id = "w-chrome", AppId = "chrome", Title = "Google Chrome" }); // first poll match

        var config = new AgentConfig
        {
            OcrEnabled = false,
            ContextBindingEnabled = false,
            OpenAppSettleDelayMs = 500
        };
        var executor = BuildExecutor(client, config);
        var plan = new ActionPlan
        {
            Steps =
            {
                new PlanStep { Type = ActionType.OpenApp, AppIdOrPath = "chrome" }
            }
        };

        var started = DateTime.UtcNow;
        var result = await executor.ExecutePlanAsync(plan, dryRun: false, CancellationToken.None);
        var elapsedMs = (DateTime.UtcNow - started).TotalMilliseconds;

        Assert.True(result.Success);
        Assert.Single(result.Steps);
        Assert.True(result.Steps[0].Success);
        Assert.True(elapsedMs < 350, $"Expected early exit before full settle timeout, got {elapsedMs:0}ms");
    }

    [Fact]
    public async Task ContextBindingBlocksWhenActiveAppChanges()
    {
        var client = new StubDesktopClient();
        client.EnqueueActiveWindow(new WindowRef { Id = "w-code", AppId = "code", Title = "Code" });
        client.EnqueueActiveWindow(new WindowRef { Id = "w-code", AppId = "code", Title = "Code" });
        client.EnqueueActiveWindow(new WindowRef { Id = "w-notepad", AppId = "notepad", Title = "Notepad" });
        client.OnFind = selector => new List<ElementRef>
        {
            new()
            {
                Id = "e-1",
                Name = "Ready",
                Role = "Button",
                Bounds = new Rect { X = 10, Y = 10, Width = 100, Height = 20 }
            }
        };

        var config = new AgentConfig
        {
            OcrEnabled = false,
            ContextBindingEnabled = true,
            ContextBindingRequireWindow = false,
            FindRetryCount = 0,
            FindRetryDelayMs = 0
        };
        var executor = BuildExecutor(client, config);

        var plan = new ActionPlan
        {
            Steps =
            {
                new PlanStep { Type = ActionType.Find, Selector = new Selector { NameContains = "Ready" } },
                new PlanStep { Type = ActionType.KeyCombo, Keys = new List<string> { "ctrl", "n" } }
            }
        };

        var result = await executor.ExecutePlanAsync(plan, dryRun: false, CancellationToken.None);

        Assert.False(result.Success);
        Assert.Contains("context binding", result.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(2, result.Steps.Count);
        Assert.True(result.Steps[0].Success);
        Assert.False(result.Steps[1].Success);
    }

    [Fact]
    public async Task SelectorSelfHealingFallsBackWithoutAutomationId()
    {
        var client = new StubDesktopClient();
        client.EnqueueActiveWindow(new WindowRef { Id = "w-code", AppId = "code", Title = "Code" });
        client.EnqueueActiveWindow(new WindowRef { Id = "w-code", AppId = "code", Title = "Code" });
        client.OnFind = selector =>
        {
            if (!string.IsNullOrWhiteSpace(selector.AutomationId))
            {
                return Array.Empty<ElementRef>();
            }

            if (selector.NameContains.Contains("Open", StringComparison.OrdinalIgnoreCase))
            {
                return new[]
                {
                    new ElementRef
                    {
                        Id = "e-open",
                        Name = "Open",
                        Role = "Button",
                        ClassName = "Button",
                        Bounds = new Rect { X = 25, Y = 25, Width = 80, Height = 24 }
                    }
                };
            }

            return Array.Empty<ElementRef>();
        };

        var config = new AgentConfig
        {
            OcrEnabled = false,
            ContextBindingEnabled = true,
            ContextBindingRequireWindow = false,
            FindRetryCount = 0,
            FindRetryDelayMs = 0
        };
        var executor = BuildExecutor(client, config);

        var plan = new ActionPlan
        {
            Steps =
            {
                new PlanStep
                {
                    Type = ActionType.Find,
                    Selector = new Selector
                    {
                        NameContains = "Open",
                        AutomationId = "changed-automation-id",
                        ClassName = "Button"
                    }
                }
            }
        };

        var result = await executor.ExecutePlanAsync(plan, dryRun: false, CancellationToken.None);

        Assert.True(result.Success);
        Assert.Single(result.Steps);
        Assert.True(result.Steps[0].Success);
        Assert.True(client.FindCalls >= 2);
    }

    [Fact]
    public async Task ExecutePlan_PreservesExecutedStepTypeInResult()
    {
        var client = new StubDesktopClient();
        client.EnqueueActiveWindow(new WindowRef { Id = "w-app", AppId = "app", Title = "App" });
        var config = new AgentConfig
        {
            OcrEnabled = false,
            ContextBindingEnabled = false
        };
        var executor = BuildExecutor(client, config);

        var plan = new ActionPlan
        {
            Steps =
            {
                new PlanStep { Type = ActionType.OpenApp, AppIdOrPath = "notepad" }
            }
        };

        var result = await executor.ExecutePlanAsync(plan, dryRun: false, CancellationToken.None);

        Assert.True(result.Success);
        Assert.Single(result.Steps);
        Assert.Equal(ActionType.OpenApp, result.Steps[0].Type);
    }

    [Fact]
    public async Task SetValue_UsesSelectorWhenElementIdMissing()
    {
        var client = new StubDesktopClient();
        client.EnqueueActiveWindow(new WindowRef { Id = "w-form", AppId = "browser", Title = "Form" });
        client.EnqueueActiveWindow(new WindowRef { Id = "w-form", AppId = "browser", Title = "Form" });
        client.OnFind = selector =>
        {
            if (!string.Equals(selector.NameContains, "Customer name", StringComparison.OrdinalIgnoreCase))
            {
                return Array.Empty<ElementRef>();
            }

            return new[]
            {
                new ElementRef
                {
                    Id = "field-name",
                    Name = "Customer name",
                    Role = "textbox",
                    Bounds = new Rect { X = 10, Y = 10, Width = 120, Height = 24 }
                }
            };
        };

        var executor = BuildExecutor(client, new AgentConfig
        {
            OcrEnabled = false,
            ContextBindingEnabled = false,
            FindRetryCount = 0,
            FindRetryDelayMs = 0
        });

        var plan = new ActionPlan
        {
            Steps =
            {
                new PlanStep
                {
                    Type = ActionType.SetValue,
                    Selector = new Selector { NameContains = "Customer name" },
                    Text = "Alessio"
                }
            }
        };

        var result = await executor.ExecutePlanAsync(plan, dryRun: false, CancellationToken.None);

        Assert.True(result.Success);
        Assert.Equal("field-name", client.LastSetElementId);
        Assert.Equal("Alessio", client.LastSetValue);
    }

    [Fact]
    public async Task OptionalGroup_SkipsRemainingCandidatesAfterFirstSuccess()
    {
        var client = new StubDesktopClient();
        client.EnqueueActiveWindow(new WindowRef { Id = "w-form", AppId = "browser", Title = "Form" });
        client.EnqueueActiveWindow(new WindowRef { Id = "w-form", AppId = "browser", Title = "Form" });
        client.EnqueueActiveWindow(new WindowRef { Id = "w-form", AppId = "browser", Title = "Form" });
        client.OnFind = selector =>
        {
            if (string.Equals(selector.NameContains, "Email", StringComparison.OrdinalIgnoreCase))
            {
                return new[]
                {
                    new ElementRef
                    {
                        Id = "field-email",
                        Name = "Email",
                        Role = "textbox",
                        Bounds = new Rect { X = 10, Y = 10, Width = 120, Height = 24 }
                    }
                };
            }

            return Array.Empty<ElementRef>();
        };

        var executor = BuildExecutor(client, new AgentConfig
        {
            OcrEnabled = false,
            ContextBindingEnabled = false,
            FindRetryCount = 0,
            FindRetryDelayMs = 0
        });

        var plan = new ActionPlan
        {
            Steps =
            {
                new PlanStep
                {
                    Type = ActionType.SetValue,
                    Selector = new Selector { NameContains = "Email" },
                    Text = "a@b.com",
                    Note = "optional-group:email;optional"
                },
                new PlanStep
                {
                    Type = ActionType.SetValue,
                    Selector = new Selector { NameContains = "E-mail" },
                    Text = "a@b.com",
                    Note = "optional-group:email;optional"
                }
            }
        };

        var result = await executor.ExecutePlanAsync(plan, dryRun: false, CancellationToken.None);

        Assert.True(result.Success);
        Assert.Equal(1, client.SetValueCalls);
        Assert.Equal(2, result.Steps.Count);
        Assert.Contains("Skipped optional candidate", result.Steps[1].Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task SetValue_FallsBackToOcr_WhenUiElementNotFound()
    {
        var client = new StubDesktopClient();
        client.EnqueueActiveWindow(new WindowRef { Id = "w-form", AppId = "browser", Title = "Form" });
        client.EnqueueActiveWindow(new WindowRef { Id = "w-form", AppId = "browser", Title = "Form" });
        client.OnFind = _ => Array.Empty<ElementRef>();

        var context = new StubContextProvider
        {
            OnFindByText = _ => new FindResult
            {
                OcrMatches = new List<OcrTextRegion>
                {
                    new()
                    {
                        Text = "Name",
                        Bounds = new Rect { X = 40, Y = 50, Width = 120, Height = 20 },
                        Confidence = 0.9f
                    }
                }
            }
        };

        var executor = BuildExecutor(client, new AgentConfig
        {
            OcrEnabled = true,
            ContextBindingEnabled = false,
            FindRetryCount = 0,
            FindRetryDelayMs = 0
        }, context);

        var plan = new ActionPlan
        {
            Steps =
            {
                new PlanStep
                {
                    Type = ActionType.SetValue,
                    Selector = new Selector { NameContains = "Name" },
                    Text = "Alessio"
                }
            }
        };

        var result = await executor.ExecutePlanAsync(plan, dryRun: false, CancellationToken.None);

        Assert.True(result.Success);
        Assert.Equal(1, client.ClickPointCalls);
        Assert.Equal(1, client.TypeTextCalls);
        Assert.Equal("Alessio", client.LastTypedText);
        Assert.Contains("OCR", result.Steps[0].Message, StringComparison.OrdinalIgnoreCase);
    }

    private static Executor BuildExecutor(StubDesktopClient client, AgentConfig config, StubContextProvider? contextProvider = null)
    {
        var policy = new PolicyEngine(config);
        return new Executor(
            client,
            contextProvider ?? new StubContextProvider(),
            new StubAppResolver(),
            policy,
            new StubRateLimiter(),
            new StubAuditLog(),
            new StubConfirmation(),
            new KillSwitch(),
            config,
            new StubOcrEngine(),
            NullLogger<Executor>.Instance);
    }

    private sealed class StubDesktopClient : IDesktopAdapterClient
    {
        private readonly Queue<WindowRef?> _windows = new();
        private WindowRef? _lastWindow;
        public Func<Selector, IReadOnlyList<ElementRef>>? OnFind { get; set; }
        public int FindCalls { get; private set; }
        public int ActiveWindowCalls { get; private set; }
        public int SetValueCalls { get; private set; }
        public string? LastSetElementId { get; private set; }
        public string? LastSetValue { get; private set; }
        public int ClickPointCalls { get; private set; }
        public int TypeTextCalls { get; private set; }
        public string? LastTypedText { get; private set; }

        public void EnqueueActiveWindow(WindowRef window)
        {
            _windows.Enqueue(window);
            _lastWindow ??= window;
        }

        public Task<WindowRef?> GetActiveWindowAsync(CancellationToken cancellationToken)
        {
            ActiveWindowCalls++;
            if (_windows.Count > 0)
            {
                _lastWindow = _windows.Dequeue();
            }

            return Task.FromResult(_lastWindow);
        }

        public Task<IReadOnlyList<WindowRef>> ListWindowsAsync(CancellationToken cancellationToken) => Task.FromResult<IReadOnlyList<WindowRef>>(Array.Empty<WindowRef>());
        public Task<UiTree?> GetUiTreeAsync(string windowId, CancellationToken cancellationToken) => Task.FromResult<UiTree?>(null);

        public Task<IReadOnlyList<ElementRef>> FindElementsAsync(Selector selector, CancellationToken cancellationToken)
        {
            FindCalls++;
            var result = OnFind?.Invoke(selector) ?? Array.Empty<ElementRef>();
            return Task.FromResult(result);
        }

        public Task<ActionResult> InvokeElementAsync(string elementId, CancellationToken cancellationToken) => Task.FromResult(new ActionResult { Success = true, Message = "ok" });
        public Task<ActionResult> SetElementValueAsync(string elementId, string value, CancellationToken cancellationToken)
        {
            SetValueCalls++;
            LastSetElementId = elementId;
            LastSetValue = value;
            return Task.FromResult(new ActionResult { Success = true, Message = "ok" });
        }
        public Task<ActionResult> ClickPointAsync(int x, int y, CancellationToken cancellationToken)
        {
            ClickPointCalls++;
            return Task.FromResult(new ActionResult { Success = true, Message = "ok" });
        }
        public Task<ActionResult> TypeTextAsync(string text, CancellationToken cancellationToken)
        {
            TypeTextCalls++;
            LastTypedText = text;
            return Task.FromResult(new ActionResult { Success = true, Message = "ok" });
        }
        public Task<ActionResult> KeyComboAsync(IEnumerable<string> keys, CancellationToken cancellationToken) => Task.FromResult(new ActionResult { Success = true, Message = "ok" });
        public Task<ActionResult> OpenAppAsync(string appIdOrPath, CancellationToken cancellationToken) => Task.FromResult(new ActionResult { Success = true, Message = "ok" });
        public Task<ScreenshotResponse> CaptureScreenAsync(ScreenshotRequest request, CancellationToken cancellationToken) => Task.FromResult(new ScreenshotResponse());
        public Task<ClipboardResponse> GetClipboardAsync(CancellationToken cancellationToken) => Task.FromResult(new ClipboardResponse { Text = string.Empty });
        public Task<ActionResult> SetClipboardAsync(string text, CancellationToken cancellationToken) => Task.FromResult(new ActionResult { Success = true, Message = "ok" });
        public Task<Status> ArmAsync(bool requireUserPresence, CancellationToken cancellationToken) => Task.FromResult(new Status { Armed = true });
        public Task<Status> DisarmAsync(CancellationToken cancellationToken) => Task.FromResult(new Status { Armed = false });
        public Task<Status> GetStatusAsync(CancellationToken cancellationToken) => Task.FromResult(new Status { Armed = true });
    }

    private sealed class StubContextProvider : IContextProvider
    {
        public Func<string, FindResult>? OnFindByText { get; set; }
        public Task<ContextSnapshot> GetSnapshotAsync(CancellationToken cancellationToken) => Task.FromResult(new ContextSnapshot());
        public Task<FindResult> FindByTextAsync(string text, CancellationToken cancellationToken)
            => Task.FromResult(OnFindByText?.Invoke(text) ?? new FindResult());
    }

    private sealed class StubAppResolver : IAppResolver
    {
        public bool TryResolveApp(string input, out string resolved)
        {
            resolved = input;
            return false;
        }

        public IReadOnlyList<AppMatch> Suggest(string input, int maxResults) => Array.Empty<AppMatch>();
    }

    private sealed class StubRateLimiter : IRateLimiter
    {
        public bool TryAcquire() => true;
    }

    private sealed class StubAuditLog : IAuditLog
    {
        public Task WriteAsync(AuditEvent auditEvent, CancellationToken cancellationToken) => Task.CompletedTask;
    }

    private sealed class StubConfirmation : IUserConfirmation
    {
        public Task<bool> ConfirmAsync(string message, CancellationToken cancellationToken) => Task.FromResult(true);
    }

    private sealed class StubOcrEngine : IOcrEngine
    {
        public Task<IReadOnlyList<OcrTextRegion>> ReadTextAsync(byte[] pngBytes, CancellationToken cancellationToken)
            => Task.FromResult<IReadOnlyList<OcrTextRegion>>(Array.Empty<OcrTextRegion>());

        public string Name => "stub";
    }
}
