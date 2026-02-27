using System.Text.Json;
using DesktopAgent.Cli;
using DesktopAgent.Core.Abstractions;
using DesktopAgent.Core.Config;
using DesktopAgent.Core.Models;
using DesktopAgent.Core.Services;
using DesktopAgent.Proto;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;

AppContext.SetSwitch("System.Net.Http.SocketsHttpHandler.Http2UnencryptedSupport", true);

var configuration = new ConfigurationBuilder()
    .SetBasePath(AppContext.BaseDirectory)
    .AddJsonFile("appsettings.json", optional: true)
    .AddEnvironmentVariables(prefix: "DESKTOP_AGENT_")
    .Build();

var agentConfig = new AgentConfig();
configuration.Bind(agentConfig);
AgentConfigSanitizer.Normalize(agentConfig);
try
{
    ApplyCliOverrides(agentConfig, args);
}
catch (ArgumentException ex)
{
    Console.WriteLine(ex.Message);
    PrintHelp();
    Environment.Exit(2);
}

var services = new ServiceCollection();
services.AddSingleton(agentConfig);
services.AddLogging(builder => builder.AddSimpleConsole(options =>
{
    options.SingleLine = true;
    options.TimestampFormat = "HH:mm:ss ";
}));
services.AddSingleton<IDesktopAdapterClient>(sp =>
{
    var logger = sp.GetRequiredService<ILogger<DesktopGrpcClient>>();
    return new DesktopGrpcClient(agentConfig.AdapterEndpoint, logger);
});
services.AddSingleton<IAppCatalog>(sp => new LocalAppCatalog(sp.GetRequiredService<AgentConfig>()));
services.AddSingleton<IAppResolver, AppResolver>();
services.AddSingleton<ILlmIntentRewriter, LocalLlmIntentRewriter>();
services.AddSingleton<RuleBasedIntentInterpreter>();
services.AddSingleton<IIntentInterpreter, FallbackIntentInterpreter>();
services.AddSingleton<IPlanner, SimplePlanner>();
services.AddSingleton<IPolicyEngine, PolicyEngine>();
services.AddSingleton<IRateLimiter>(_ => new SlidingWindowRateLimiter(() => agentConfig.MaxActionsPerSecond));
services.AddSingleton<IAuditLog, JsonlAuditLog>();
services.AddSingleton<IKillSwitch, KillSwitch>();
services.AddSingleton<IUserConfirmation, ConsoleConfirmation>();
services.AddSingleton<IContextProvider, ContextProvider>();
services.AddSingleton<IExecutor, Executor>();
services.AddSingleton<AgentOrchestrator>();
services.AddSingleton<IOcrEngine>(sp =>
{
    var logger = sp.GetRequiredService<ILogger<TesseractOcrEngine>>();
    if (!agentConfig.OcrEnabled)
    {
        return new OcrEngineStub();
    }

    return new TesseractOcrEngine(agentConfig.Ocr.TesseractPath, logger);
});

var provider = services.BuildServiceProvider();
var logger = provider.GetRequiredService<ILoggerFactory>().CreateLogger("Cli");
var client = provider.GetRequiredService<IDesktopAdapterClient>();
var executor = provider.GetRequiredService<IExecutor>();
var planner = provider.GetRequiredService<IPlanner>();
var orchestrator = provider.GetRequiredService<AgentOrchestrator>();
var killSwitch = provider.GetRequiredService<IKillSwitch>();

if (args.Length == 0 || args[0] is "help" or "--help" or "-h")
{
    PrintHelp();
    return;
}

var command = args[0].ToLowerInvariant();
var cancellationToken = CancellationToken.None;

switch (command)
{
    case "profile":
        var requestedProfile = args.Length > 1 ? args[1] : "balanced";
        agentConfig.ActiveProfile = AgentProfileService.NormalizeProfile(requestedProfile);
        agentConfig.ProfileModeEnabled = true;
        AgentProfileService.ApplyActiveProfile(agentConfig);
        Console.WriteLine($"Profile: {agentConfig.ActiveProfile}");
        Console.WriteLine($"RequireConfirmation: {agentConfig.RequireConfirmation}");
        Console.WriteLine($"MaxActionsPerSecond: {agentConfig.MaxActionsPerSecond}");
        Console.WriteLine($"QuizSafeModeEnabled: {agentConfig.QuizSafeModeEnabled}");
        Console.WriteLine($"PostCheckStrict: {agentConfig.PostCheckStrict}");
        Console.WriteLine($"ContextBindingEnabled: {agentConfig.ContextBindingEnabled}");
        Console.WriteLine($"ContextBindingRequireWindow: {agentConfig.ContextBindingRequireWindow}");
        break;
    case "arm":
        killSwitch.Reset();
        var requirePresence = args.Contains("--require-user-presence", StringComparer.OrdinalIgnoreCase);
        var armStatus = await client.ArmAsync(requirePresence, cancellationToken);
        Console.WriteLine($"Armed: {armStatus.Armed}, RequireUserPresence: {armStatus.RequireUserPresence}, KillSwitch: OFF");
        break;
    case "disarm":
        killSwitch.Trip("CLI requested disarm");
        var disarmStatus = await client.DisarmAsync(cancellationToken);
        Console.WriteLine($"Armed: {disarmStatus.Armed}. Running actions stopped.");
        break;
    case "status":
        var status = await client.GetStatusAsync(cancellationToken);
        Console.WriteLine($"Adapter armed: {status.Armed}, RequireUserPresence: {status.RequireUserPresence}");
        Console.WriteLine($"Allowlist: {string.Join(", ", agentConfig.AllowedApps)}");
        Console.WriteLine($"OCR enabled: {agentConfig.OcrEnabled} ({agentConfig.Ocr.Engine})");
        Console.WriteLine($"Quiz safe mode: {agentConfig.QuizSafeModeEnabled}");
        break;
    case "active-window":
        var active = await client.GetActiveWindowAsync(cancellationToken);
        Console.WriteLine(active == null ? "No active window" : $"{active.Title} [{active.AppId}] ({active.Id})");
        break;
    case "list-windows":
        var windows = await client.ListWindowsAsync(cancellationToken);
        foreach (var window in windows)
        {
            Console.WriteLine($"{window.Id}: {window.Title} [{window.AppId}] ({window.Bounds.Width}x{window.Bounds.Height})");
        }
        break;
    case "apps":
    case "list-apps":
        var allowedFlags = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "--allowed-only", "--allowed", "--allowlist" };
        var allowedOnly = args.Skip(1).Any(arg => allowedFlags.Contains(arg));
        var queryArgs = args.Skip(1).Where(arg => !allowedFlags.Contains(arg)).ToArray();
        var query = string.Join(" ", queryArgs);
        if (!allowedOnly && (query.Equals("allowed", StringComparison.OrdinalIgnoreCase) || query.Equals("allowlist", StringComparison.OrdinalIgnoreCase)))
        {
            allowedOnly = true;
            query = string.Empty;
        }
        var resolver = provider.GetRequiredService<IAppResolver>();
        var matches = resolver.Suggest(query, 20);
        if (allowedOnly)
        {
            matches = matches.Where(match => match.IsAllowed).ToList();
        }
        if (matches.Count == 0)
        {
            Console.WriteLine("No apps found.");
            break;
        }

        foreach (var match in matches)
        {
            var allowedTag = match.IsAllowed ? " [allowed]" : string.Empty;
            Console.WriteLine($"{match.Entry.Name}{allowedTag} score={match.Score:0.00} ({match.Entry.Path})");
        }
        break;
    case "find":
        var findText = JoinArgs(args, 1);
        await ExecutePlanAsync(executor, BuildFindPlan(findText), dryRun: false, cancellationToken);
        break;
    case "click":
        var selector = ParseSelector(args.Skip(1).ToArray());
        await ExecutePlanAsync(executor, BuildClickPlan(selector), dryRun: false, cancellationToken);
        break;
    case "type":
        var typeText = JoinArgs(args, 1);
        await ExecutePlanAsync(executor, BuildTypePlan(typeText), dryRun: false, cancellationToken);
        break;
    case "open":
        var app = JoinArgs(args, 1);
        var openResolver = provider.GetRequiredService<IAppResolver>();
        if (openResolver.TryResolveApp(app, out var resolvedApp))
        {
            app = resolvedApp;
        }
        await ExecutePlanAsync(executor, BuildOpenPlan(app), dryRun: false, cancellationToken);
        break;
    case "tutor":
        await ExecuteTutorAsync(executor, cancellationToken);
        break;
    case "run-script":
        var path = args.Length > 1 ? args[1] : string.Empty;
        await RunScriptAsync(path, executor, cancellationToken);
        break;
    case "dry-run":
        var intent = args.Length > 1 && args[1].Equals("intent", StringComparison.OrdinalIgnoreCase)
            ? JoinArgs(args, 2)
            : JoinArgs(args, 1);
        var plan = planner.PlanFromIntent(intent);
        PrintPlan(plan);
        await executor.ExecutePlanAsync(plan, dryRun: true, cancellationToken);
        break;
    case "intent":
        var directIntent = JoinArgs(args, 1);
        var execResult = await orchestrator.ExecuteIntentAsync(directIntent, dryRun: false, cancellationToken);
        PrintExecution(execResult);
        break;
    case "kill":
        killSwitch.Trip("CLI requested kill");
        Console.WriteLine("Kill switch tripped");
        break;
    case "reset-kill":
    case "unkill":
        killSwitch.Reset();
        Console.WriteLine("Kill switch reset");
        break;
    default:
        Console.WriteLine($"Unknown command: {command}");
        PrintHelp();
        break;
}

static async Task ExecutePlanAsync(IExecutor executor, ActionPlan plan, bool dryRun, CancellationToken cancellationToken)
{
    var result = await executor.ExecutePlanAsync(plan, dryRun, cancellationToken);
    PrintExecution(result);
}

static async Task RunScriptAsync(string path, IExecutor executor, CancellationToken cancellationToken)
{
    if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
    {
        Console.WriteLine("Plan file not found");
        return;
    }

    var json = await File.ReadAllTextAsync(path, cancellationToken);
    var plan = JsonSerializer.Deserialize<ActionPlan>(json, new JsonSerializerOptions
    {
        PropertyNameCaseInsensitive = true
    });

    if (plan == null)
    {
        Console.WriteLine("Invalid plan file");
        return;
    }

    var result = await executor.ExecutePlanAsync(plan, dryRun: false, cancellationToken);
    PrintExecution(result);
}

static ActionPlan BuildFindPlan(string text)
{
    return new ActionPlan
    {
        Intent = $"find {text}",
        Steps = new List<PlanStep>
        {
            new() { Type = ActionType.Find, Selector = new Selector { NameContains = text } }
        }
    };
}

static ActionPlan BuildClickPlan(Selector selector)
{
    return new ActionPlan
    {
        Intent = "click",
        Steps = new List<PlanStep>
        {
            new() { Type = ActionType.Click, Selector = selector }
        }
    };
}

static ActionPlan BuildTypePlan(string text)
{
    return new ActionPlan
    {
        Intent = $"type {text}",
        Steps = new List<PlanStep>
        {
            new() { Type = ActionType.TypeText, Text = text }
        }
    };
}

static ActionPlan BuildOpenPlan(string app)
{
    return new ActionPlan
    {
        Intent = $"open {app}",
        Steps = new List<PlanStep>
        {
            new() { Type = ActionType.OpenApp, AppIdOrPath = app }
        }
    };
}

static ActionPlan BuildReadPlan()
{
    return new ActionPlan
    {
        Intent = "tutor",
        Steps = new List<PlanStep>
        {
            new() { Type = ActionType.ReadText }
        }
    };
}

static async Task ExecuteTutorAsync(IExecutor executor, CancellationToken cancellationToken)
{
    var plan = BuildReadPlan();
    var result = await executor.ExecutePlanAsync(plan, dryRun: false, cancellationToken);
    PrintExecution(result);

    foreach (var step in result.Steps)
    {
        if (step.Data is IEnumerable<DesktopAgent.Core.Models.OcrTextRegion> regions)
        {
            var text = string.Join(" ", regions.Select(r => r.Text));
            Console.WriteLine($"Tutor mode text: {text}");
        }
    }
}

static Selector ParseSelector(string[] args)
{
    var selector = new Selector();
    for (var i = 0; i < args.Length; i++)
    {
        switch (args[i])
        {
            case "--name" when i + 1 < args.Length:
                selector.NameContains = args[++i];
                break;
            case "--role" when i + 1 < args.Length:
                selector.Role = args[++i];
                break;
            case "--automationId" when i + 1 < args.Length:
                selector.AutomationId = args[++i];
                break;
            case "--class" when i + 1 < args.Length:
                selector.ClassName = args[++i];
                break;
            case "--ancestor" when i + 1 < args.Length:
                selector.AncestorNameContains = args[++i];
                break;
            case "--index" when i + 1 < args.Length && int.TryParse(args[i + 1], out var idx):
                selector.Index = idx;
                i++;
                break;
            case "--window" when i + 1 < args.Length:
                selector.WindowId = args[++i];
                break;
        }
    }

    return selector;
}

static string JoinArgs(string[] args, int start)
{
    return args.Length <= start ? string.Empty : string.Join(' ', args.Skip(start));
}

static void PrintPlan(ActionPlan plan)
{
    Console.WriteLine($"Intent: {plan.Intent}");
    for (var i = 0; i < plan.Steps.Count; i++)
    {
        var step = plan.Steps[i];
        var details = $"{i + 1}. {step.Type} {step.Text} {step.Target} {step.AppIdOrPath}".TrimEnd();
        if (!string.IsNullOrWhiteSpace(step.Note))
        {
            details = $"{details} ({step.Note})";
        }
        Console.WriteLine(details);
    }
}

static void PrintExecution(ExecutionResult result)
{
    Console.WriteLine($"Success: {result.Success} - {result.Message}");
    foreach (var step in result.Steps)
    {
        Console.WriteLine($"[{step.Index}] {step.Type} => {step.Success} ({step.Message})");
    }
}

static void PrintHelp()
{
    Console.WriteLine("DesktopAgent CLI commands:");
    Console.WriteLine("  profile <safe|balanced|power>");
    Console.WriteLine("  arm [--require-user-presence]");
    Console.WriteLine("  disarm");
    Console.WriteLine("  status");
    Console.WriteLine("  active-window");
    Console.WriteLine("  list-windows");
    Console.WriteLine("  list-apps [query] [--allowed-only]");
    Console.WriteLine("  find <text>");
    Console.WriteLine("  click --name <text> [--role <role>] [--automationId <id>] [--class <class>] [--ancestor <text>] [--index <n>] [--window <id>]");
    Console.WriteLine("  type <text>");
    Console.WriteLine("  open <app>");
    Console.WriteLine("  tutor");
    Console.WriteLine("  run-script <plan.json>");
    Console.WriteLine("  dry-run intent <text>");
    Console.WriteLine("  intent <text>");
    Console.WriteLine("Plugin intents:");
    Console.WriteLine("  file write <path> <text>");
    Console.WriteLine("  file append <path> <text>");
    Console.WriteLine("  file read <path>");
    Console.WriteLine("  file list [path]");
    Console.WriteLine("  open url <url>");
    Console.WriteLine("  double click <target>");
    Console.WriteLine("  right click <target>");
    Console.WriteLine("  drag <source> to <target>");
    Console.WriteLine("  wait until <text> [for <seconds>]");
    Console.WriteLine("  minimize/maximize/restore/switch window, focus <app>, scroll up/down [n]");
    Console.WriteLine("  notify <text>");
    Console.WriteLine("  clipboard history");
    Console.WriteLine("  volume up/down/mute [n], brightness up/down [n], lock screen");
    Console.WriteLine("  kill");
    Console.WriteLine("  reset-kill");
    Console.WriteLine("Global flags:");
    Console.WriteLine("  --strict-postcheck | --lenient-postcheck");
    Console.WriteLine("  --postcheck-menuitem <window-change|none>");
    Console.WriteLine("  --postcheck-checkbox <present|none>");
    Console.WriteLine("  --postcheck-button <disappear-or-window|none>");
    Console.WriteLine("Example:");
    Console.WriteLine("  desktop-agent --lenient-postcheck click --name \"OK\"");
}

static void ApplyCliOverrides(AgentConfig config, string[] args)
{
    if (args.Any(arg => arg.Equals("--lenient-postcheck", StringComparison.OrdinalIgnoreCase)))
    {
        config.PostCheckStrict = false;
    }

    if (args.Any(arg => arg.Equals("--strict-postcheck", StringComparison.OrdinalIgnoreCase)))
    {
        config.PostCheckStrict = true;
    }

    var menuRule = GetArgValue(args, "--postcheck-menuitem");
    if (!string.IsNullOrWhiteSpace(menuRule))
    {
        if (IsValidRule(menuRule))
        {
            config.PostCheckRules.MenuItem = menuRule;
        }
        else
        {
            throw new ArgumentException($"Invalid value for --postcheck-menuitem: {menuRule}. Allowed: window-change, none.");
        }
    }

    var checkboxRule = GetArgValue(args, "--postcheck-checkbox");
    if (!string.IsNullOrWhiteSpace(checkboxRule))
    {
        if (IsValidRule(checkboxRule))
        {
            config.PostCheckRules.Checkbox = checkboxRule;
        }
        else
        {
            throw new ArgumentException($"Invalid value for --postcheck-checkbox: {checkboxRule}. Allowed: present, none.");
        }
    }

    var buttonRule = GetArgValue(args, "--postcheck-button");
    if (!string.IsNullOrWhiteSpace(buttonRule))
    {
        if (IsValidRule(buttonRule))
        {
            config.PostCheckRules.Button = buttonRule;
        }
        else
        {
            throw new ArgumentException($"Invalid value for --postcheck-button: {buttonRule}. Allowed: disappear-or-window, none.");
        }
    }
}

static string? GetArgValue(string[] args, string name)
{
    for (var i = 0; i < args.Length - 1; i++)
    {
        if (args[i].Equals(name, StringComparison.OrdinalIgnoreCase))
        {
            return args[i + 1];
        }
    }
    return null;
}

static bool IsValidRule(string value)
{
    var normalized = value.Trim().ToLowerInvariant();
    return normalized is "window-change" or "present" or "disappear-or-window" or "none";
}
