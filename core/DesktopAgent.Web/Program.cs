using System.Diagnostics;
using System.Net.Http.Json;
using System.Reflection;
using System.Text.Json;
using System.Text.RegularExpressions;
using DesktopAgent.Core.Abstractions;
using DesktopAgent.Core.Config;
using DesktopAgent.Core.Models;
using DesktopAgent.Core.Services;
using DesktopAgent.Proto;
using Microsoft.Extensions.FileProviders;

AppContext.SetSwitch("System.Net.Http.SocketsHttpHandler.Http2UnencryptedSupport", true);

var builder = WebApplication.CreateBuilder(args);

builder.Configuration
    .SetBasePath(AppContext.BaseDirectory)
    .AddJsonFile("appsettings.json", optional: true)
    .AddEnvironmentVariables(prefix: "DESKTOP_AGENT_");

var agentConfig = new AgentConfig();
builder.Configuration.Bind(agentConfig);
AgentConfigSanitizer.Normalize(agentConfig);

builder.Services.AddSingleton(agentConfig);
builder.Services.AddLogging();
builder.Services.AddSingleton<IDesktopAdapterClient>(sp =>
{
    var logger = sp.GetRequiredService<ILogger<DesktopGrpcClient>>();
    return new DesktopGrpcClient(agentConfig.AdapterEndpoint, logger);
});
builder.Services.AddSingleton<IAuditLog, JsonlAuditLog>();
builder.Services.AddSingleton<IContextProvider, ContextProvider>();
builder.Services.AddSingleton<IAppCatalog>(sp => new LocalAppCatalog(sp.GetRequiredService<AgentConfig>()));
builder.Services.AddSingleton<IAppResolver, AppResolver>();
builder.Services.AddSingleton<ILlmIntentRewriter, LocalLlmIntentRewriter>();
builder.Services.AddSingleton<RuleBasedIntentInterpreter>();
builder.Services.AddSingleton<IIntentInterpreter, FallbackIntentInterpreter>();
builder.Services.AddSingleton<IPlanner, SimplePlanner>();
builder.Services.AddSingleton<IPolicyEngine, PolicyEngine>();
builder.Services.AddSingleton<IRateLimiter>(_ => new SlidingWindowRateLimiter(() => agentConfig.MaxActionsPerSecond));
builder.Services.AddSingleton<IKillSwitch, KillSwitch>();
builder.Services.AddSingleton<IOcrEngine>(sp =>
{
    var logger = sp.GetRequiredService<ILogger<TesseractOcrEngine>>();
    if (!agentConfig.OcrEnabled)
    {
        return new OcrEngineStub();
    }

    return new TesseractOcrEngine(agentConfig.Ocr.TesseractPath, logger);
});
builder.Services.AddSingleton<ChatActionStore>();
builder.Services.AddSingleton<TargetMemoryStore>();
builder.Services.AddSingleton<MacroRecorderStore>();
builder.Services.AddSingleton<ContextLockStore>();
builder.Services.AddSingleton<LlmAvailabilityCache>();
builder.Services.AddSingleton<RestartRequirementTracker>();
builder.Services.AddSingleton<ConfigFileStore>(sp =>
{
    var env = sp.GetRequiredService<IHostEnvironment>();
    var path = Path.Combine(env.ContentRootPath, "appsettings.json");
    return new ConfigFileStore(path);
});
builder.Services.AddSingleton<TaskLibraryStore>();
builder.Services.AddSingleton<ScheduleLibraryStore>();
builder.Services.AddSingleton<ScheduledTaskRunner>();
builder.Services.AddHostedService(sp => sp.GetRequiredService<ScheduledTaskRunner>());

var app = builder.Build();
var serverVersion = AppVersionHelper.Resolve();

app.UseDefaultFiles();
app.UseStaticFiles(new StaticFileOptions
{
    FileProvider = new PhysicalFileProvider(Path.Combine(app.Environment.ContentRootPath, "wwwroot"))
});

app.MapGet("/api/status", async (IDesktopAdapterClient client, AgentConfig config, LlmAvailabilityCache llmCache, RestartRequirementTracker restartTracker, TargetMemoryStore memory, MacroRecorderStore recorder, ContextLockStore contextLock, IKillSwitch killSwitch) =>
{
    var status = await client.GetStatusAsync(CancellationToken.None);
    var llmStatus = await llmCache.GetAsync(CancellationToken.None);
    return Results.Ok(new
    {
        version = serverVersion,
        adapter = new { status.Armed, status.RequireUserPresence, status.Message },
        config = new
        {
            config.AdapterEndpoint,
            config.ProfileModeEnabled,
            config.ActiveProfile,
            config.AllowedApps,
            config.PostCheckStrict,
            config.PostCheckRules,
            config.ContextBindingEnabled,
            config.ContextBindingRequireWindow
        },
        restart = new
        {
            ocrRequired = restartTracker.OcrRestartRequired
        },
        memory = memory.GetSnapshot(),
        recording = recorder.GetStatus(),
        contextLock = contextLock.GetSnapshot(),
        killSwitch = new
        {
            tripped = killSwitch.IsTripped,
            reason = killSwitch.Reason
        },
        llm = new
        {
            llmStatus.Enabled,
            llmStatus.Available,
            llmStatus.Provider,
            llmStatus.Message,
            llmStatus.Endpoint
        }
    });
});

app.MapGet("/api/config", (AgentConfig config, RestartRequirementTracker restartTracker) =>
{
    return Results.Ok(new
    {
        allowedApps = config.AllowedApps,
        appAliases = config.AppAliases,
        profileModeEnabled = config.ProfileModeEnabled,
        activeProfile = config.ActiveProfile,
        profiles = new
        {
            safe = config.Profiles.Safe,
            balanced = config.Profiles.Balanced,
            power = config.Profiles.Power
        },
        requireConfirmation = config.RequireConfirmation,
        maxActionsPerSecond = config.MaxActionsPerSecond,
        quizSafeModeEnabled = config.QuizSafeModeEnabled,
        ocrEnabled = config.OcrEnabled,
        ocr = new
        {
            engine = config.Ocr.Engine,
            tesseractPath = config.Ocr.TesseractPath
        },
        ocrRestartRequired = restartTracker.OcrRestartRequired,
        adapterRestartCommand = config.AdapterRestartCommand,
        adapterRestartWorkingDir = config.AdapterRestartWorkingDir,
        findRetryCount = config.FindRetryCount,
        findRetryDelayMs = config.FindRetryDelayMs,
        postCheckTimeoutMs = config.PostCheckTimeoutMs,
        postCheckPollMs = config.PostCheckPollMs,
        clipboardHistoryMaxItems = config.ClipboardHistoryMaxItems,
        filesystemAllowedRoots = config.FilesystemAllowedRoots,
        contextBindingEnabled = config.ContextBindingEnabled,
        contextBindingRequireWindow = config.ContextBindingRequireWindow,
        taskLibraryPath = config.TaskLibraryPath,
        scheduleLibraryPath = config.ScheduleLibraryPath,
        auditLlmInteractions = config.AuditLlmInteractions,
        auditLlmIncludeRawText = config.AuditLlmIncludeRawText,
        llm = new
        {
            enabled = config.LlmFallbackEnabled,
            allowNonLoopbackEndpoint = config.AllowNonLoopbackLlmEndpoint,
            provider = config.LlmFallback.Provider,
            endpoint = config.LlmFallback.Endpoint,
            model = config.LlmFallback.Model,
            timeoutSeconds = config.LlmFallback.TimeoutSeconds,
            maxTokens = config.LlmFallback.MaxTokens
        }
    });
});

app.MapPost("/api/config", async (ConfigUpdateRequest request, AgentConfig config, ConfigFileStore store, LlmAvailabilityCache llmCache, RestartRequirementTracker restartTracker) =>
{
    if (request.Llm == null
        && request.AllowedApps == null
        && request.AppAliases == null
        && request.RequireConfirmation == null
        && request.MaxActionsPerSecond == null
        && request.QuizSafeModeEnabled == null
        && request.OcrEnabled == null
        && request.Ocr == null
        && request.AdapterRestartCommand == null
        && request.AdapterRestartWorkingDir == null
        && request.FindRetryCount == null
        && request.FindRetryDelayMs == null
        && request.PostCheckTimeoutMs == null
        && request.PostCheckPollMs == null
        && request.ProfileModeEnabled == null
        && request.ActiveProfile == null
        && request.ClipboardHistoryMaxItems == null
        && request.FilesystemAllowedRoots == null
        && request.ContextBindingEnabled == null
        && request.ContextBindingRequireWindow == null
        && request.TaskLibraryPath == null
        && request.ScheduleLibraryPath == null
        && request.AuditLlmInteractions == null
        && request.AuditLlmIncludeRawText == null)
    {
        return Results.BadRequest(new { message = "No changes provided." });
    }

    var errors = new List<string>();
    var provider = request.Llm?.Provider?.Trim();
    if (!string.IsNullOrWhiteSpace(provider) && !ConfigValidators.IsAllowedProvider(provider))
    {
        errors.Add("Unsupported provider. Use ollama, openai, or llama.cpp.");
    }

    var endpoint = request.Llm?.Endpoint?.Trim();
    if (!string.IsNullOrWhiteSpace(endpoint))
    {
        if (!Uri.TryCreate(endpoint, UriKind.Absolute, out var uri))
        {
            errors.Add("Endpoint must be a valid absolute URL.");
        }
        else if (!ConfigValidators.IsEndpointAllowed(uri, request.Llm?.AllowNonLoopbackEndpoint ?? config.AllowNonLoopbackLlmEndpoint))
        {
            errors.Add("Endpoint must be local (loopback) unless remote endpoints are enabled.");
        }
    }

    if (request.Llm?.TimeoutSeconds is < 1)
    {
        errors.Add("TimeoutSeconds must be >= 1.");
    }

    if (request.Llm?.MaxTokens is < 1)
    {
        errors.Add("MaxTokens must be >= 1.");
    }

    if (request.MaxActionsPerSecond is < 1 or > 60)
    {
        errors.Add("MaxActionsPerSecond must be between 1 and 60.");
    }

    if (request.Ocr?.Engine != null && string.IsNullOrWhiteSpace(request.Ocr.Engine))
    {
        errors.Add("OCR engine cannot be empty.");
    }

    if (request.AdapterRestartWorkingDir != null)
    {
        var dir = request.AdapterRestartWorkingDir.Trim();
        if (!string.IsNullOrWhiteSpace(dir) && !Directory.Exists(dir))
        {
            errors.Add("Adapter restart working directory does not exist.");
        }
    }

    if (request.FindRetryCount is < 0 or > 20)
    {
        errors.Add("FindRetryCount must be between 0 and 20.");
    }

    if (request.FindRetryDelayMs is < 0 or > 2000)
    {
        errors.Add("FindRetryDelayMs must be between 0 and 2000.");
    }

    if (request.PostCheckTimeoutMs is < 100 or > 5000)
    {
        errors.Add("PostCheckTimeoutMs must be between 100 and 5000.");
    }

    if (request.PostCheckPollMs is < 20 or > 1000)
    {
        errors.Add("PostCheckPollMs must be between 20 and 1000.");
    }

    if (request.ClipboardHistoryMaxItems is < 1 or > 500)
    {
        errors.Add("ClipboardHistoryMaxItems must be between 1 and 500.");
    }

    if (request.ActiveProfile != null)
    {
        var normalizedProfile = AgentProfileService.NormalizeProfile(request.ActiveProfile);
        if (normalizedProfile != request.ActiveProfile.Trim().ToLowerInvariant()
            && !string.Equals(request.ActiveProfile.Trim(), "balanced", StringComparison.OrdinalIgnoreCase))
        {
            errors.Add("ActiveProfile must be safe, balanced, or power.");
        }
    }

    if (errors.Count > 0)
    {
        return Results.BadRequest(new { message = string.Join(" ", errors) });
    }

    var previousOcr = new OcrSnapshot(config.OcrEnabled, config.Ocr.Engine, config.Ocr.TesseractPath);

    if (request.AllowedApps != null)
    {
        config.AllowedApps = request.AllowedApps
            .Select(item => item?.Trim())
            .Where(item => !string.IsNullOrWhiteSpace(item))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList()!;
    }

    if (request.AppAliases != null)
    {
        var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var kvp in request.AppAliases)
        {
            var key = kvp.Key?.Trim();
            var value = kvp.Value?.Trim();
            if (string.IsNullOrWhiteSpace(key) || string.IsNullOrWhiteSpace(value))
            {
                continue;
            }
            dict[key] = value;
        }
        config.AppAliases = dict;
    }

    if (request.RequireConfirmation.HasValue)
    {
        config.RequireConfirmation = request.RequireConfirmation.Value;
    }

    if (request.MaxActionsPerSecond.HasValue)
    {
        config.MaxActionsPerSecond = request.MaxActionsPerSecond.Value;
    }

    if (request.QuizSafeModeEnabled.HasValue)
    {
        config.QuizSafeModeEnabled = request.QuizSafeModeEnabled.Value;
    }

    if (request.OcrEnabled.HasValue)
    {
        config.OcrEnabled = request.OcrEnabled.Value;
    }

    if (request.Ocr?.Engine != null)
    {
        config.Ocr.Engine = request.Ocr.Engine.Trim();
    }

    if (request.Ocr?.TesseractPath != null)
    {
        config.Ocr.TesseractPath = request.Ocr.TesseractPath.Trim();
    }

    if (request.AdapterRestartCommand != null)
    {
        config.AdapterRestartCommand = request.AdapterRestartCommand.Trim();
    }

    if (request.AdapterRestartWorkingDir != null)
    {
        config.AdapterRestartWorkingDir = request.AdapterRestartWorkingDir.Trim();
    }

    if (request.FindRetryCount.HasValue)
    {
        config.FindRetryCount = request.FindRetryCount.Value;
    }

    if (request.FindRetryDelayMs.HasValue)
    {
        config.FindRetryDelayMs = request.FindRetryDelayMs.Value;
    }

    if (request.PostCheckTimeoutMs.HasValue)
    {
        config.PostCheckTimeoutMs = request.PostCheckTimeoutMs.Value;
    }

    if (request.PostCheckPollMs.HasValue)
    {
        config.PostCheckPollMs = request.PostCheckPollMs.Value;
    }

    if (request.ProfileModeEnabled.HasValue)
    {
        config.ProfileModeEnabled = request.ProfileModeEnabled.Value;
    }

    if (request.ActiveProfile != null)
    {
        config.ActiveProfile = AgentProfileService.NormalizeProfile(request.ActiveProfile);
    }

    if (request.ClipboardHistoryMaxItems.HasValue)
    {
        config.ClipboardHistoryMaxItems = request.ClipboardHistoryMaxItems.Value;
    }

    if (request.FilesystemAllowedRoots != null)
    {
        config.FilesystemAllowedRoots = request.FilesystemAllowedRoots
            .Select(root => root?.Trim())
            .Where(root => !string.IsNullOrWhiteSpace(root))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList()!;
    }

    if (request.ContextBindingEnabled.HasValue)
    {
        config.ContextBindingEnabled = request.ContextBindingEnabled.Value;
    }

    if (request.ContextBindingRequireWindow.HasValue)
    {
        config.ContextBindingRequireWindow = request.ContextBindingRequireWindow.Value;
    }

    if (request.TaskLibraryPath != null)
    {
        config.TaskLibraryPath = request.TaskLibraryPath.Trim();
    }

    if (request.ScheduleLibraryPath != null)
    {
        config.ScheduleLibraryPath = request.ScheduleLibraryPath.Trim();
    }

    if (request.AuditLlmInteractions.HasValue)
    {
        config.AuditLlmInteractions = request.AuditLlmInteractions.Value;
    }

    if (request.AuditLlmIncludeRawText.HasValue)
    {
        config.AuditLlmIncludeRawText = request.AuditLlmIncludeRawText.Value;
    }

    if (request.Llm?.Enabled.HasValue == true)
    {
        config.LlmFallbackEnabled = request.Llm.Enabled.Value;
    }

    if (request.Llm?.AllowNonLoopbackEndpoint.HasValue == true)
    {
        config.AllowNonLoopbackLlmEndpoint = request.Llm.AllowNonLoopbackEndpoint.Value;
    }

    if (!string.IsNullOrWhiteSpace(provider))
    {
        config.LlmFallback.Provider = provider;
    }

    if (request.Llm?.Endpoint != null)
    {
        config.LlmFallback.Endpoint = endpoint ?? string.Empty;
    }

    if (request.Llm?.Model != null)
    {
        config.LlmFallback.Model = request.Llm.Model.Trim();
    }

    if (request.Llm?.TimeoutSeconds.HasValue == true)
    {
        config.LlmFallback.TimeoutSeconds = request.Llm.TimeoutSeconds.Value;
    }

    if (request.Llm?.MaxTokens.HasValue == true)
    {
        config.LlmFallback.MaxTokens = request.Llm.MaxTokens.Value;
    }

    if (config.ProfileModeEnabled)
    {
        AgentProfileService.ApplyActiveProfile(config);
    }

    var updatedOcr = new OcrSnapshot(config.OcrEnabled, config.Ocr.Engine, config.Ocr.TesseractPath);
    if (!previousOcr.Equals(updatedOcr))
    {
        restartTracker.OcrRestartRequired = true;
    }

    await store.SaveAsync(config);
    llmCache.Invalidate();

    return Results.Ok(new
    {
        allowedApps = config.AllowedApps,
        appAliases = config.AppAliases,
        profileModeEnabled = config.ProfileModeEnabled,
        activeProfile = config.ActiveProfile,
        profiles = new
        {
            safe = config.Profiles.Safe,
            balanced = config.Profiles.Balanced,
            power = config.Profiles.Power
        },
        requireConfirmation = config.RequireConfirmation,
        maxActionsPerSecond = config.MaxActionsPerSecond,
        quizSafeModeEnabled = config.QuizSafeModeEnabled,
        ocrEnabled = config.OcrEnabled,
        ocr = new
        {
            engine = config.Ocr.Engine,
            tesseractPath = config.Ocr.TesseractPath
        },
        ocrRestartRequired = restartTracker.OcrRestartRequired,
        adapterRestartCommand = config.AdapterRestartCommand,
        adapterRestartWorkingDir = config.AdapterRestartWorkingDir,
        findRetryCount = config.FindRetryCount,
        findRetryDelayMs = config.FindRetryDelayMs,
        postCheckTimeoutMs = config.PostCheckTimeoutMs,
        postCheckPollMs = config.PostCheckPollMs,
        clipboardHistoryMaxItems = config.ClipboardHistoryMaxItems,
        filesystemAllowedRoots = config.FilesystemAllowedRoots,
        contextBindingEnabled = config.ContextBindingEnabled,
        contextBindingRequireWindow = config.ContextBindingRequireWindow,
        taskLibraryPath = config.TaskLibraryPath,
        scheduleLibraryPath = config.ScheduleLibraryPath,
        auditLlmInteractions = config.AuditLlmInteractions,
        auditLlmIncludeRawText = config.AuditLlmIncludeRawText,
        llm = new
        {
            enabled = config.LlmFallbackEnabled,
            allowNonLoopbackEndpoint = config.AllowNonLoopbackLlmEndpoint,
            provider = config.LlmFallback.Provider,
            endpoint = config.LlmFallback.Endpoint,
            model = config.LlmFallback.Model,
            timeoutSeconds = config.LlmFallback.TimeoutSeconds,
            maxTokens = config.LlmFallback.MaxTokens
        }
    });
});

app.MapPost("/api/restart", (HttpContext context, IHostApplicationLifetime lifetime) =>
{
    var remoteIp = context.Connection.RemoteIpAddress;
    if (remoteIp != null && !System.Net.IPAddress.IsLoopback(remoteIp))
    {
        return Results.StatusCode(StatusCodes.Status403Forbidden);
    }

    _ = Task.Run(async () =>
    {
        await Task.Delay(250);
        lifetime.StopApplication();
    });

    return Results.Ok(new { message = "Restarting server..." });
});

app.MapGet("/api/version", () => Results.Ok(new { version = serverVersion }));

app.MapPost("/api/adapter/restart", (HttpContext context, AgentConfig config) =>
{
    var remoteIp = context.Connection.RemoteIpAddress;
    if (remoteIp != null && !System.Net.IPAddress.IsLoopback(remoteIp))
    {
        return Results.StatusCode(StatusCodes.Status403Forbidden);
    }

    if (string.IsNullOrWhiteSpace(config.AdapterRestartCommand))
    {
        return Results.BadRequest(new { message = "Adapter restart command not configured." });
    }

    if (!CommandLineParser.TryParseCommand(config.AdapterRestartCommand, out var fileName, out var args))
    {
        return Results.BadRequest(new { message = "Invalid adapter restart command." });
    }

    try
    {
        var psi = new ProcessStartInfo
        {
            FileName = fileName,
            Arguments = args,
            UseShellExecute = false,
            CreateNoWindow = true
        };
        if (!string.IsNullOrWhiteSpace(config.AdapterRestartWorkingDir))
        {
            psi.WorkingDirectory = config.AdapterRestartWorkingDir;
        }
        Process.Start(psi);
        return Results.Ok(new { message = "Adapter restart command launched." });
    }
    catch (Exception ex)
    {
        return Results.Problem($"Failed to launch adapter restart command: {ex.Message}");
    }
});

app.MapGet("/api/utilities/status", (HttpContext context, AgentConfig config) =>
{
    if (!RequestGuards.IsLoopback(context))
    {
        return Results.StatusCode(StatusCodes.Status403Forbidden);
    }

    var ffmpeg = UtilityInstaller.Probe("ffmpeg", "-version");
    var tesseractCommand = string.IsNullOrWhiteSpace(config.Ocr.TesseractPath) ? "tesseract" : config.Ocr.TesseractPath;
    var tesseract = UtilityInstaller.Probe(tesseractCommand, "--version");

    return Results.Ok(new
    {
        os = UtilityInstaller.GetOsName(),
        packageManager = UtilityInstaller.GetPreferredPackageManager(),
        ffmpeg,
        tesseract,
        ocrEnabled = config.OcrEnabled,
        configuredTesseractPath = config.Ocr.TesseractPath
    });
});

app.MapPost("/api/utilities/install", async (HttpContext context, UtilityInstallRequest request, AgentConfig config, ConfigFileStore store, RestartRequirementTracker restartTracker, CancellationToken cancellationToken) =>
{
    if (!RequestGuards.IsLoopback(context))
    {
        return Results.StatusCode(StatusCodes.Status403Forbidden);
    }

    var normalizedTool = (request.Tool ?? string.Empty).Trim().ToLowerInvariant();
    if (normalizedTool is not ("ffmpeg" or "ocr" or "tesseract"))
    {
        return Results.BadRequest(new { message = "Tool must be ffmpeg or ocr." });
    }

    var result = await UtilityInstaller.InstallAsync(normalizedTool, request.EnableOcr, config, store, restartTracker, cancellationToken);
    if (!result.Success)
    {
        return Results.BadRequest(new
        {
            message = result.Message,
            stdout = result.StdOut,
            stderr = result.StdErr,
            ffmpeg = UtilityInstaller.Probe("ffmpeg", "-version"),
            tesseract = UtilityInstaller.Probe(string.IsNullOrWhiteSpace(config.Ocr.TesseractPath) ? "tesseract" : config.Ocr.TesseractPath, "--version")
        });
    }

    return Results.Ok(new
    {
        message = result.Message,
        stdout = result.StdOut,
        stderr = result.StdErr,
        ffmpeg = UtilityInstaller.Probe("ffmpeg", "-version"),
        tesseract = UtilityInstaller.Probe(string.IsNullOrWhiteSpace(config.Ocr.TesseractPath) ? "tesseract" : config.Ocr.TesseractPath, "--version"),
        ocrEnabled = config.OcrEnabled,
        ocrRestartRequired = restartTracker.OcrRestartRequired
    });
});

app.MapPost("/api/utilities/enable-ocr", async (HttpContext context, UtilityEnableOcrRequest? request, AgentConfig config, ConfigFileStore store, RestartRequirementTracker restartTracker, CancellationToken cancellationToken) =>
{
    if (!RequestGuards.IsLoopback(context))
    {
        return Results.StatusCode(StatusCodes.Status403Forbidden);
    }

    var changed = false;
    if (!config.OcrEnabled)
    {
        config.OcrEnabled = true;
        changed = true;
    }

    var path = request?.TesseractPath?.Trim();
    if (!string.IsNullOrWhiteSpace(path) && !string.Equals(config.Ocr.TesseractPath, path, StringComparison.Ordinal))
    {
        config.Ocr.TesseractPath = path;
        changed = true;
    }

    if (changed)
    {
        restartTracker.OcrRestartRequired = true;
        await store.SaveAsync(config);
    }

    return Results.Ok(new
    {
        message = changed ? "OCR enabled. Restart required." : "OCR already enabled.",
        ocrEnabled = config.OcrEnabled,
        configuredTesseractPath = config.Ocr.TesseractPath,
        ocrRestartRequired = restartTracker.OcrRestartRequired,
        tesseract = UtilityInstaller.Probe(string.IsNullOrWhiteSpace(config.Ocr.TesseractPath) ? "tesseract" : config.Ocr.TesseractPath, "--version")
    });
});

app.MapGet("/api/tasks", async (TaskLibraryStore tasks, CancellationToken cancellationToken) =>
{
    var items = await tasks.ListAsync(cancellationToken);
    return Results.Ok(new { tasks = items });
});

app.MapPost("/api/tasks", async (TaskUpsertRequest request, TaskLibraryStore tasks, CancellationToken cancellationToken) =>
{
    var name = request.Name?.Trim() ?? string.Empty;
    var intent = request.Intent?.Trim() ?? string.Empty;
    var planJson = request.PlanJson?.Trim();
    if (string.IsNullOrWhiteSpace(name) || (string.IsNullOrWhiteSpace(intent) && string.IsNullOrWhiteSpace(planJson)))
    {
        return Results.BadRequest(new { message = "Task name and intent/plan are required." });
    }

    if (!string.IsNullOrWhiteSpace(planJson))
    {
        try
        {
            var parsed = JsonSerializer.Deserialize<ActionPlan>(planJson, new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true
            });
            if (parsed == null || parsed.Steps.Count == 0)
            {
                return Results.BadRequest(new { message = "Invalid plan JSON." });
            }
        }
        catch (Exception ex)
        {
            return Results.BadRequest(new { message = $"Invalid plan JSON: {ex.Message}" });
        }
    }

    await tasks.UpsertAsync(name, string.IsNullOrWhiteSpace(intent) ? $"recorded-macro:{name}" : intent, request.Description?.Trim(), planJson, cancellationToken);
    return Results.Ok(new { message = "Task saved." });
});

app.MapDelete("/api/tasks/{name}", async (string name, TaskLibraryStore tasks, CancellationToken cancellationToken) =>
{
    if (string.IsNullOrWhiteSpace(name))
    {
        return Results.BadRequest(new { message = "Task name is required." });
    }

    var deleted = await tasks.DeleteAsync(name, cancellationToken);
    return deleted ? Results.Ok(new { message = "Task deleted." }) : Results.NotFound(new { message = "Task not found." });
});

app.MapPost("/api/tasks/{name}/run", async (string name, RunTaskRequest? request, TaskLibraryStore tasks, IServiceProvider sp, CancellationToken cancellationToken) =>
{
    var task = await tasks.GetAsync(name, cancellationToken);
    if (task == null)
    {
        return Results.NotFound(new { message = "Task not found." });
    }

    if (!string.IsNullOrWhiteSpace(task.PlanJson))
    {
        try
        {
            var plan = JsonSerializer.Deserialize<ActionPlan>(task.PlanJson, new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
            if (plan == null || plan.Steps.Count == 0)
            {
                return Results.BadRequest(new { message = "Stored task plan is invalid." });
            }

            return await ExecutePlanAsync($"task:{name}", plan, request?.DryRun ?? false, sp);
        }
        catch (Exception ex)
        {
            return Results.BadRequest(new { message = $"Stored task plan error: {ex.Message}" });
        }
    }

    return await ExecuteIntentAsync(task.Intent, request?.DryRun ?? false, sp);
});

app.MapGet("/api/schedules", async (ScheduleLibraryStore schedules, CancellationToken cancellationToken) =>
{
    var items = await schedules.ListAsync(cancellationToken);
    return Results.Ok(new { schedules = items });
});

app.MapPost("/api/schedules", async (ScheduleUpsertRequest request, ScheduleLibraryStore schedules, TaskLibraryStore tasks, CancellationToken cancellationToken) =>
{
    var taskName = request.TaskName?.Trim() ?? string.Empty;
    if (string.IsNullOrWhiteSpace(taskName))
    {
        return Results.BadRequest(new { message = "Task name is required." });
    }

    var task = await tasks.GetAsync(taskName, cancellationToken);
    if (task == null)
    {
        return Results.BadRequest(new { message = $"Task '{taskName}' not found." });
    }

    if (request.IntervalSeconds is < 5 or > 86400)
    {
        return Results.BadRequest(new { message = "IntervalSeconds must be between 5 and 86400." });
    }

    var item = await schedules.UpsertAsync(
        request.Id,
        taskName,
        request.StartAtUtc,
        request.IntervalSeconds,
        request.Enabled ?? true,
        cancellationToken);

    return Results.Ok(new { message = "Schedule saved.", schedule = item });
});

app.MapDelete("/api/schedules/{id}", async (string id, ScheduleLibraryStore schedules, CancellationToken cancellationToken) =>
{
    if (string.IsNullOrWhiteSpace(id))
    {
        return Results.BadRequest(new { message = "Schedule id is required." });
    }

    var deleted = await schedules.DeleteAsync(id, cancellationToken);
    return deleted ? Results.Ok(new { message = "Schedule deleted." }) : Results.NotFound(new { message = "Schedule not found." });
});

app.MapPost("/api/schedules/{id}/run", async (string id, ScheduledTaskRunner runner, CancellationToken cancellationToken) =>
{
    if (string.IsNullOrWhiteSpace(id))
    {
        return Results.BadRequest(new { message = "Schedule id is required." });
    }

    var run = await runner.RunNowAsync(id, cancellationToken);
    if (!run.Found)
    {
        return Results.NotFound(new { message = run.Message });
    }

    return Results.Ok(new { message = run.Message, success = run.Success });
});

app.MapGet("/api/inspector", async (string? query, IDesktopAdapterClient client, CancellationToken cancellationToken) =>
{
    var active = await client.GetActiveWindowAsync(cancellationToken);
    var windows = await client.ListWindowsAsync(cancellationToken);
    var topWindows = windows.Take(15).Select(window => new
    {
        window.Id,
        window.Title,
        window.AppId,
        Bounds = new { window.Bounds.X, window.Bounds.Y, window.Bounds.Width, window.Bounds.Height }
    }).ToList();

    object? uiTree = null;
    if (active != null)
    {
        var tree = await client.GetUiTreeAsync(active.Id, cancellationToken);
        if (tree?.Root != null)
        {
            uiTree = SummarizeUiTree(tree.Root, maxNodes: 60);
        }
    }

    var matches = new List<object>();
    if (!string.IsNullOrWhiteSpace(query))
    {
        var selector = new DesktopAgent.Proto.Selector { NameContains = query.Trim() };
        var found = await client.FindElementsAsync(selector, cancellationToken);
        matches = found.Take(25).Select(element => (object)new
        {
            element.Id,
            element.Role,
            element.Name,
            element.AutomationId,
            element.ClassName,
            Bounds = new { element.Bounds.X, element.Bounds.Y, element.Bounds.Width, element.Bounds.Height }
        }).ToList();
    }

    return Results.Ok(new
    {
        activeWindow = active == null ? null : new
        {
            active.Id,
            active.Title,
            active.AppId,
            Bounds = new { active.Bounds.X, active.Bounds.Y, active.Bounds.Width, active.Bounds.Height }
        },
        windows = topWindows,
        uiTree,
        matches
    });
});

app.MapGet("/api/memory", (TargetMemoryStore memory) =>
{
    return Results.Ok(memory.GetSnapshot());
});

app.MapGet("/api/lock", (ContextLockStore contextLock) =>
{
    return Results.Ok(contextLock.GetSnapshot());
});

app.MapPost("/api/memory/clear", (TargetMemoryStore memory) =>
{
    memory.Clear();
    return Results.Ok(new { message = "Memory cleared." });
});

app.MapPost("/api/lock/clear", (ContextLockStore contextLock) =>
{
    contextLock.Unlock();
    return Results.Ok(new { message = "Context lock cleared." });
});

app.MapGet("/api/recording/status", (MacroRecorderStore recorder) =>
{
    return Results.Ok(recorder.GetStatus());
});

app.MapPost("/api/recording/start", (MacroRecordStartRequest? request, MacroRecorderStore recorder) =>
{
    recorder.Start(request?.Name);
    return Results.Ok(recorder.GetStatus());
});

app.MapPost("/api/recording/stop", (MacroRecorderStore recorder) =>
{
    recorder.Stop();
    var plan = recorder.BuildRecordedPlan();
    return Results.Ok(new
    {
        status = recorder.GetStatus(),
        planJson = plan == null ? null : PlanToJson(plan)
    });
});

app.MapPost("/api/recording/save", async (MacroRecordSaveRequest request, MacroRecorderStore recorder, TaskLibraryStore tasks, CancellationToken cancellationToken) =>
{
    var name = request.Name?.Trim() ?? string.Empty;
    if (string.IsNullOrWhiteSpace(name))
    {
        return Results.BadRequest(new { message = "Task name is required." });
    }

    var plan = recorder.BuildRecordedPlan();
    if (plan == null || plan.Steps.Count == 0)
    {
        return Results.BadRequest(new { message = "No recorded steps to save." });
    }

    var planJson = PlanToJson(plan);
    var intent = $"recorded-macro:{name}";
    await tasks.UpsertAsync(name, intent, request.Description?.Trim(), planJson, cancellationToken);
    return Results.Ok(new { message = "Recorded macro saved.", task = name });
});

app.MapPost("/api/translate", async (TranslateRequest request, AgentConfig config, IAuditLog auditLog, CancellationToken cancellationToken) =>
{
    var text = request.Text?.Trim();
    var target = request.TargetLanguage?.Trim();
    if (string.IsNullOrWhiteSpace(text) || string.IsNullOrWhiteSpace(target))
    {
        return Results.BadRequest(new { message = "Text and targetLanguage are required." });
    }

    var source = string.IsNullOrWhiteSpace(request.SourceLanguage) ? null : request.SourceLanguage.Trim();
    var result = await LlmTranslationService.TranslateAsync(new TranslationIntent(text, target, source), config, auditLog, cancellationToken);
    if (!result.Success)
    {
        return Results.BadRequest(new { message = result.Message });
    }

    return Results.Ok(new
    {
        translation = result.TranslatedText,
        provider = result.Provider,
        model = config.LlmFallback.Model
    });
});

app.MapPost("/api/chat", async (ChatRequest request, IDesktopAdapterClient client, ChatActionStore store, IAuditLog auditLog, AgentConfig config, IServiceProvider sp, TargetMemoryStore memory, MacroRecorderStore recorder, ContextLockStore contextLock, IKillSwitch killSwitch) =>
{
    var message = request.Message?.Trim() ?? string.Empty;
    if (string.IsNullOrWhiteSpace(message))
    {
        return Results.Ok(ChatResponse.Simple("Enter a valid command."));
    }

    var normalized = message.ToLowerInvariant();

    if (normalized is "kill" or "kill switch" or "panic" or "panic stop" or "stop now" or "abort")
    {
        killSwitch.Trip("Kill requested from chat");
        await auditLog.WriteAsync(new AuditEvent { Timestamp = DateTimeOffset.UtcNow, EventType = "kill", Message = "Kill switch tripped by user" }, CancellationToken.None);
        return Results.Ok(ChatResponse.Simple("Kill switch enabled. Running actions will stop immediately."));
    }

    if (normalized is "reset kill" or "clear kill" or "unkill")
    {
        killSwitch.Reset();
        await auditLog.WriteAsync(new AuditEvent { Timestamp = DateTimeOffset.UtcNow, EventType = "kill_reset", Message = "Kill switch reset by user" }, CancellationToken.None);
        return Results.Ok(ChatResponse.Simple("Kill switch reset."));
    }

    if (normalized.Contains("status"))
    {
        var status = await client.GetStatusAsync(CancellationToken.None);
        var lockStatus = FormatContextLock(contextLock.GetSnapshot());
        var killStatus = killSwitch.IsTripped
            ? $"Kill switch: ON ({killSwitch.Reason ?? "manual"})."
            : "Kill switch: OFF.";
        return Results.Ok(ChatResponse.Simple($"Adapter armed: {status.Armed}, require presence: {status.RequireUserPresence}. {killStatus} {lockStatus}"));
    }

    if (normalized is "lock status" or "context lock status")
    {
        return Results.Ok(ChatResponse.Simple(FormatContextLock(contextLock.GetSnapshot())));
    }

    if (normalized is "unlock" or "unlock context" or "unlock context lock" or "release lock")
    {
        contextLock.Unlock();
        await auditLog.WriteAsync(new AuditEvent { Timestamp = DateTimeOffset.UtcNow, EventType = "context_unlock", Message = "Context lock disabled" }, CancellationToken.None);
        return Results.Ok(ChatResponse.Simple("Context lock disabled."));
    }

    if (normalized is "lock current" or "lock current window" or "lock on current" or "lock on current window")
    {
        var active = await client.GetActiveWindowAsync(CancellationToken.None);
        if (active == null)
        {
            return Results.Ok(ChatResponse.Simple("Cannot lock current window: no active window detected."));
        }

        contextLock.LockToWindow(active);
        await auditLog.WriteAsync(new AuditEvent { Timestamp = DateTimeOffset.UtcNow, EventType = "context_lock", Message = $"Context locked to window {active.Id}", Data = new { active.Id, active.AppId, active.Title } }, CancellationToken.None);
        return Results.Ok(ChatResponse.Simple($"Context locked to window '{active.Title}' [{active.AppId}] ({active.Id})."));
    }

    if (normalized.StartsWith("lock on "))
    {
        var target = message["lock on ".Length..].Trim();
        if (string.IsNullOrWhiteSpace(target))
        {
            return Results.Ok(ChatResponse.Simple("Specify target: lock on current window | lock on <app>."));
        }

        var lowerTarget = target.ToLowerInvariant();
        if (lowerTarget is "current" or "current window" or "this window")
        {
            var active = await client.GetActiveWindowAsync(CancellationToken.None);
            if (active == null)
            {
                return Results.Ok(ChatResponse.Simple("Cannot lock current window: no active window detected."));
            }

            contextLock.LockToWindow(active);
            await auditLog.WriteAsync(new AuditEvent { Timestamp = DateTimeOffset.UtcNow, EventType = "context_lock", Message = $"Context locked to window {active.Id}", Data = new { active.Id, active.AppId, active.Title } }, CancellationToken.None);
            return Results.Ok(ChatResponse.Simple($"Context locked to window '{active.Title}' [{active.AppId}] ({active.Id})."));
        }

        if (lowerTarget.StartsWith("window "))
        {
            var windowId = target["window ".Length..].Trim();
            if (string.IsNullOrWhiteSpace(windowId))
            {
                return Results.Ok(ChatResponse.Simple("Specify window id: lock on window <id>."));
            }

            contextLock.LockToWindowId(windowId);
            await auditLog.WriteAsync(new AuditEvent { Timestamp = DateTimeOffset.UtcNow, EventType = "context_lock", Message = $"Context locked to window id {windowId}" }, CancellationToken.None);
            return Results.Ok(ChatResponse.Simple($"Context locked to window id '{windowId}'."));
        }

        var appTarget = target;
        var resolver = sp.GetRequiredService<IAppResolver>();
        if (resolver.TryResolveApp(target, out var resolvedTarget))
        {
            appTarget = resolvedTarget;
        }

        contextLock.LockToApp(appTarget);
        await auditLog.WriteAsync(new AuditEvent { Timestamp = DateTimeOffset.UtcNow, EventType = "context_lock", Message = $"Context locked to app {appTarget}" }, CancellationToken.None);
        return Results.Ok(ChatResponse.Simple($"Context locked to app '{appTarget}'."));
    }

    if (normalized.StartsWith("profile "))
    {
        var profile = AgentProfileService.NormalizeProfile(message["profile ".Length..].Trim());
        config.ActiveProfile = profile;
        config.ProfileModeEnabled = true;
        AgentProfileService.ApplyActiveProfile(config);
        return Results.Ok(ChatResponse.Simple($"Profile set to {profile}. Limits/policy updated."));
    }

    if (normalized.StartsWith("arm"))
    {
        killSwitch.Reset();
        var status = await client.ArmAsync(true, CancellationToken.None);
        await auditLog.WriteAsync(new AuditEvent { Timestamp = DateTimeOffset.UtcNow, EventType = "arm", Message = "Adapter armed" }, CancellationToken.None);
        return Results.Ok(ChatResponse.Simple($"Armed: {status.Armed}, require presence: {status.RequireUserPresence}. Kill switch reset."));
    }

    if (normalized.StartsWith("disarm"))
    {
        killSwitch.Trip("Disarm requested from chat");
        var status = await client.DisarmAsync(CancellationToken.None);
        await auditLog.WriteAsync(new AuditEvent { Timestamp = DateTimeOffset.UtcNow, EventType = "disarm", Message = "Adapter disarmed" }, CancellationToken.None);
        return Results.Ok(ChatResponse.Simple($"Armed: {status.Armed}. Running actions stopped."));
    }

    if (normalized.Contains("simulate presence"))
    {
        var token = store.CreatePending(ChatActionType.SimulatePresence, "Simulate presence: the adapter will be armed without user presence.", null, false);
        return Results.Ok(ChatResponse.Confirm("Simulate presence? This will arm the adapter without user presence.", token));
    }

    if (normalized.Contains("require presence"))
    {
        killSwitch.Reset();
        var status = await client.ArmAsync(true, CancellationToken.None);
        await auditLog.WriteAsync(new AuditEvent { Timestamp = DateTimeOffset.UtcNow, EventType = "presence_required", Message = "Require presence enabled" }, CancellationToken.None);
        return Results.Ok(ChatResponse.Simple($"Presence required. Armed: {status.Armed}. Kill switch reset."));
    }

    if (normalized.StartsWith("list apps") || normalized.StartsWith("apps"))
    {
        var rawArgs = normalized.StartsWith("list apps")
            ? message.Substring("list apps".Length)
            : message.Substring("apps".Length);
        var tokens = rawArgs.Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).ToList();
        var allowFlags = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "allowed", "allowlist", "--allowed", "--allowed-only", "--allowlist"
        };
        var allowedOnly = tokens.RemoveAll(token => allowFlags.Contains(token)) > 0;
        var query = string.Join(" ", tokens);
        var resolver = sp.GetRequiredService<IAppResolver>();
        var matches = resolver.Suggest(query, 12);
        if (allowedOnly)
        {
            matches = matches.Where(match => match.IsAllowed).ToList();
        }
        if (matches.Count == 0)
        {
            return Results.Ok(ChatResponse.Simple("No apps found."));
        }

        var lines = matches.Select(m =>
        {
            var tag = m.IsAllowed ? " [allowed]" : string.Empty;
            return $"{m.Entry.Name}{tag} score={m.Score:0.00} ({m.Entry.Path})";
        }).ToList();
        return Results.Ok(ChatResponse.WithSteps("Top apps:", lines, null, null));
    }

    if (normalized is "memory" or "show memory" or "last targets")
    {
        var snapshot = memory.GetSnapshot();
        var lines = new List<string>
        {
            $"Last window: {(string.IsNullOrWhiteSpace(snapshot.LastWindowId) ? "none" : $"{snapshot.LastWindowTitle} [{snapshot.LastWindowAppId}] ({snapshot.LastWindowId})")}",
            $"Last element: {(string.IsNullOrWhiteSpace(snapshot.LastElementId) ? "none" : $"{snapshot.LastElementName} role={snapshot.LastElementRole} id={snapshot.LastElementId}")}"
        };
        return Results.Ok(ChatResponse.WithSteps("Memory snapshot", lines, null, null));
    }

    if (normalized is "start recording" or "record start" or "start macro recording")
    {
        recorder.Start(null);
        return Results.Ok(ChatResponse.Simple("Recording started."));
    }

    if (normalized is "stop recording" or "record stop" or "stop macro recording")
    {
        recorder.Stop();
        var plan = recorder.BuildRecordedPlan();
        var count = plan?.Steps.Count ?? 0;
        return Results.Ok(ChatResponse.Simple($"Recording stopped. Captured steps: {count}."));
    }

    if (normalized.StartsWith("save recording"))
    {
        var name = message["save recording".Length..].Trim();
        if (string.IsNullOrWhiteSpace(name))
        {
            return Results.Ok(ChatResponse.Simple("Specify a task name. Example: save recording my-macro"));
        }

        var plan = recorder.BuildRecordedPlan();
        if (plan == null || plan.Steps.Count == 0)
        {
            return Results.Ok(ChatResponse.Simple("No recorded steps to save."));
        }

        var tasks = sp.GetRequiredService<TaskLibraryStore>();
        var planJson = PlanToJson(plan);
        await tasks.UpsertAsync(name, $"recorded-macro:{name}", "Recorded from chat", planJson, CancellationToken.None);
        return Results.Ok(ChatResponse.Simple($"Recording saved as task '{name}'."));
    }

    if (normalized is "click last" or "click last element")
    {
        var plan = memory.BuildClickLastPlan();
        if (plan == null)
        {
            return Results.Ok(ChatResponse.Simple("No remembered element to click."));
        }

        return await ExecutePlanAsync("click last", plan, dryRun: false, sp);
    }

    if (TryParseTypeInLast(message, out var textInLast))
    {
        var plan = memory.BuildTypeInLastPlan(textInLast);
        if (plan == null)
        {
            return Results.Ok(ChatResponse.Simple("No remembered editable element. Use 'find ...' first."));
        }

        return await ExecutePlanAsync("type in last", plan, dryRun: false, sp);
    }

    if (TryParseTranslationIntent(message, out var translationIntent))
    {
        var translation = await LlmTranslationService.TranslateAsync(translationIntent, config, auditLog, CancellationToken.None);
        if (!translation.Success)
        {
            return Results.Ok(ChatResponse.Simple(translation.Message));
        }

        return Results.Ok(ChatResponse.Simple($"Translation ({translation.Provider}):\n{translation.TranslatedText}"));
    }

    if (IsDirectIntent(normalized))
    {
        return await ExecuteIntentAsync(message, dryRun: false, sp);
    }

    if (normalized.StartsWith("run "))
    {
        var intent = message.Substring("run ".Length).Trim();
        return await ExecuteIntentAsync(intent, dryRun: false, sp);
    }

    if (normalized.StartsWith("dry-run ") || normalized.StartsWith("dryrun "))
    {
        var intent = message.Contains(' ') ? message[(message.IndexOf(' ') + 1)..].Trim() : string.Empty;
        return await ExecuteIntentAsync(intent, dryRun: true, sp);
    }

    var planner = sp.GetRequiredService<IPlanner>();
    var inferredPlan = planner.PlanFromIntent(message);
    if (!IsUnrecognizedPlan(inferredPlan))
    {
        var token = store.CreatePending(ChatActionType.ExecutePlan, "Confirm free-form execution", inferredPlan, false);
        var notice = GetRewriteNotice(inferredPlan);
        var prompt = string.IsNullOrWhiteSpace(notice)
            ? "I interpreted your request. Confirm execution?"
            : $"I interpreted your request. {notice}. Confirm execution?";
        return Results.Ok(ChatResponse.ConfirmWithSteps(prompt, token, PlanToLines(inferredPlan), PlanToJson(inferredPlan), GetModeLabel(inferredPlan)));
    }

    return Results.Ok(ChatResponse.Simple("Available commands: status, kill, reset kill, lock status, lock on <current window|app>, unlock, profile <safe|balanced|power>, arm, disarm, simulate presence, require presence, list apps [query] [allowed], memory, click last, type in last <text>, start/stop recording, save recording <name>, run <intent>, dry-run <intent>, translate <text> to <language> (or 'translate to <language>: <text>'). Supported intents: open/run/start, find/click/type/press, double click/right click/drag, wait until <text> [for <seconds>], window actions (minimize/maximize/restore/switch/focus), scroll/page/home/end, browser back/forward/refresh/find in page, file write/read/list/append, open url/search (on specific browser), notify, clipboard history, jiggle mouse for <duration>, volume up/down/mute, brightness up/down, lock screen."));
});

app.MapPost("/api/confirm", async (ConfirmRequest request, IDesktopAdapterClient client, ChatActionStore store, IAuditLog auditLog, IServiceProvider sp, IKillSwitch killSwitch) =>
{
    if (string.IsNullOrWhiteSpace(request.Token))
    {
        return Results.BadRequest(new { message = "Token missing" });
    }

    var pending = store.Take(request.Token);
    if (pending == null)
    {
        return Results.NotFound(new { message = "Token not found" });
    }

    if (!request.Approve)
    {
        return Results.Ok(ChatResponse.Simple("Cancelled."));
    }

    switch (pending.Type)
    {
        case ChatActionType.SimulatePresence:
            killSwitch.Reset();
            var status = await client.ArmAsync(false, CancellationToken.None);
            await auditLog.WriteAsync(new AuditEvent { Timestamp = DateTimeOffset.UtcNow, EventType = "simulate_presence", Message = "Presence simulation enabled" }, CancellationToken.None);
            return Results.Ok(ChatResponse.Simple($"Presence simulated. Armed: {status.Armed}, require presence: {status.RequireUserPresence}. Kill switch reset."));
        case ChatActionType.ExecutePlan:
            var planToExecute = pending.Plan;
            if (!string.IsNullOrWhiteSpace(request.PlanJson))
            {
                if (!TryParseActionPlanJson(request.PlanJson, out var editedPlan, out var parseError))
                {
                    return Results.BadRequest(new { message = $"Invalid edited plan JSON: {parseError}" });
                }

                planToExecute = editedPlan;
            }

            if (planToExecute == null)
            {
                return Results.Ok(ChatResponse.Simple("Plan missing."));
            }
            var executor = CreateExecutor(sp, new AutoApproveConfirmation());
            var execResult = await executor.ExecutePlanAsync(planToExecute, pending.DryRun, CancellationToken.None);
            await CaptureExecutionSideEffectsAsync(pending.Message, planToExecute, execResult, pending.DryRun, sp);
            var notice = GetRewriteNotice(planToExecute);
            var reply = AppendNotice(FormatExecution(execResult), notice);
            return Results.Ok(ChatResponse.WithSteps(reply, ExecutionToLines(execResult), PlanToJson(planToExecute), GetModeLabel(planToExecute)));
        default:
            return Results.Ok(ChatResponse.Simple("Action not supported."));
    }
});

app.MapPost("/api/intent", async (IntentRequest request, IServiceProvider sp) =>
{
    var intent = request.Intent?.Trim() ?? string.Empty;
    if (string.IsNullOrWhiteSpace(intent))
    {
        return Results.BadRequest(new { message = "Intent missing" });
    }
    return await ExecuteIntentAsync(intent, request.DryRun, sp);
});

app.MapGet("/api/audit", (int? take, AgentConfig config) =>
{
    var path = Path.GetFullPath(config.AuditLogPath);
    if (!File.Exists(path))
    {
        return Results.Ok(new { lines = Array.Empty<string>() });
    }

    var max = take.GetValueOrDefault(100);
    max = Math.Clamp(max, 1, 500);
    var queue = new Queue<string>();
    foreach (var line in File.ReadLines(path))
    {
        if (queue.Count == max)
        {
            queue.Dequeue();
        }
        queue.Enqueue(line);
    }

    return Results.Ok(new { lines = queue.ToArray() });
});

app.MapGet("/api/audit/stream", async (HttpContext context, AgentConfig config) =>
{
    context.Response.Headers.ContentType = "text/event-stream";
    context.Response.Headers.CacheControl = "no-cache";
    context.Response.Headers.Connection = "keep-alive";

    var path = Path.GetFullPath(config.AuditLogPath);
    long lastSize = 0;
    var cancellation = context.RequestAborted;

    while (!cancellation.IsCancellationRequested)
    {
        if (File.Exists(path))
        {
            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            if (stream.Length < lastSize)
            {
                lastSize = 0;
            }

            if (stream.Length > lastSize)
            {
                stream.Seek(lastSize, SeekOrigin.Begin);
                using var reader = new StreamReader(stream);
                string? line;
                while ((line = await reader.ReadLineAsync()) != null)
                {
                    await context.Response.WriteAsync($"data: {line}\n\n", cancellation);
                }
                lastSize = stream.Position;
                await context.Response.Body.FlushAsync(cancellation);
            }
        }

        await Task.Delay(1000, cancellation);
    }
});

app.Run();

static IExecutor CreateExecutor(IServiceProvider sp, IUserConfirmation confirmation)
{
    return new Executor(
        sp.GetRequiredService<IDesktopAdapterClient>(),
        sp.GetRequiredService<IContextProvider>(),
        sp.GetRequiredService<IAppResolver>(),
        sp.GetRequiredService<IPolicyEngine>(),
        sp.GetRequiredService<IRateLimiter>(),
        sp.GetRequiredService<IAuditLog>(),
        confirmation,
        sp.GetRequiredService<IKillSwitch>(),
        sp.GetRequiredService<AgentConfig>(),
        sp.GetRequiredService<IOcrEngine>(),
        sp.GetRequiredService<ILogger<Executor>>());
}

static string FormatExecution(ExecutionResult result)
{
    var summary = $"Success: {result.Success}. {result.Message}";
    if (result.Steps.Count == 0)
    {
        return summary;
    }

    var details = string.Join(" | ", result.Steps.Select(step => $"{step.Type}:{step.Success}"));
    return $"{summary} Steps: {details}";
}

static string AppendNotice(string reply, string? notice)
{
    if (string.IsNullOrWhiteSpace(notice))
    {
        return reply;
    }

    return $"{reply} {notice}";
}

static string? GetRewriteNotice(ActionPlan plan)
{
    var note = plan.Steps.Select(step => step.Note)
        .FirstOrDefault(value => value != null && value.StartsWith("Rewritten intent:", StringComparison.OrdinalIgnoreCase));
    return note;
}

static string GetModeLabel(ActionPlan plan)
{
    var rewritten = plan.Steps.Select(step => step.Note)
        .FirstOrDefault(value => value != null && value.StartsWith("Rewritten intent:", StringComparison.OrdinalIgnoreCase));
    return string.IsNullOrWhiteSpace(rewritten) ? "Mode: Rule-based" : "Mode: LLM interpreter";
}

static async Task<IResult> ExecuteIntentAsync(string intent, bool dryRun, IServiceProvider sp)
{
    if (string.IsNullOrWhiteSpace(intent))
    {
        return Results.Ok(ChatResponse.Simple("Intent missing."));
    }

    var planner = sp.GetRequiredService<IPlanner>();
    var plan = planner.PlanFromIntent(intent);
    return await ExecutePlanAsync(intent, plan, dryRun, sp);
}

static async Task<IResult> ExecutePlanAsync(string source, ActionPlan plan, bool dryRun, IServiceProvider sp)
{
    var policy = sp.GetRequiredService<IPolicyEngine>();
    var client = sp.GetRequiredService<IDesktopAdapterClient>();
    var contextLock = sp.GetRequiredService<ContextLockStore>();
    var activeWindow = await client.GetActiveWindowAsync(CancellationToken.None);
    var contextLockState = contextLock.GetSnapshot();

    if (contextLockState.Enabled)
    {
        if (!MatchesContextLock(activeWindow, contextLockState))
        {
            return Results.Ok(ChatResponse.Simple($"Blocked: context lock active ({FormatContextLock(contextLockState)})."));
        }

        var lockReason = ApplyContextLock(plan, contextLockState);
        if (!string.IsNullOrWhiteSpace(lockReason))
        {
            return Results.Ok(ChatResponse.Simple($"Blocked: {lockReason}"));
        }
    }

    foreach (var step in plan.Steps)
    {
        var decision = policy.Evaluate(step, activeWindow);
        if (!decision.Allowed)
        {
            return Results.Ok(ChatResponse.Simple($"Blocked: {decision.Reason}"));
        }
        if (decision.RequiresConfirmation)
        {
            var store = sp.GetRequiredService<ChatActionStore>();
            var token = store.CreatePending(ChatActionType.ExecutePlan, "Confirm execution", plan, dryRun);
            return Results.Ok(ChatResponse.ConfirmWithSteps($"Confirm plan execution for: {source}", token, PlanToLines(plan), PlanToJson(plan), GetModeLabel(plan)));
        }
    }

    var executor = CreateExecutor(sp, new AutoApproveConfirmation());
    var result = await executor.ExecutePlanAsync(plan, dryRun, CancellationToken.None);
    await CaptureExecutionSideEffectsAsync(source, plan, result, dryRun, sp);
    var notice = GetRewriteNotice(plan);
    var reply = AppendNotice(FormatExecution(result), notice);
    return Results.Ok(ChatResponse.WithSteps(reply, ExecutionToLines(result), PlanToJson(plan), GetModeLabel(plan)));
}

static async Task CaptureExecutionSideEffectsAsync(string source, ActionPlan plan, ExecutionResult result, bool dryRun, IServiceProvider sp)
{
    if (dryRun || plan.Steps.Count == 0)
    {
        return;
    }

    var client = sp.GetRequiredService<IDesktopAdapterClient>();
    var memory = sp.GetRequiredService<TargetMemoryStore>();
    var recorder = sp.GetRequiredService<MacroRecorderStore>();

    WindowRef? active = null;
    try
    {
        active = await client.GetActiveWindowAsync(CancellationToken.None);
    }
    catch
    {
        // Keep side-effects best-effort.
    }

    memory.Capture(plan, result, active);
    recorder.Capture(source, plan, result);
}

static IReadOnlyList<string> PlanToLines(ActionPlan plan)
{
    var lines = new List<string>();
    for (var i = 0; i < plan.Steps.Count; i++)
    {
        var step = plan.Steps[i];
        lines.Add($"{i + 1}. {DescribeStep(step)}");
    }
    return lines;
}

static IReadOnlyList<string> ExecutionToLines(ExecutionResult result)
{
    var lines = new List<string>();
    for (var i = 0; i < result.Steps.Count; i++)
    {
        var step = result.Steps[i];
        var line = $"{i + 1}. {step.Type} => {step.Success} ({step.Message})";
        var hint = BuildStepDataHint(step);
        if (!string.IsNullOrWhiteSpace(hint))
        {
            line = $"{line} | {hint}";
        }
        else if (step.Data != null)
        {
            line = $"{line} data={ToInlineJson(step.Data, 300)}";
        }
        lines.Add(line);
    }
    return lines;
}

static string? BuildStepDataHint(StepResult step)
{
    if (step.Data == null)
    {
        return null;
    }

    if (!TrySerializeToElement(step.Data, out var element))
    {
        return null;
    }

    switch (step.Type)
    {
        case ActionType.ClipboardHistory:
            if (element.ValueKind == JsonValueKind.Array)
            {
                var items = element.EnumerateArray().ToList();
                var snippets = new List<string>();
                foreach (var entry in items.Take(3))
                {
                    var text = entry.TryGetProperty("text", out var textEl) ? textEl.GetString() ?? string.Empty : string.Empty;
                    text = text.Replace(Environment.NewLine, " ").Trim();
                    if (text.Length > 32)
                    {
                        text = text[..32] + "...";
                    }

                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        snippets.Add($"\"{text}\"");
                    }
                }

                return snippets.Count > 0
                    ? $"clipboard items={items.Count}, latest: {string.Join(", ", snippets)}"
                    : $"clipboard items={items.Count}";
            }
            break;

        case ActionType.FileRead:
            if (element.ValueKind == JsonValueKind.Object)
            {
                var path = element.TryGetProperty("path", out var pathEl) ? pathEl.GetString() ?? string.Empty : string.Empty;
                var content = element.TryGetProperty("content", out var contentEl) ? contentEl.GetString() ?? string.Empty : string.Empty;
                content = content.Replace(Environment.NewLine, " ").Trim();
                if (content.Length > 80)
                {
                    content = content[..80] + "...";
                }
                return $"file={path}, preview=\"{content}\"";
            }
            break;

        case ActionType.FileList:
            if (element.ValueKind == JsonValueKind.Object && element.TryGetProperty("entries", out var entriesEl) && entriesEl.ValueKind == JsonValueKind.Array)
            {
                var entries = entriesEl.EnumerateArray().Select(item => item.GetString()).Where(item => !string.IsNullOrWhiteSpace(item)).ToList();
                var preview = string.Join(", ", entries.Take(5));
                return entries.Count > 0
                    ? $"entries={entries.Count}: {preview}"
                    : "entries=0";
            }
            break;

        case ActionType.Find:
            if (element.ValueKind == JsonValueKind.Array)
            {
                var names = new List<string>();
                var total = 0;
                foreach (var item in element.EnumerateArray())
                {
                    total++;
                    if (names.Count >= 4)
                    {
                        continue;
                    }

                    var name = item.TryGetProperty("name", out var nameEl) ? nameEl.GetString() : null;
                    if (!string.IsNullOrWhiteSpace(name))
                    {
                        names.Add(name.Trim());
                    }
                }

                return names.Count > 0
                    ? $"matches={total}: {string.Join(", ", names)}"
                    : $"matches={total}";
            }
            break;

        case ActionType.ReadText:
            if (element.ValueKind == JsonValueKind.Array)
            {
                var texts = new List<string>();
                foreach (var item in element.EnumerateArray())
                {
                    if (texts.Count >= 3)
                    {
                        break;
                    }

                    var text = item.TryGetProperty("text", out var textEl) ? textEl.GetString() : null;
                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        text = text.Replace(Environment.NewLine, " ").Trim();
                        if (text.Length > 28)
                        {
                            text = text[..28] + "...";
                        }
                        texts.Add($"\"{text}\"");
                    }
                }

                return texts.Count > 0 ? $"ocr: {string.Join(", ", texts)}" : null;
            }
            break;

        case ActionType.MouseJiggle:
            if (element.ValueKind == JsonValueKind.Object)
            {
                var moves = element.TryGetProperty("moves", out var movesEl) ? movesEl.GetInt32() : 0;
                var seconds = element.TryGetProperty("seconds", out var secondsEl) ? secondsEl.GetDouble() : 0;
                return $"moves={moves}, duration={seconds:0.#}s";
            }
            break;
    }

    return null;
}

static bool TrySerializeToElement(object data, out JsonElement element)
{
    try
    {
        element = JsonSerializer.SerializeToElement(data);
        return true;
    }
    catch
    {
        element = default;
        return false;
    }
}

static string ToInlineJson(object value, int maxChars)
{
    try
    {
        var json = JsonSerializer.Serialize(value);
        if (json.Length <= maxChars)
        {
            return json;
        }

        return json[..maxChars] + "...";
    }
    catch
    {
        return "\"<unavailable>\"";
    }
}

static string DescribeStep(PlanStep step)
{
    var parts = new List<string> { step.Type.ToString() };
    if (!string.IsNullOrWhiteSpace(step.Text))
    {
        parts.Add($"text=\"{step.Text}\"");
    }
    if (!string.IsNullOrWhiteSpace(step.Target))
    {
        parts.Add($"target=\"{step.Target}\"");
    }
    if (!string.IsNullOrWhiteSpace(step.AppIdOrPath))
    {
        parts.Add($"app=\"{step.AppIdOrPath}\"");
    }
    if (step.Keys is { Count: > 0 })
    {
        parts.Add($"keys={string.Join('+', step.Keys)}");
    }
    if (step.Selector != null)
    {
        if (!string.IsNullOrWhiteSpace(step.Selector.Role))
            parts.Add($"role={step.Selector.Role}");
        if (!string.IsNullOrWhiteSpace(step.Selector.NameContains))
            parts.Add($"name~=\"{step.Selector.NameContains}\"");
        if (!string.IsNullOrWhiteSpace(step.Selector.AutomationId))
            parts.Add($"automationId={step.Selector.AutomationId}");
        if (!string.IsNullOrWhiteSpace(step.Selector.ClassName))
            parts.Add($"class={step.Selector.ClassName}");
        if (!string.IsNullOrWhiteSpace(step.Selector.AncestorNameContains))
            parts.Add($"ancestor~=\"{step.Selector.AncestorNameContains}\"");
    }
    return string.Join(" ", parts);
}

static string? ApplyContextLock(ActionPlan plan, ContextLockSnapshot snapshot)
{
    if (!snapshot.Enabled)
    {
        return null;
    }

    foreach (var step in plan.Steps)
    {
        if (step.Type == ActionType.OpenApp)
        {
            return "Context lock active: unlock before switching/opening apps.";
        }

        if (!RequiresUiContext(step.Type))
        {
            continue;
        }

        if (!string.IsNullOrWhiteSpace(snapshot.AppId))
        {
            if (!string.IsNullOrWhiteSpace(step.ExpectedAppId) && !AppIdMatches(step.ExpectedAppId, snapshot.AppId))
            {
                return $"Context lock mismatch on app: expected '{snapshot.AppId}'.";
            }

            step.ExpectedAppId ??= snapshot.AppId;
        }

        if (!string.IsNullOrWhiteSpace(snapshot.WindowId))
        {
            if (!string.IsNullOrWhiteSpace(step.ExpectedWindowId)
                && !string.Equals(step.ExpectedWindowId, snapshot.WindowId, StringComparison.OrdinalIgnoreCase))
            {
                return $"Context lock mismatch on window: expected '{snapshot.WindowId}'.";
            }

            step.ExpectedWindowId ??= snapshot.WindowId;
        }
    }

    return null;
}

static bool MatchesContextLock(WindowRef? activeWindow, ContextLockSnapshot snapshot)
{
    if (!snapshot.Enabled)
    {
        return true;
    }

    if (activeWindow == null)
    {
        return false;
    }

    if (!string.IsNullOrWhiteSpace(snapshot.AppId) && !AppIdMatches(activeWindow.AppId, snapshot.AppId))
    {
        return false;
    }

    if (!string.IsNullOrWhiteSpace(snapshot.WindowId)
        && !string.Equals(activeWindow.Id, snapshot.WindowId, StringComparison.OrdinalIgnoreCase))
    {
        return false;
    }

    return true;
}

static string FormatContextLock(ContextLockSnapshot snapshot)
{
    if (!snapshot.Enabled)
    {
        return "Context lock: off.";
    }

    if (string.Equals(snapshot.Scope, "window", StringComparison.OrdinalIgnoreCase))
    {
        var title = string.IsNullOrWhiteSpace(snapshot.WindowTitle) ? "window" : snapshot.WindowTitle;
        return $"Context lock: window '{title}' ({snapshot.WindowId ?? "n/a"}).";
    }

    return $"Context lock: app '{snapshot.AppId ?? "n/a"}'.";
}

static bool RequiresUiContext(ActionType type)
{
    return type is ActionType.Find
        or ActionType.Click
        or ActionType.TypeText
        or ActionType.KeyCombo
        or ActionType.Invoke
        or ActionType.SetValue
        or ActionType.ReadText
        or ActionType.CaptureScreen;
}

static bool AppIdMatches(string? actual, string? expected)
{
    if (string.IsNullOrWhiteSpace(actual) || string.IsNullOrWhiteSpace(expected))
    {
        return false;
    }

    var actualNormalized = NormalizeAppToken(actual);
    var expectedNormalized = NormalizeAppToken(expected);
    if (string.Equals(actualNormalized, expectedNormalized, StringComparison.OrdinalIgnoreCase))
    {
        return true;
    }

    return actualNormalized.Contains(expectedNormalized, StringComparison.OrdinalIgnoreCase)
        || expectedNormalized.Contains(actualNormalized, StringComparison.OrdinalIgnoreCase);
}

static string NormalizeAppToken(string value)
{
    var trimmed = value.Trim().Trim('"', '\'');
    if (string.IsNullOrWhiteSpace(trimmed))
    {
        return string.Empty;
    }

    if (trimmed.Contains(Path.DirectorySeparatorChar)
        || trimmed.Contains(Path.AltDirectorySeparatorChar))
    {
        var fileName = Path.GetFileNameWithoutExtension(trimmed);
        if (!string.IsNullOrWhiteSpace(fileName))
        {
            trimmed = fileName;
        }
    }

    return trimmed.Replace(" ", string.Empty).ToLowerInvariant();
}

static string PlanToJson(ActionPlan plan)
{
    return JsonSerializer.Serialize(plan, new JsonSerializerOptions
    {
        WriteIndented = true
    });
}

static bool TryParseActionPlanJson(string? planJson, out ActionPlan? plan, out string error)
{
    plan = null;
    error = string.Empty;

    var json = planJson?.Trim();
    if (string.IsNullOrWhiteSpace(json))
    {
        error = "Empty plan JSON.";
        return false;
    }

    try
    {
        plan = JsonSerializer.Deserialize<ActionPlan>(json, new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true
        });
        if (plan == null || plan.Steps.Count == 0)
        {
            error = "Plan is empty.";
            plan = null;
            return false;
        }

        return true;
    }
    catch (Exception ex)
    {
        error = ex.Message;
        return false;
    }
}

static UiTreeSummaryNode SummarizeUiTree(DesktopAgent.Proto.UiNode root, int maxNodes)
{
    var limit = Math.Clamp(maxNodes, 1, 1000);
    var count = 0;
    return SummarizeNode(root, ref count, limit, 0);
}

static UiTreeSummaryNode SummarizeNode(DesktopAgent.Proto.UiNode node, ref int count, int maxNodes, int depth)
{
    count++;
    var summary = new UiTreeSummaryNode
    {
        Role = node.Role,
        Name = node.Name,
        AutomationId = node.AutomationId,
        Depth = depth
    };

    if (count >= maxNodes)
    {
        return summary;
    }

    foreach (var child in node.Children)
    {
        if (count >= maxNodes)
        {
            break;
        }

        summary.Children.Add(SummarizeNode(child, ref count, maxNodes, depth + 1));
    }

    return summary;
}

static bool TryParseTypeInLast(string message, out string text)
{
    text = string.Empty;
    if (string.IsNullOrWhiteSpace(message))
    {
        return false;
    }

    var trimmed = message.Trim();
    var lower = trimmed.ToLowerInvariant();
    var prefixes = new[]
    {
        "type in last ",
        "type into last ",
        "type in last textbox ",
        "type in last field ",
        "scrivi in ultimo ",
        "digita in ultimo "
    };

    foreach (var prefix in prefixes)
    {
        if (!lower.StartsWith(prefix))
        {
            continue;
        }

        text = trimmed[prefix.Length..].Trim();
        return !string.IsNullOrWhiteSpace(text);
    }

    return false;
}

static bool TryParseTranslationIntent(string message, out TranslationIntent intent)
{
    intent = default!;
    if (string.IsNullOrWhiteSpace(message))
    {
        return false;
    }

    var text = message.Trim();
    var lower = text.ToLowerInvariant();
    if (!lower.Contains("traduc", StringComparison.Ordinal)
        && !lower.Contains("tradur", StringComparison.Ordinal)
        && !lower.Contains("translat", StringComparison.Ordinal))
    {
        return false;
    }

    if (TryParseHeadAndBody(text, out var headTarget, out var headSource, out var headText))
    {
        intent = new TranslationIntent(headText, headTarget, headSource);
        return true;
    }

    var inMarker = lower.LastIndexOf(" in ", StringComparison.Ordinal);
    var toMarker = lower.LastIndexOf(" to ", StringComparison.Ordinal);
    var marker = Math.Max(inMarker, toMarker);
    if (marker <= 0)
    {
        return false;
    }

    var before = text[..marker].Trim();
    var after = text[(marker + 4)..].Trim();
    if (!string.IsNullOrWhiteSpace(after))
    {
        var sep = after.IndexOfAny(new[] { '?', ':', '\n', ';', '!' });
        var descriptor = sep >= 0 ? after[..sep].Trim() : after;
        var bodyAfter = sep >= 0 ? after[(sep + 1)..].Trim() : string.Empty;

        if (TryParseLanguageDescriptor(descriptor, out var target, out var source))
        {
            var body = !string.IsNullOrWhiteSpace(bodyAfter)
                ? bodyAfter
                : NormalizeTranslationLeadIn(before);
            if (string.IsNullOrWhiteSpace(body))
            {
                return false;
            }

            intent = new TranslationIntent(body, target, source);
            return true;
        }
    }

    if (TryParseImplicitTranslation(text, out intent))
    {
        return true;
    }

    return false;
}

static string NormalizeTranslationLeadIn(string value)
{
    var text = value.Trim();
    text = Regex.Replace(text, "^(can\\s+you\\s+)?(please\\s+)?translate\\s+", string.Empty, RegexOptions.IgnoreCase);
    text = Regex.Replace(text, "^(puoi\\s+)?(per\\s+favore\\s+)?tradur(?:re|mi|re\\s+mi)?\\s+", string.Empty, RegexOptions.IgnoreCase);
    text = Regex.Replace(text, "^(pui\\s+)?(per\\s+favore\\s+)?tradur(?:re|mi|re\\s+mi)?\\s+", string.Empty, RegexOptions.IgnoreCase);
    text = Regex.Replace(text, "^(traduci|tradurre)\\s+", string.Empty, RegexOptions.IgnoreCase);
    return text.Trim().Trim('?', ':', '.', '!', ',', ';');
}

static bool TryParseImplicitTranslation(string text, out TranslationIntent intent)
{
    intent = default!;
    if (string.IsNullOrWhiteSpace(text))
    {
        return false;
    }

    var separatorIndex = text.IndexOfAny(new[] { '?', ':', '\n' });
    if (separatorIndex <= 0 || separatorIndex >= text.Length - 1)
    {
        return false;
    }

    var leadIn = text[..separatorIndex].Trim();
    var body = text[(separatorIndex + 1)..].Trim();
    if (string.IsNullOrWhiteSpace(body))
    {
        return false;
    }

    intent = new TranslationIntent(body, InferDefaultTargetLanguage(leadIn, body), null);
    return true;
}

static string InferDefaultTargetLanguage(string leadIn, string body)
{
    var lead = (leadIn ?? string.Empty).ToLowerInvariant();
    var text = (body ?? string.Empty).ToLowerInvariant();

    if (lead.Contains("italian", StringComparison.Ordinal) || lead.Contains("italiano", StringComparison.Ordinal))
    {
        return "italian";
    }

    if (lead.Contains("english", StringComparison.Ordinal) || lead.Contains("inglese", StringComparison.Ordinal))
    {
        return "english";
    }

    var italianHints = new[]
    {
        "ciao", "sono", "come va", "grazie", "per favore", "buongiorno", "arrivederci", "oggi", "domani", "prego"
    };
    if (italianHints.Any(hint => text.Contains(hint, StringComparison.Ordinal)))
    {
        return "english";
    }

    var englishHints = new[]
    {
        "hello", "how are you", "thanks", "please", "good morning", "today", "tomorrow"
    };
    if (englishHints.Any(hint => text.Contains(hint, StringComparison.Ordinal)))
    {
        return "italian";
    }

    return "english";
}

static bool TryParseHeadAndBody(string remainder, out string targetLanguage, out string? sourceLanguage, out string text)
{
    targetLanguage = string.Empty;
    sourceLanguage = null;
    text = string.Empty;

    var newlineIdx = remainder.IndexOf('\n');
    if (newlineIdx > 0)
    {
        var head = remainder[..newlineIdx].Trim().TrimEnd(':');
        var body = remainder[(newlineIdx + 1)..].Trim();
        if (TryParseLanguageHead(head, out targetLanguage, out sourceLanguage) && !string.IsNullOrWhiteSpace(body))
        {
            text = body;
            return true;
        }
    }

    var colonIdx = remainder.IndexOf(':');
    if (colonIdx > 0)
    {
        var head = remainder[..colonIdx].Trim();
        var body = remainder[(colonIdx + 1)..].Trim();
        if (TryParseLanguageHead(head, out targetLanguage, out sourceLanguage) && !string.IsNullOrWhiteSpace(body))
        {
            text = body;
            return true;
        }
    }

    return false;
}

static bool TryParseLanguageHead(string head, out string targetLanguage, out string? sourceLanguage)
{
    targetLanguage = string.Empty;
    sourceLanguage = null;
    if (string.IsNullOrWhiteSpace(head))
    {
        return false;
    }

    var trimmed = head.Trim();
    if (trimmed.StartsWith("translate to ", StringComparison.OrdinalIgnoreCase))
    {
        return TryParseLanguageDescriptor(trimmed["translate to ".Length..].Trim(), out targetLanguage, out sourceLanguage);
    }
    if (trimmed.StartsWith("translate in ", StringComparison.OrdinalIgnoreCase))
    {
        return TryParseLanguageDescriptor(trimmed["translate in ".Length..].Trim(), out targetLanguage, out sourceLanguage);
    }
    if (trimmed.StartsWith("traduci in ", StringComparison.OrdinalIgnoreCase))
    {
        return TryParseLanguageDescriptor(trimmed["traduci in ".Length..].Trim(), out targetLanguage, out sourceLanguage);
    }
    if (trimmed.StartsWith("tradurre in ", StringComparison.OrdinalIgnoreCase))
    {
        return TryParseLanguageDescriptor(trimmed["tradurre in ".Length..].Trim(), out targetLanguage, out sourceLanguage);
    }
    if (trimmed.StartsWith("to ", StringComparison.OrdinalIgnoreCase))
    {
        return TryParseLanguageDescriptor(trimmed["to ".Length..].Trim(), out targetLanguage, out sourceLanguage);
    }
    if (trimmed.StartsWith("in ", StringComparison.OrdinalIgnoreCase))
    {
        return TryParseLanguageDescriptor(trimmed["in ".Length..].Trim(), out targetLanguage, out sourceLanguage);
    }

    return false;
}

static bool TryParseLanguageDescriptor(string descriptor, out string targetLanguage, out string? sourceLanguage)
{
    targetLanguage = string.Empty;
    sourceLanguage = null;
    if (string.IsNullOrWhiteSpace(descriptor))
    {
        return false;
    }

    var cleaned = descriptor.Trim().TrimEnd(':').Trim();
    var lower = cleaned.ToLowerInvariant();
    var sourceSep = lower.IndexOf(" from ", StringComparison.Ordinal);
    var sourceSepLen = 6;
    if (sourceSep <= 0)
    {
        sourceSep = lower.IndexOf(" da ", StringComparison.Ordinal);
        sourceSepLen = 4;
    }

    if (sourceSep > 0)
    {
        targetLanguage = cleaned[..sourceSep].Trim().Trim(',', ';', '.');
        sourceLanguage = cleaned[(sourceSep + sourceSepLen)..].Trim().Trim(',', ';', '.');
    }
    else
    {
        targetLanguage = cleaned.Trim().Trim(',', ';', '.');
    }

    if (string.IsNullOrWhiteSpace(targetLanguage))
    {
        return false;
    }

    if (string.IsNullOrWhiteSpace(sourceLanguage))
    {
        sourceLanguage = null;
    }

    return true;
}

static bool IsDirectIntent(string normalized)
{
    if (normalized.StartsWith("http://") || normalized.StartsWith("https://"))
    {
        return true;
    }

    return normalized.StartsWith("open ")
        || normalized.StartsWith("start ")
        || normalized.StartsWith("launch ")
        || normalized.StartsWith("apri ")
        || normalized.StartsWith("avvia ")
        || normalized.StartsWith("esegui ")
        || normalized.StartsWith("lancia ")
        || normalized.StartsWith("in ")
        || normalized.StartsWith("nel ")
        || normalized.StartsWith("nella ")
        || normalized.StartsWith("find ")
        || normalized.StartsWith("cerca ")
        || normalized.StartsWith("trova ")
        || normalized.StartsWith("click ")
        || normalized.StartsWith("double click ")
        || normalized.StartsWith("doppio clic ")
        || normalized.StartsWith("right click ")
        || normalized.StartsWith("clic destro ")
        || normalized.StartsWith("drag ")
        || normalized.StartsWith("trascina ")
        || normalized.StartsWith("clicca ")
        || normalized.StartsWith("clic ")
        || normalized.StartsWith("type ")
        || normalized.StartsWith("scrivi ")
        || normalized.StartsWith("digita ")
        || normalized.StartsWith("press ")
        || normalized.StartsWith("premi ")
        || normalized.StartsWith("create")
        || normalized.StartsWith("crea")
        || normalized.StartsWith("new file")
        || normalized.StartsWith("save")
        || normalized.StartsWith("salva")
        || normalized.StartsWith("new tab")
        || normalized.StartsWith("nuova")
        || normalized.StartsWith("close tab")
        || normalized.Equals("back")
        || normalized.StartsWith("go back")
        || normalized.StartsWith("indietro")
        || normalized.Equals("forward")
        || normalized.StartsWith("go forward")
        || normalized.StartsWith("avanti")
        || normalized.StartsWith("refresh")
        || normalized.StartsWith("reload")
        || normalized.StartsWith("ricarica")
        || normalized.StartsWith("find in page ")
        || normalized.StartsWith("search in page ")
        || normalized.StartsWith("cerca nella pagina ")
        || normalized.StartsWith("trova nella pagina ")
        || normalized.StartsWith("go to ")
        || normalized.StartsWith("navigate to ")
        || normalized.StartsWith("vai a ")
        || normalized.StartsWith("naviga a ")
        || normalized.StartsWith("wait until ")
        || normalized.StartsWith("aspetta finche ")
        || normalized.StartsWith("aspetta finché ")
        || normalized.StartsWith("chiudi")
        || normalized.StartsWith("minimize")
        || normalized.StartsWith("maximize")
        || normalized.StartsWith("restore")
        || normalized.StartsWith("switch window")
        || normalized.StartsWith("focus ")
        || normalized.StartsWith("scroll ")
        || normalized.StartsWith("scorri ")
        || normalized.StartsWith("page up")
        || normalized.StartsWith("page down")
        || normalized.Equals("home")
        || normalized.Equals("end")
        || normalized.StartsWith("copy")
        || normalized.StartsWith("copia")
        || normalized.StartsWith("paste")
        || normalized.StartsWith("incolla")
        || normalized.StartsWith("undo")
        || normalized.StartsWith("annulla")
        || normalized.StartsWith("redo")
        || normalized.StartsWith("ripeti")
        || normalized.StartsWith("select all")
        || normalized.StartsWith("seleziona")
        || normalized.StartsWith("file ")
        || normalized.StartsWith("notify ")
        || normalized.StartsWith("notification ")
        || normalized.StartsWith("notifica ")
        || normalized.StartsWith("clipboard history")
        || normalized.StartsWith("show clipboard history")
        || normalized.StartsWith("cronologia clipboard")
        || normalized.StartsWith("jiggle mouse")
        || normalized.StartsWith("mouse jiggle")
        || normalized.StartsWith("move mouse")
        || normalized.StartsWith("move the mouse")
        || normalized.StartsWith("muovi mouse")
        || normalized.StartsWith("muovi il mouse")
        || normalized.StartsWith("muovi randomicamente")
        || normalized.StartsWith("volume ")
        || normalized.StartsWith("audio ")
        || normalized.StartsWith("brightness ")
        || normalized.StartsWith("luminosita ")
        || normalized.StartsWith("luminosità ")
        || normalized.StartsWith("lock screen")
        || normalized.StartsWith("lock workstation")
        || normalized.StartsWith("blocca schermo")
        || normalized.StartsWith("key ")
        || normalized.StartsWith("keys ");
}

static bool IsUnrecognizedPlan(ActionPlan plan)
{
    if (plan.Steps.Count == 0)
    {
        return true;
    }

    if (plan.Steps.Count == 1 && plan.Steps[0].Type == ActionType.ReadText)
    {
        var note = plan.Steps[0].Note ?? string.Empty;
        if (string.IsNullOrWhiteSpace(note))
        {
            return true;
        }

        if (note.StartsWith("Default to read text", StringComparison.OrdinalIgnoreCase)
            || note.StartsWith("Unrecognized", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }
    }

    return false;
}

internal static class LlmTranslationService
{
    private static readonly JsonSerializerOptions JsonOptions = new(JsonSerializerDefaults.Web);

    public static async Task<TranslationResult> TranslateAsync(TranslationIntent intent, AgentConfig config, IAuditLog auditLog, CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(intent.Text) || string.IsNullOrWhiteSpace(intent.TargetLanguage))
        {
            return TranslationResult.Fail("Missing text or target language.");
        }

        if (!config.LlmFallbackEnabled)
        {
            return TranslationResult.Fail("LLM is disabled. Enable it in Configuration to use translation.");
        }

        var endpoint = config.LlmFallback.Endpoint?.Trim();
        if (string.IsNullOrWhiteSpace(endpoint))
        {
            return TranslationResult.Fail("LLM endpoint is not configured.");
        }

        if (!Uri.TryCreate(endpoint, UriKind.Absolute, out var uri)
            || (!uri.IsLoopback && !config.AllowNonLoopbackLlmEndpoint))
        {
            return TranslationResult.Fail("LLM endpoint is invalid or blocked by policy.");
        }

        var provider = (config.LlmFallback.Provider ?? "ollama").Trim().ToLowerInvariant();
        uri = NormalizeLlmEndpoint(uri, provider);
        var prompt = BuildPrompt(intent);
        var maxTokens = ResolveTranslationMaxTokens(intent.Text, config.LlmFallback.MaxTokens);
        var includeRaw = config.AuditLlmIncludeRawText;
        var canAudit = config.AuditLlmInteractions;
        var sw = Stopwatch.StartNew();

        if (canAudit)
        {
            await auditLog.WriteAsync(new AuditEvent
            {
                Timestamp = DateTimeOffset.UtcNow,
                EventType = "llm_translate_request",
                Message = "Translation requested",
                Data = new
                {
                    provider,
                    model = config.LlmFallback.Model,
                    endpoint,
                    target = intent.TargetLanguage,
                    source = intent.SourceLanguage,
                    input = includeRaw ? intent.Text : "[redacted]",
                    inputLength = intent.Text.Length
                }
            }, CancellationToken.None);
        }

        try
        {
            using var client = new HttpClient
            {
                Timeout = TimeSpan.FromSeconds(Math.Max(2, config.LlmFallback.TimeoutSeconds))
            };

            string? translated = provider switch
            {
                "openai" => await CallOpenAiCompatibleAsync(client, uri, prompt, config, maxTokens, cancellationToken),
                "llama.cpp" => await CallLlamaCppAsync(client, uri, prompt, config, maxTokens, cancellationToken),
                _ => await CallOllamaAsync(client, uri, prompt, config, maxTokens, cancellationToken)
            };

            translated = CleanTranslationOutput(translated);
            sw.Stop();

            if (string.IsNullOrWhiteSpace(translated))
            {
                return TranslationResult.Fail("Translation failed: model returned an empty response.");
            }

            if (canAudit)
            {
                await auditLog.WriteAsync(new AuditEvent
                {
                    Timestamp = DateTimeOffset.UtcNow,
                    EventType = "llm_translate_response",
                    Message = "Translation completed",
                    Data = new
                    {
                        provider,
                        model = config.LlmFallback.Model,
                        latencyMs = sw.ElapsedMilliseconds,
                        output = includeRaw ? translated : "[redacted]",
                        outputLength = translated.Length
                    }
                }, CancellationToken.None);
            }

            return TranslationResult.Ok(translated, provider);
        }
        catch (Exception ex)
        {
            sw.Stop();
            if (canAudit)
            {
                await auditLog.WriteAsync(new AuditEvent
                {
                    Timestamp = DateTimeOffset.UtcNow,
                    EventType = "llm_translate_error",
                    Message = "Translation failed",
                    Data = new
                    {
                        provider,
                        model = config.LlmFallback.Model,
                        latencyMs = sw.ElapsedMilliseconds,
                        error = ex.Message
                    }
                }, CancellationToken.None);
            }

            return TranslationResult.Fail($"Translation failed: {ex.Message}");
        }
    }

    private static async Task<string?> CallOllamaAsync(HttpClient client, Uri uri, string prompt, AgentConfig config, int maxTokens, CancellationToken cancellationToken)
    {
        var payload = new
        {
            model = config.LlmFallback.Model,
            prompt,
            stream = false,
            options = new { temperature = 0.1, num_predict = maxTokens }
        };

        using var response = await client.PostAsJsonAsync(uri, payload, JsonOptions, cancellationToken);
        if (!response.IsSuccessStatusCode)
        {
            throw await BuildLlmHttpExceptionAsync(response, cancellationToken);
        }

        await using var stream = await response.Content.ReadAsStreamAsync(cancellationToken);
        using var doc = await JsonDocument.ParseAsync(stream, cancellationToken: cancellationToken);
        if (doc.RootElement.TryGetProperty("response", out var resp))
        {
            return resp.GetString();
        }

        return null;
    }

    private static async Task<string?> CallOpenAiCompatibleAsync(HttpClient client, Uri uri, string prompt, AgentConfig config, int maxTokens, CancellationToken cancellationToken)
    {
        var payload = new
        {
            model = config.LlmFallback.Model,
            messages = new[]
            {
                new { role = "system", content = "You are a precise translator. Return only the translation." },
                new { role = "user", content = prompt }
            },
            temperature = 0.1,
            max_tokens = maxTokens
        };

        using var response = await client.PostAsJsonAsync(uri, payload, JsonOptions, cancellationToken);
        if (!response.IsSuccessStatusCode)
        {
            throw await BuildLlmHttpExceptionAsync(response, cancellationToken);
        }

        await using var stream = await response.Content.ReadAsStreamAsync(cancellationToken);
        using var doc = await JsonDocument.ParseAsync(stream, cancellationToken: cancellationToken);
        if (doc.RootElement.TryGetProperty("choices", out var choices) && choices.ValueKind == JsonValueKind.Array && choices.GetArrayLength() > 0)
        {
            var first = choices[0];
            if (first.TryGetProperty("message", out var message) && message.TryGetProperty("content", out var content))
            {
                return content.GetString();
            }
            if (first.TryGetProperty("text", out var text))
            {
                return text.GetString();
            }
        }

        return null;
    }

    private static async Task<string?> CallLlamaCppAsync(HttpClient client, Uri uri, string prompt, AgentConfig config, int maxTokens, CancellationToken cancellationToken)
    {
        var payload = new
        {
            prompt,
            n_predict = maxTokens,
            temperature = 0.1
        };

        using var response = await client.PostAsJsonAsync(uri, payload, JsonOptions, cancellationToken);
        if (!response.IsSuccessStatusCode)
        {
            throw await BuildLlmHttpExceptionAsync(response, cancellationToken);
        }

        await using var stream = await response.Content.ReadAsStreamAsync(cancellationToken);
        using var doc = await JsonDocument.ParseAsync(stream, cancellationToken: cancellationToken);
        if (doc.RootElement.TryGetProperty("content", out var content))
        {
            return content.GetString();
        }
        if (doc.RootElement.TryGetProperty("choices", out var choices) && choices.ValueKind == JsonValueKind.Array && choices.GetArrayLength() > 0)
        {
            var first = choices[0];
            if (first.TryGetProperty("text", out var text))
            {
                return text.GetString();
            }
        }

        return null;
    }

    private static Uri NormalizeLlmEndpoint(Uri uri, string provider)
    {
        if (!string.IsNullOrWhiteSpace(uri.AbsolutePath) && uri.AbsolutePath != "/")
        {
            return uri;
        }

        return provider switch
        {
            "openai" => new Uri(uri, "/v1/chat/completions"),
            "llama.cpp" => new Uri(uri, "/completion"),
            _ => new Uri(uri, "/api/generate")
        };
    }

    private static int ResolveTranslationMaxTokens(string input, int configuredMaxTokens)
    {
        var configured = Math.Clamp(configuredMaxTokens, 32, 4096);
        var estimated = Math.Clamp((int)Math.Ceiling((input?.Length ?? 0) / 2.8), 128, 4096);
        return Math.Max(configured, estimated);
    }

    private static async Task<Exception> BuildLlmHttpExceptionAsync(HttpResponseMessage response, CancellationToken cancellationToken)
    {
        var detail = await TryReadLlmErrorAsync(response, cancellationToken);
        return new InvalidOperationException($"LLM HTTP {(int)response.StatusCode} ({response.ReasonPhrase ?? "HTTP error"}): {detail}");
    }

    private static async Task<string> TryReadLlmErrorAsync(HttpResponseMessage response, CancellationToken cancellationToken)
    {
        try
        {
            var body = await response.Content.ReadAsStringAsync(cancellationToken);
            if (string.IsNullOrWhiteSpace(body))
            {
                return "empty response body";
            }

            var trimmed = body.Trim();
            try
            {
                using var doc = JsonDocument.Parse(trimmed);
                if (doc.RootElement.ValueKind == JsonValueKind.Object)
                {
                    if (doc.RootElement.TryGetProperty("error", out var errorProp))
                    {
                        return Compact(errorProp.ToString(), 220);
                    }
                    if (doc.RootElement.TryGetProperty("message", out var messageProp))
                    {
                        return Compact(messageProp.ToString(), 220);
                    }
                }
            }
            catch
            {
                // Best effort.
            }

            return Compact(trimmed, 220);
        }
        catch
        {
            return "unable to read response body";
        }
    }

    private static string Compact(string? value, int maxChars)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        var normalized = value.Replace("\r", " ").Replace("\n", " ").Trim();
        return normalized.Length <= maxChars ? normalized : normalized[..maxChars] + "...";
    }

    private static string BuildPrompt(TranslationIntent intent)
    {
        var sourcePart = string.IsNullOrWhiteSpace(intent.SourceLanguage)
            ? string.Empty
            : $"Source language: {intent.SourceLanguage}\n";

        return
            "You are a translation engine.\n" +
            "Task: translate the text exactly, preserving meaning, tone, formatting and line breaks.\n" +
            "Do not explain. Do not add notes. Return only translated text.\n" +
            $"Target language: {intent.TargetLanguage}\n" +
            sourcePart +
            "TEXT START\n" +
            intent.Text +
            "\nTEXT END";
    }

    private static string? CleanTranslationOutput(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw))
        {
            return null;
        }

        var value = raw.Trim();
        if (value.StartsWith("```", StringComparison.Ordinal))
        {
            value = value.Trim('`').Trim();
        }

        if (value.StartsWith("translation:", StringComparison.OrdinalIgnoreCase))
        {
            value = value["translation:".Length..].Trim();
        }

        return string.IsNullOrWhiteSpace(value) ? null : value;
    }
}

internal sealed record UtilityInstallRequest(string Tool, bool EnableOcr);
internal sealed record UtilityEnableOcrRequest(string? TesseractPath);
internal sealed record TranslateRequest(string Text, string TargetLanguage, string? SourceLanguage);
internal sealed record TranslationIntent(string Text, string TargetLanguage, string? SourceLanguage);
internal sealed record TranslationResult(bool Success, string Message, string? TranslatedText, string Provider)
{
    public static TranslationResult Ok(string text, string provider) => new(true, "ok", text, provider);
    public static TranslationResult Fail(string message) => new(false, message, null, "none");
}

internal static class RequestGuards
{
    public static bool IsLoopback(HttpContext context)
    {
        var remoteIp = context.Connection.RemoteIpAddress;
        return remoteIp == null || System.Net.IPAddress.IsLoopback(remoteIp);
    }
}

internal static class UtilityInstaller
{
    private static readonly string[] WindowsExecutableExtensions = (Environment.GetEnvironmentVariable("PATHEXT") ?? ".EXE;.CMD;.BAT")
        .Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

    public static string GetOsName()
    {
        if (OperatingSystem.IsWindows())
        {
            return "windows";
        }

        if (OperatingSystem.IsMacOS())
        {
            return "macos";
        }

        if (OperatingSystem.IsLinux())
        {
            return "linux";
        }

        return "unknown";
    }

    public static string GetPreferredPackageManager()
    {
        if (OperatingSystem.IsWindows())
        {
            return TryResolveExecutable("winget") != null ? "winget" : "manual";
        }

        if (OperatingSystem.IsMacOS())
        {
            return TryResolveExecutable("brew") != null ? "brew" : "manual";
        }

        if (OperatingSystem.IsLinux())
        {
            if (TryResolveExecutable("apt-get") != null) return "apt";
            if (TryResolveExecutable("dnf") != null) return "dnf";
            if (TryResolveExecutable("pacman") != null) return "pacman";
            return "manual";
        }

        return "manual";
    }

    public static ToolProbeResult Probe(string commandOrPath, string versionArgs)
    {
        var resolvedPath = TryResolveExecutable(commandOrPath);
        if (resolvedPath == null)
        {
            return new ToolProbeResult(false, null, null, "not found");
        }

        var run = RunCommandAsync(resolvedPath, versionArgs, TimeSpan.FromSeconds(6), CancellationToken.None)
            .ConfigureAwait(false)
            .GetAwaiter()
            .GetResult();
        var versionLine = FirstNonEmptyLine(run.StdOut) ?? FirstNonEmptyLine(run.StdErr);
        return new ToolProbeResult(run.ExitCode == 0, resolvedPath, versionLine, run.ExitCode == 0 ? "ok" : $"exit {run.ExitCode}");
    }

    public static async Task<UtilityInstallResult> InstallAsync(
        string tool,
        bool enableOcr,
        AgentConfig config,
        ConfigFileStore store,
        RestartRequirementTracker restartTracker,
        CancellationToken cancellationToken)
    {
        if (OperatingSystem.IsWindows())
        {
            if (TryResolveExecutable("winget") == null)
            {
                return UtilityInstallResult.Fail("winget not found. Install App Installer from Microsoft Store, then retry.");
            }

            var commands = BuildWindowsInstallCommands(tool);
            return await RunInstallCommandsAsync(commands, tool, enableOcr, config, store, restartTracker, cancellationToken);
        }

        if (OperatingSystem.IsMacOS())
        {
            if (TryResolveExecutable("brew") == null)
            {
                return UtilityInstallResult.Fail("Homebrew not found. Install Homebrew, then run: brew install ffmpeg tesseract");
            }

            var packageName = tool == "ffmpeg" ? "ffmpeg" : "tesseract";
            var installResult = await RunCommandAsync("brew", $"install {packageName}", TimeSpan.FromMinutes(15), cancellationToken);
            if (installResult.ExitCode != 0)
            {
                return UtilityInstallResult.Fail($"brew install {packageName} failed.", installResult.StdOut, installResult.StdErr);
            }

            if (tool != "ffmpeg" && enableOcr)
            {
                var changed = !config.OcrEnabled || !string.Equals(config.Ocr.TesseractPath, "tesseract", StringComparison.Ordinal);
                config.OcrEnabled = true;
                config.Ocr.TesseractPath = "tesseract";
                if (changed)
                {
                    restartTracker.OcrRestartRequired = true;
                    await store.SaveAsync(config);
                }
            }

            return UtilityInstallResult.Ok($"Installed {packageName}.", installResult.StdOut, installResult.StdErr);
        }

        if (OperatingSystem.IsLinux())
        {
            var hint = tool == "ffmpeg"
                ? "Run manually: sudo apt-get update && sudo apt-get install -y ffmpeg"
                : "Run manually: sudo apt-get update && sudo apt-get install -y tesseract-ocr";
            return UtilityInstallResult.Fail($"Automatic install is not enabled on Linux web endpoint. {hint}");
        }

        return UtilityInstallResult.Fail("Unsupported OS for automatic install.");
    }

    private static IReadOnlyList<(string FileName, string Arguments)> BuildWindowsInstallCommands(string tool)
    {
        if (tool == "ffmpeg")
        {
            return
            [
                ("winget", "install -e --id Gyan.FFmpeg --accept-package-agreements --accept-source-agreements"),
                ("winget", "install -e --id BtbN.FFmpeg --accept-package-agreements --accept-source-agreements")
            ];
        }

        return
        [
            ("winget", "install -e --id UB-Mannheim.TesseractOCR --accept-package-agreements --accept-source-agreements"),
            ("winget", "install -e --id Tesseract-OCR.Tesseract --accept-package-agreements --accept-source-agreements")
        ];
    }

    private static async Task<UtilityInstallResult> RunInstallCommandsAsync(
        IReadOnlyList<(string FileName, string Arguments)> commands,
        string tool,
        bool enableOcr,
        AgentConfig config,
        ConfigFileStore store,
        RestartRequirementTracker restartTracker,
        CancellationToken cancellationToken)
    {
        UtilityCommandResult? last = null;
        foreach (var command in commands)
        {
            last = await RunCommandAsync(command.FileName, command.Arguments, TimeSpan.FromMinutes(15), cancellationToken);
            if (last.ExitCode == 0)
            {
                break;
            }
        }

        if (last == null || last.ExitCode != 0)
        {
            return UtilityInstallResult.Fail($"Failed to install {tool}.", last?.StdOut, last?.StdErr);
        }

        if (tool != "ffmpeg" && enableOcr)
        {
            var probe = Probe("tesseract", "--version");
            var path = probe.Path ?? "tesseract";
            var changed = !config.OcrEnabled || !string.Equals(config.Ocr.TesseractPath, path, StringComparison.Ordinal);
            config.OcrEnabled = true;
            config.Ocr.TesseractPath = path;
            if (changed)
            {
                restartTracker.OcrRestartRequired = true;
                await store.SaveAsync(config);
            }
        }

        var label = tool == "ffmpeg" ? "FFmpeg" : "Tesseract OCR";
        return UtilityInstallResult.Ok($"Installed {label}.", last.StdOut, last.StdErr);
    }

    private static string? TryResolveExecutable(string commandOrPath)
    {
        if (string.IsNullOrWhiteSpace(commandOrPath))
        {
            return null;
        }

        var candidate = commandOrPath.Trim().Trim('"');
        if (candidate.Length == 0)
        {
            return null;
        }

        if (Path.IsPathRooted(candidate) || candidate.Contains(Path.DirectorySeparatorChar) || candidate.Contains(Path.AltDirectorySeparatorChar))
        {
            return ResolveCandidatePath(candidate);
        }

        var pathEntries = (Environment.GetEnvironmentVariable("PATH") ?? string.Empty)
            .Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        foreach (var entry in pathEntries)
        {
            var combined = Path.Combine(entry, candidate);
            var resolved = ResolveCandidatePath(combined);
            if (resolved != null)
            {
                return resolved;
            }
        }

        return null;
    }

    private static string? ResolveCandidatePath(string candidate)
    {
        if (File.Exists(candidate))
        {
            return Path.GetFullPath(candidate);
        }

        if (!OperatingSystem.IsWindows() || Path.GetExtension(candidate).Length > 0)
        {
            return null;
        }

        foreach (var ext in WindowsExecutableExtensions)
        {
            var withExt = candidate + ext;
            if (File.Exists(withExt))
            {
                return Path.GetFullPath(withExt);
            }
        }

        return null;
    }

    private static async Task<UtilityCommandResult> RunCommandAsync(string fileName, string arguments, TimeSpan timeout, CancellationToken cancellationToken)
    {
        using var process = new Process
        {
            StartInfo = new ProcessStartInfo
            {
                FileName = fileName,
                Arguments = arguments,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            }
        };

        try
        {
            if (!process.Start())
            {
                return new UtilityCommandResult(-1, string.Empty, "Process failed to start.");
            }

            var outputTask = process.StandardOutput.ReadToEndAsync();
            var errorTask = process.StandardError.ReadToEndAsync();

            using var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            timeoutCts.CancelAfter(timeout);
            await process.WaitForExitAsync(timeoutCts.Token);

            var output = await outputTask;
            var error = await errorTask;
            return new UtilityCommandResult(process.ExitCode, output, error);
        }
        catch (OperationCanceledException)
        {
            try
            {
                if (!process.HasExited)
                {
                    process.Kill(entireProcessTree: true);
                }
            }
            catch
            {
                // ignore
            }

            return new UtilityCommandResult(-2, string.Empty, $"Command timed out after {timeout.TotalSeconds:F0}s.");
        }
        catch (Exception ex)
        {
            return new UtilityCommandResult(-1, string.Empty, ex.Message);
        }
    }

    private static string? FirstNonEmptyLine(string? text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return null;
        }

        return text.Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .FirstOrDefault(line => !string.IsNullOrWhiteSpace(line));
    }
}

internal sealed record ToolProbeResult(bool Installed, string? Path, string? Version, string Status);
internal sealed record UtilityCommandResult(int ExitCode, string StdOut, string StdErr);

internal sealed record UtilityInstallResult(bool Success, string Message, string? StdOut, string? StdErr)
{
    public static UtilityInstallResult Ok(string message, string? stdOut, string? stdErr)
        => new(true, message, stdOut, stdErr);

    public static UtilityInstallResult Fail(string message, string? stdOut = null, string? stdErr = null)
        => new(false, message, stdOut, stdErr);
}

internal sealed record ChatRequest(string Message);
internal sealed record TaskUpsertRequest(string Name, string Intent, string? Description, string? PlanJson);
internal sealed record RunTaskRequest(bool DryRun);
internal sealed record ScheduleUpsertRequest(string? Id, string TaskName, DateTimeOffset? StartAtUtc, int? IntervalSeconds, bool? Enabled);
internal sealed record MacroRecordStartRequest(string? Name);
internal sealed record MacroRecordSaveRequest(string Name, string? Description);

internal sealed record ConfirmRequest(string Token, bool Approve, string? PlanJson);

internal sealed record IntentRequest(string Intent, bool DryRun);

internal sealed record ChatResponse(string Reply, bool NeedsConfirmation, string? Token, string? ActionLabel, IReadOnlyList<string>? Steps, string? PlanJson, string? ModeLabel)
{
    public static ChatResponse Simple(string reply) => new(reply, false, null, null, null, null, null);
    public static ChatResponse WithSteps(string reply, IReadOnlyList<string> steps, string? planJson, string? modeLabel) => new(reply, false, null, null, steps, planJson, modeLabel);
    public static ChatResponse Confirm(string reply, string token) => new(reply, true, token, "Confirm", null, null, null);
    public static ChatResponse ConfirmWithSteps(string reply, string token, IReadOnlyList<string> steps, string? planJson, string? modeLabel) => new(reply, true, token, "Confirm", steps, planJson, modeLabel);
}

internal enum ChatActionType
{
    SimulatePresence,
    ExecutePlan
}

internal sealed class ChatActionStore
{
    private readonly Dictionary<string, PendingAction> _actions = new();
    private readonly object _lock = new();

    public string CreatePending(ChatActionType type, string message, ActionPlan? plan, bool dryRun)
    {
        var token = Guid.NewGuid().ToString("n");
        lock (_lock)
        {
            _actions[token] = new PendingAction(type, DateTimeOffset.UtcNow, message, plan, dryRun);
        }
        return token;
    }

    public PendingAction? Take(string token)
    {
        lock (_lock)
        {
            if (_actions.TryGetValue(token, out var action))
            {
                _actions.Remove(token);
                return action;
            }
        }
        return null;
    }
}

internal sealed record PendingAction(ChatActionType Type, DateTimeOffset CreatedAt, string Message, ActionPlan? Plan, bool DryRun);

internal sealed class TargetMemoryStore
{
    private readonly object _lock = new();
    private MemoryWindow? _lastWindow;
    private MemoryElement? _lastElement;

    public void Capture(ActionPlan plan, ExecutionResult result, WindowRef? activeWindow)
    {
        lock (_lock)
        {
            if (activeWindow != null)
            {
                _lastWindow = new MemoryWindow(activeWindow.Id, activeWindow.Title, activeWindow.AppId);
            }

            foreach (var executed in result.Steps.Where(s => s.Success))
            {
                if (executed.Index < 0 || executed.Index >= plan.Steps.Count)
                {
                    continue;
                }

                var planned = plan.Steps[executed.Index];
                TryCaptureElementFromData(executed.Data);

                if (_lastElement == null && planned.Selector != null && !string.IsNullOrWhiteSpace(planned.Selector.NameContains))
                {
                    _lastElement = new MemoryElement(
                        Id: planned.ElementId ?? string.Empty,
                        Role: planned.Selector.Role ?? string.Empty,
                        Name: planned.Selector.NameContains ?? string.Empty,
                        AutomationId: planned.Selector.AutomationId ?? string.Empty,
                        ClassName: planned.Selector.ClassName ?? string.Empty,
                        PathHints: planned.Selector.AncestorNameContains ?? string.Empty,
                        Bounds: planned.Selector.BoundsHint == null ? null : new Rect
                        {
                            X = planned.Selector.BoundsHint.X,
                            Y = planned.Selector.BoundsHint.Y,
                            Width = planned.Selector.BoundsHint.Width,
                            Height = planned.Selector.BoundsHint.Height
                        });
                }
            }
        }
    }

    public ActionPlan? BuildClickLastPlan()
    {
        lock (_lock)
        {
            if (_lastElement == null)
            {
                return null;
            }

            var step = new PlanStep
            {
                Type = ActionType.Click,
                Selector = _lastElement.ToSelector(),
                ExpectedAppId = _lastWindow?.AppId,
                ExpectedWindowId = _lastWindow?.Id
            };

            if (_lastElement.Bounds != null && _lastElement.Bounds.Width > 0 && _lastElement.Bounds.Height > 0)
            {
                step.Point = new Rect
                {
                    X = _lastElement.Bounds.X + (_lastElement.Bounds.Width / 2),
                    Y = _lastElement.Bounds.Y + (_lastElement.Bounds.Height / 2),
                    Width = 0,
                    Height = 0
                };
            }

            return new ActionPlan
            {
                Intent = "click last",
                Steps = new List<PlanStep> { step }
            };
        }
    }

    public ActionPlan? BuildTypeInLastPlan(string text)
    {
        lock (_lock)
        {
            if (_lastElement == null || string.IsNullOrWhiteSpace(text))
            {
                return null;
            }

            var steps = new List<PlanStep>();
            if (!string.IsNullOrWhiteSpace(_lastElement.Id))
            {
                steps.Add(new PlanStep
                {
                    Type = ActionType.SetValue,
                    ElementId = _lastElement.Id,
                    Text = text,
                    ExpectedAppId = _lastWindow?.AppId,
                    ExpectedWindowId = _lastWindow?.Id
                });
            }
            else
            {
                steps.Add(new PlanStep
                {
                    Type = ActionType.Click,
                    Selector = _lastElement.ToSelector(),
                    ExpectedAppId = _lastWindow?.AppId,
                    ExpectedWindowId = _lastWindow?.Id
                });
                steps.Add(new PlanStep
                {
                    Type = ActionType.TypeText,
                    Text = text,
                    ExpectedAppId = _lastWindow?.AppId,
                    ExpectedWindowId = _lastWindow?.Id
                });
            }

            return new ActionPlan
            {
                Intent = "type in last",
                Steps = steps
            };
        }
    }

    public TargetMemorySnapshot GetSnapshot()
    {
        lock (_lock)
        {
            return new TargetMemorySnapshot(
                _lastWindow?.Id,
                _lastWindow?.Title,
                _lastWindow?.AppId,
                _lastElement?.Id,
                _lastElement?.Role,
                _lastElement?.Name,
                _lastElement?.AutomationId,
                _lastElement?.ClassName,
                _lastElement?.PathHints);
        }
    }

    public void Clear()
    {
        lock (_lock)
        {
            _lastWindow = null;
            _lastElement = null;
        }
    }

    private void TryCaptureElementFromData(object? data)
    {
        if (data is ElementRef element)
        {
            _lastElement = MemoryElement.FromElementRef(element);
            return;
        }

        if (data is IReadOnlyList<ElementRef> elements && elements.Count > 0)
        {
            _lastElement = MemoryElement.FromElementRef(elements[0]);
            return;
        }

        if (data is IEnumerable<ElementRef> enumerable)
        {
            var first = enumerable.FirstOrDefault();
            if (first != null)
            {
                _lastElement = MemoryElement.FromElementRef(first);
            }
        }
    }

    private sealed record MemoryWindow(string Id, string Title, string AppId);

    private sealed record MemoryElement(
        string Id,
        string Role,
        string Name,
        string AutomationId,
        string ClassName,
        string PathHints,
        Rect? Bounds)
    {
        public static MemoryElement FromElementRef(ElementRef element)
        {
            return new MemoryElement(
                element.Id,
                element.Role,
                element.Name,
                element.AutomationId,
                element.ClassName,
                element.PathHints,
                element.Bounds == null
                    ? null
                    : new Rect
                    {
                        X = element.Bounds.X,
                        Y = element.Bounds.Y,
                        Width = element.Bounds.Width,
                        Height = element.Bounds.Height
                    });
        }

        public Selector ToSelector()
        {
            return new Selector
            {
                Role = Role,
                NameContains = Name,
                AutomationId = AutomationId,
                ClassName = ClassName,
                AncestorNameContains = PathHints,
                BoundsHint = Bounds == null ? null : new Rect
                {
                    X = Bounds.X,
                    Y = Bounds.Y,
                    Width = Bounds.Width,
                    Height = Bounds.Height
                }
            };
        }
    }
}

internal sealed record TargetMemorySnapshot(
    string? LastWindowId,
    string? LastWindowTitle,
    string? LastWindowAppId,
    string? LastElementId,
    string? LastElementRole,
    string? LastElementName,
    string? LastElementAutomationId,
    string? LastElementClassName,
    string? LastElementPathHints);

internal sealed record ContextLockSnapshot(
    bool Enabled,
    string? AppId,
    string? WindowId,
    string? WindowTitle,
    DateTimeOffset? LockedAt,
    string Scope)
{
    public static ContextLockSnapshot Unlocked { get; } = new(false, null, null, null, null, "none");
}

internal sealed class ContextLockStore
{
    private readonly object _lock = new();
    private ContextLockSnapshot _state = ContextLockSnapshot.Unlocked;

    public ContextLockSnapshot GetSnapshot()
    {
        lock (_lock)
        {
            return _state;
        }
    }

    public void Unlock()
    {
        lock (_lock)
        {
            _state = ContextLockSnapshot.Unlocked;
        }
    }

    public void LockToWindow(WindowRef window)
    {
        lock (_lock)
        {
            _state = new ContextLockSnapshot(
                Enabled: true,
                AppId: window.AppId,
                WindowId: window.Id,
                WindowTitle: window.Title,
                LockedAt: DateTimeOffset.UtcNow,
                Scope: "window");
        }
    }

    public void LockToWindowId(string windowId)
    {
        lock (_lock)
        {
            _state = new ContextLockSnapshot(
                Enabled: true,
                AppId: null,
                WindowId: windowId.Trim(),
                WindowTitle: null,
                LockedAt: DateTimeOffset.UtcNow,
                Scope: "window");
        }
    }

    public void LockToApp(string appId)
    {
        lock (_lock)
        {
            _state = new ContextLockSnapshot(
                Enabled: true,
                AppId: appId.Trim(),
                WindowId: null,
                WindowTitle: null,
                LockedAt: DateTimeOffset.UtcNow,
                Scope: "app");
        }
    }
}

internal sealed class MacroRecorderStore
{
    private readonly object _lock = new();
    private readonly List<PlanStep> _steps = new();
    private bool _isRecording;
    private string? _name;
    private DateTimeOffset? _startedAt;

    public void Start(string? name)
    {
        lock (_lock)
        {
            _steps.Clear();
            _isRecording = true;
            _name = string.IsNullOrWhiteSpace(name) ? null : name.Trim();
            _startedAt = DateTimeOffset.UtcNow;
        }
    }

    public void Stop()
    {
        lock (_lock)
        {
            _isRecording = false;
        }
    }

    public void Capture(string source, ActionPlan plan, ExecutionResult result)
    {
        lock (_lock)
        {
            if (!_isRecording || plan.Steps.Count == 0 || result.Steps.Count == 0)
            {
                return;
            }

            foreach (var executed in result.Steps.Where(step => step.Success))
            {
                if (executed.Index < 0 || executed.Index >= plan.Steps.Count)
                {
                    continue;
                }

                _steps.Add(CloneStep(plan.Steps[executed.Index]));
                if (_steps.Count > 2000)
                {
                    _steps.RemoveAt(0);
                }
            }
        }
    }

    public ActionPlan? BuildRecordedPlan()
    {
        lock (_lock)
        {
            if (_steps.Count == 0)
            {
                return null;
            }

            return new ActionPlan
            {
                Intent = string.IsNullOrWhiteSpace(_name) ? "recorded-macro" : $"recorded-macro:{_name}",
                Steps = _steps.Select(CloneStep).ToList()
            };
        }
    }

    public object GetStatus()
    {
        lock (_lock)
        {
            return new
            {
                isRecording = _isRecording,
                name = _name,
                startedAt = _startedAt,
                capturedSteps = _steps.Count
            };
        }
    }

    private static PlanStep CloneStep(PlanStep step)
    {
        return new PlanStep
        {
            Type = step.Type,
            Selector = step.Selector == null ? null : new Selector
            {
                Role = step.Selector.Role,
                NameContains = step.Selector.NameContains,
                AutomationId = step.Selector.AutomationId,
                ClassName = step.Selector.ClassName,
                AncestorNameContains = step.Selector.AncestorNameContains,
                Index = step.Selector.Index,
                WindowId = step.Selector.WindowId,
                BoundsHint = step.Selector.BoundsHint == null ? null : new Rect
                {
                    X = step.Selector.BoundsHint.X,
                    Y = step.Selector.BoundsHint.Y,
                    Width = step.Selector.BoundsHint.Width,
                    Height = step.Selector.BoundsHint.Height
                }
            },
            ExpectedAppId = step.ExpectedAppId,
            ExpectedWindowId = step.ExpectedWindowId,
            Text = step.Text,
            Target = step.Target,
            AppIdOrPath = step.AppIdOrPath,
            Keys = step.Keys == null ? null : new List<string>(step.Keys),
            Point = step.Point == null ? null : new Rect { X = step.Point.X, Y = step.Point.Y, Width = step.Point.Width, Height = step.Point.Height },
            ElementId = step.ElementId,
            WaitFor = step.WaitFor,
            Note = step.Note
        };
    }
}

internal sealed class AutoApproveConfirmation : IUserConfirmation
{
    public Task<bool> ConfirmAsync(string message, CancellationToken cancellationToken)
    {
        return Task.FromResult(true);
    }
}

internal sealed record LlmStatus(bool Enabled, bool Available, string Provider, string Message, string? Endpoint);
internal sealed record LlmConfigUpdate(bool? Enabled, bool? AllowNonLoopbackEndpoint, string? Provider, string? Endpoint, string? Model, int? TimeoutSeconds, int? MaxTokens);
internal sealed record OcrConfigUpdate(string? Engine, string? TesseractPath);
internal sealed record ConfigUpdateRequest(
    LlmConfigUpdate? Llm,
    List<string>? AllowedApps,
    Dictionary<string, string>? AppAliases,
    bool? ProfileModeEnabled,
    string? ActiveProfile,
    bool? RequireConfirmation,
    int? MaxActionsPerSecond,
    bool? QuizSafeModeEnabled,
    bool? OcrEnabled,
    OcrConfigUpdate? Ocr,
    string? AdapterRestartCommand,
    string? AdapterRestartWorkingDir,
    int? FindRetryCount,
    int? FindRetryDelayMs,
    int? PostCheckTimeoutMs,
    int? PostCheckPollMs,
    int? ClipboardHistoryMaxItems,
    List<string>? FilesystemAllowedRoots,
    bool? ContextBindingEnabled,
    bool? ContextBindingRequireWindow,
    string? TaskLibraryPath,
    string? ScheduleLibraryPath,
    bool? AuditLlmInteractions,
    bool? AuditLlmIncludeRawText);

internal sealed record OcrSnapshot(bool Enabled, string? Engine, string? TesseractPath);
internal sealed class UiTreeSummaryNode
{
    public string Role { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
    public string AutomationId { get; set; } = string.Empty;
    public int Depth { get; set; }
    public List<UiTreeSummaryNode> Children { get; set; } = new();
}

internal sealed record TaskLibraryItem(string Name, string Intent, string? Description, string? PlanJson, DateTimeOffset UpdatedAt);

internal sealed class TaskLibraryStore
{
    private readonly string _path;
    private readonly object _lock = new();

    public TaskLibraryStore(AgentConfig config)
    {
        var relative = string.IsNullOrWhiteSpace(config.TaskLibraryPath) ? "tasks.library.json" : config.TaskLibraryPath;
        _path = Path.GetFullPath(relative, AppContext.BaseDirectory);
    }

    public Task<IReadOnlyList<TaskLibraryItem>> ListAsync(CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        lock (_lock)
        {
            return Task.FromResult((IReadOnlyList<TaskLibraryItem>)LoadUnsafe()
                .OrderBy(item => item.Name, StringComparer.OrdinalIgnoreCase)
                .ToList());
        }
    }

    public Task<TaskLibraryItem?> GetAsync(string name, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        lock (_lock)
        {
            var item = LoadUnsafe().FirstOrDefault(task => string.Equals(task.Name, name, StringComparison.OrdinalIgnoreCase));
            return Task.FromResult(item);
        }
    }

    public Task UpsertAsync(string name, string intent, string? description, string? planJson, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        lock (_lock)
        {
            var items = LoadUnsafe();
            var index = items.FindIndex(item => string.Equals(item.Name, name, StringComparison.OrdinalIgnoreCase));
            var entry = new TaskLibraryItem(name, intent, description, planJson, DateTimeOffset.UtcNow);
            if (index >= 0)
            {
                items[index] = entry;
            }
            else
            {
                items.Add(entry);
            }
            SaveUnsafe(items);
        }
        return Task.CompletedTask;
    }

    public Task<bool> DeleteAsync(string name, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        lock (_lock)
        {
            var items = LoadUnsafe();
            var removed = items.RemoveAll(item => string.Equals(item.Name, name, StringComparison.OrdinalIgnoreCase)) > 0;
            if (removed)
            {
                SaveUnsafe(items);
            }
            return Task.FromResult(removed);
        }
    }

    private List<TaskLibraryItem> LoadUnsafe()
    {
        try
        {
            if (!File.Exists(_path))
            {
                return new List<TaskLibraryItem>();
            }

            var json = File.ReadAllText(_path);
            var items = JsonSerializer.Deserialize<List<TaskLibraryItem>>(json);
            return items ?? new List<TaskLibraryItem>();
        }
        catch
        {
            return new List<TaskLibraryItem>();
        }
    }

    private void SaveUnsafe(List<TaskLibraryItem> items)
    {
        var dir = Path.GetDirectoryName(_path);
        if (!string.IsNullOrWhiteSpace(dir))
        {
            Directory.CreateDirectory(dir);
        }

        var json = JsonSerializer.Serialize(items, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(_path, json);
    }
}

internal sealed record ScheduledTaskItem(
    string Id,
    string TaskName,
    bool Enabled,
    DateTimeOffset? StartAtUtc,
    int? IntervalSeconds,
    DateTimeOffset? NextRunAtUtc,
    DateTimeOffset? LastRunAtUtc,
    bool? LastSuccess,
    string? LastMessage);

internal sealed class ScheduleLibraryStore
{
    private readonly string _path;
    private readonly object _lock = new();

    public ScheduleLibraryStore(AgentConfig config)
    {
        var relative = string.IsNullOrWhiteSpace(config.ScheduleLibraryPath) ? "schedules.library.json" : config.ScheduleLibraryPath;
        _path = Path.GetFullPath(relative, AppContext.BaseDirectory);
    }

    public Task<IReadOnlyList<ScheduledTaskItem>> ListAsync(CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        lock (_lock)
        {
            return Task.FromResult((IReadOnlyList<ScheduledTaskItem>)LoadUnsafe()
                .OrderBy(item => item.TaskName, StringComparer.OrdinalIgnoreCase)
                .ThenBy(item => item.Id, StringComparer.OrdinalIgnoreCase)
                .ToList());
        }
    }

    public Task<ScheduledTaskItem?> GetAsync(string id, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        lock (_lock)
        {
            var item = LoadUnsafe().FirstOrDefault(entry => string.Equals(entry.Id, id, StringComparison.OrdinalIgnoreCase));
            return Task.FromResult(item);
        }
    }

    public Task<ScheduledTaskItem> UpsertAsync(
        string? id,
        string taskName,
        DateTimeOffset? startAtUtc,
        int? intervalSeconds,
        bool enabled,
        CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        lock (_lock)
        {
            var items = LoadUnsafe();
            var now = DateTimeOffset.UtcNow;
            var normalizedId = string.IsNullOrWhiteSpace(id) ? Guid.NewGuid().ToString("n") : id.Trim();
            var normalizedTask = taskName.Trim();
            var normalizedInterval = intervalSeconds.HasValue && intervalSeconds.Value > 0 ? intervalSeconds.Value : (int?)null;
            var normalizedStart = startAtUtc?.ToUniversalTime() ?? now;
            DateTimeOffset? nextRun = enabled ? ComputeNextRun(normalizedStart, normalizedInterval, now) : null;

            var existingIndex = items.FindIndex(item => string.Equals(item.Id, normalizedId, StringComparison.OrdinalIgnoreCase));
            var existing = existingIndex >= 0 ? items[existingIndex] : null;

            var updated = new ScheduledTaskItem(
                Id: normalizedId,
                TaskName: normalizedTask,
                Enabled: enabled,
                StartAtUtc: normalizedStart,
                IntervalSeconds: normalizedInterval,
                NextRunAtUtc: nextRun,
                LastRunAtUtc: existing?.LastRunAtUtc,
                LastSuccess: existing?.LastSuccess,
                LastMessage: existing?.LastMessage);

            if (existingIndex >= 0)
            {
                items[existingIndex] = updated;
            }
            else
            {
                items.Add(updated);
            }

            SaveUnsafe(items);
            return Task.FromResult(updated);
        }
    }

    public Task<bool> DeleteAsync(string id, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        lock (_lock)
        {
            var items = LoadUnsafe();
            var removed = items.RemoveAll(item => string.Equals(item.Id, id, StringComparison.OrdinalIgnoreCase)) > 0;
            if (removed)
            {
                SaveUnsafe(items);
            }

            return Task.FromResult(removed);
        }
    }

    public Task<IReadOnlyList<ScheduledTaskItem>> GetDueAsync(DateTimeOffset nowUtc, int maxCount, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        lock (_lock)
        {
            var due = LoadUnsafe()
                .Where(item => item.Enabled && item.NextRunAtUtc.HasValue && item.NextRunAtUtc.Value <= nowUtc)
                .OrderBy(item => item.NextRunAtUtc)
                .Take(Math.Clamp(maxCount, 1, 100))
                .ToList();
            return Task.FromResult((IReadOnlyList<ScheduledTaskItem>)due);
        }
    }

    public Task<ScheduledTaskItem?> MarkRunResultAsync(string id, DateTimeOffset runAtUtc, bool success, string message, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        lock (_lock)
        {
            var items = LoadUnsafe();
            var index = items.FindIndex(item => string.Equals(item.Id, id, StringComparison.OrdinalIgnoreCase));
            if (index < 0)
            {
                return Task.FromResult<ScheduledTaskItem?>(null);
            }

            var existing = items[index];
            DateTimeOffset? nextRun = null;
            var enabled = existing.Enabled;

            if (existing.Enabled)
            {
                if (existing.IntervalSeconds.HasValue && existing.IntervalSeconds.Value > 0)
                {
                    var start = existing.StartAtUtc ?? runAtUtc;
                    nextRun = ComputeNextRun(start, existing.IntervalSeconds, runAtUtc);
                }
                else
                {
                    enabled = false;
                }
            }

            var updated = existing with
            {
                Enabled = enabled,
                NextRunAtUtc = nextRun,
                LastRunAtUtc = runAtUtc,
                LastSuccess = success,
                LastMessage = message
            };
            items[index] = updated;
            SaveUnsafe(items);
            return Task.FromResult<ScheduledTaskItem?>(updated);
        }
    }

    private static DateTimeOffset ComputeNextRun(DateTimeOffset startAtUtc, int? intervalSeconds, DateTimeOffset nowUtc)
    {
        if (!intervalSeconds.HasValue || intervalSeconds.Value <= 0)
        {
            return startAtUtc > nowUtc ? startAtUtc : nowUtc;
        }

        if (startAtUtc > nowUtc)
        {
            return startAtUtc;
        }

        var interval = TimeSpan.FromSeconds(intervalSeconds.Value);
        if (interval <= TimeSpan.Zero)
        {
            return nowUtc;
        }

        var elapsed = nowUtc - startAtUtc;
        var steps = (long)Math.Floor(elapsed.TotalSeconds / interval.TotalSeconds) + 1;
        return startAtUtc + TimeSpan.FromSeconds(steps * interval.TotalSeconds);
    }

    private List<ScheduledTaskItem> LoadUnsafe()
    {
        try
        {
            if (!File.Exists(_path))
            {
                return new List<ScheduledTaskItem>();
            }

            var json = File.ReadAllText(_path);
            var items = JsonSerializer.Deserialize<List<ScheduledTaskItem>>(json);
            return items ?? new List<ScheduledTaskItem>();
        }
        catch
        {
            return new List<ScheduledTaskItem>();
        }
    }

    private void SaveUnsafe(List<ScheduledTaskItem> items)
    {
        var dir = Path.GetDirectoryName(_path);
        if (!string.IsNullOrWhiteSpace(dir))
        {
            Directory.CreateDirectory(dir);
        }

        var json = JsonSerializer.Serialize(items, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(_path, json);
    }
}

internal sealed record ScheduledRunResult(bool Found, bool Success, string Message);

internal sealed class ScheduledTaskRunner : BackgroundService
{
    private readonly IServiceScopeFactory _scopeFactory;
    private readonly ScheduleLibraryStore _schedules;
    private readonly IAuditLog _auditLog;
    private readonly ILogger<ScheduledTaskRunner> _logger;
    private readonly HashSet<string> _running = new(StringComparer.OrdinalIgnoreCase);
    private readonly object _runningLock = new();

    public ScheduledTaskRunner(
        IServiceScopeFactory scopeFactory,
        ScheduleLibraryStore schedules,
        IAuditLog auditLog,
        ILogger<ScheduledTaskRunner> logger)
    {
        _scopeFactory = scopeFactory;
        _schedules = schedules;
        _auditLog = auditLog;
        _logger = logger;
    }

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        while (!stoppingToken.IsCancellationRequested)
        {
            try
            {
                var due = await _schedules.GetDueAsync(DateTimeOffset.UtcNow, 8, stoppingToken);
                foreach (var item in due)
                {
                    await RunScheduleInternalAsync(item, triggeredByTimer: true, stoppingToken);
                }
            }
            catch (OperationCanceledException) when (stoppingToken.IsCancellationRequested)
            {
                break;
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Schedule runner iteration failed");
            }

            await Task.Delay(1000, stoppingToken);
        }
    }

    public async Task<ScheduledRunResult> RunNowAsync(string scheduleId, CancellationToken cancellationToken)
    {
        var schedule = await _schedules.GetAsync(scheduleId, cancellationToken);
        if (schedule == null)
        {
            return new ScheduledRunResult(false, false, "Schedule not found.");
        }

        return await RunScheduleInternalAsync(schedule, triggeredByTimer: false, cancellationToken);
    }

    private async Task<ScheduledRunResult> RunScheduleInternalAsync(ScheduledTaskItem schedule, bool triggeredByTimer, CancellationToken cancellationToken)
    {
        if (!TryEnterRunning(schedule.Id))
        {
            return new ScheduledRunResult(true, false, "Schedule already running.");
        }

        try
        {
            await _auditLog.WriteAsync(new AuditEvent
            {
                Timestamp = DateTimeOffset.UtcNow,
                EventType = "schedule_trigger",
                Message = triggeredByTimer ? "Scheduled task triggered" : "Scheduled task triggered manually",
                Data = new { schedule.Id, schedule.TaskName, schedule.IntervalSeconds }
            }, cancellationToken);

            using var scope = _scopeFactory.CreateScope();
            var tasks = scope.ServiceProvider.GetRequiredService<TaskLibraryStore>();
            var planner = scope.ServiceProvider.GetRequiredService<IPlanner>();

            var storedTask = await tasks.GetAsync(schedule.TaskName, cancellationToken);
            if (storedTask == null)
            {
                await _schedules.MarkRunResultAsync(schedule.Id, DateTimeOffset.UtcNow, success: false, "Task not found.", cancellationToken);
                return new ScheduledRunResult(true, false, $"Task '{schedule.TaskName}' not found.");
            }

            ActionPlan plan;
            if (!string.IsNullOrWhiteSpace(storedTask.PlanJson))
            {
                ActionPlan? parsedPlan;
                try
                {
                    parsedPlan = JsonSerializer.Deserialize<ActionPlan>(storedTask.PlanJson, new JsonSerializerOptions
                    {
                        PropertyNameCaseInsensitive = true
                    });
                }
                catch (Exception ex)
                {
                    parsedPlan = null;
                    await _schedules.MarkRunResultAsync(schedule.Id, DateTimeOffset.UtcNow, success: false, $"Invalid task plan: {ex.Message}", cancellationToken);
                    return new ScheduledRunResult(true, false, $"Invalid task plan: {ex.Message}");
                }

                if (parsedPlan == null || parsedPlan.Steps.Count == 0)
                {
                    await _schedules.MarkRunResultAsync(schedule.Id, DateTimeOffset.UtcNow, success: false, "Invalid task plan: empty plan", cancellationToken);
                    return new ScheduledRunResult(true, false, "Invalid task plan: empty plan");
                }

                plan = parsedPlan;
            }
            else
            {
                plan = planner.PlanFromIntent(storedTask.Intent);
            }

            var executor = new Executor(
                scope.ServiceProvider.GetRequiredService<IDesktopAdapterClient>(),
                scope.ServiceProvider.GetRequiredService<IContextProvider>(),
                scope.ServiceProvider.GetRequiredService<IAppResolver>(),
                scope.ServiceProvider.GetRequiredService<IPolicyEngine>(),
                scope.ServiceProvider.GetRequiredService<IRateLimiter>(),
                scope.ServiceProvider.GetRequiredService<IAuditLog>(),
                new RejectConfirmation(),
                scope.ServiceProvider.GetRequiredService<IKillSwitch>(),
                scope.ServiceProvider.GetRequiredService<AgentConfig>(),
                scope.ServiceProvider.GetRequiredService<IOcrEngine>(),
                scope.ServiceProvider.GetRequiredService<ILogger<Executor>>());

            var result = await executor.ExecutePlanAsync(plan, dryRun: false, cancellationToken);
            await _schedules.MarkRunResultAsync(schedule.Id, DateTimeOffset.UtcNow, result.Success, result.Message, cancellationToken);
            await _auditLog.WriteAsync(new AuditEvent
            {
                Timestamp = DateTimeOffset.UtcNow,
                EventType = "schedule_result",
                Message = $"Schedule execution {(result.Success ? "succeeded" : "failed")}",
                Data = new { schedule.Id, schedule.TaskName, result.Success, result.Message }
            }, cancellationToken);

            return new ScheduledRunResult(true, result.Success, result.Message);
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            return new ScheduledRunResult(true, false, "Cancelled.");
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Schedule execution failed for {ScheduleId}", schedule.Id);
            await _schedules.MarkRunResultAsync(schedule.Id, DateTimeOffset.UtcNow, success: false, ex.Message, cancellationToken);
            await _auditLog.WriteAsync(new AuditEvent
            {
                Timestamp = DateTimeOffset.UtcNow,
                EventType = "schedule_result",
                Message = "Schedule execution failed",
                Data = new { schedule.Id, schedule.TaskName, success = false, error = ex.Message }
            }, cancellationToken);
            return new ScheduledRunResult(true, false, ex.Message);
        }
        finally
        {
            ExitRunning(schedule.Id);
        }
    }

    private bool TryEnterRunning(string id)
    {
        lock (_runningLock)
        {
            return _running.Add(id);
        }
    }

    private void ExitRunning(string id)
    {
        lock (_runningLock)
        {
            _running.Remove(id);
        }
    }

    private sealed class RejectConfirmation : IUserConfirmation
    {
        public Task<bool> ConfirmAsync(string message, CancellationToken cancellationToken)
        {
            return Task.FromResult(false);
        }
    }
}

internal sealed class LlmAvailabilityCache
{
    private readonly AgentConfig _config;
    private readonly object _lock = new();
    private readonly TimeSpan _ttl = TimeSpan.FromSeconds(5);
    private DateTimeOffset _nextCheck = DateTimeOffset.MinValue;
    private LlmStatus _cached = new(false, false, "unknown", "Disabled in config", null);

    public LlmAvailabilityCache(AgentConfig config)
    {
        _config = config;
    }

    public async Task<LlmStatus> GetAsync(CancellationToken cancellationToken)
    {
        var now = DateTimeOffset.UtcNow;
        lock (_lock)
        {
            if (now < _nextCheck)
            {
                return _cached;
            }
        }

        var status = await ProbeAsync(cancellationToken);
        lock (_lock)
        {
            _cached = status;
            _nextCheck = DateTimeOffset.UtcNow.Add(_ttl);
        }
        return status;
    }

    public void Invalidate()
    {
        lock (_lock)
        {
            _nextCheck = DateTimeOffset.MinValue;
        }
    }

    private async Task<LlmStatus> ProbeAsync(CancellationToken cancellationToken)
    {
        var provider = NormalizeProvider(_config.LlmFallback.Provider);
        var endpoint = _config.LlmFallback.Endpoint?.Trim();
        if (!_config.LlmFallbackEnabled)
        {
            return new LlmStatus(false, false, provider, "Disabled in config", endpoint);
        }

        if (string.IsNullOrWhiteSpace(endpoint))
        {
            return new LlmStatus(true, false, provider, "Endpoint not configured", null);
        }

        if (!Uri.TryCreate(endpoint, UriKind.Absolute, out var uri))
        {
            return new LlmStatus(true, false, provider, "Invalid endpoint", endpoint);
        }

        if (!ConfigValidators.IsEndpointAllowed(uri, _config.AllowNonLoopbackLlmEndpoint))
        {
            return new LlmStatus(true, false, provider, "Endpoint blocked by policy (local only)", endpoint);
        }

        if (provider == "ollama")
        {
            return await ProbeOllamaStatusAsync(uri, endpoint, cancellationToken);
        }

        var port = uri.Port > 0 ? uri.Port : (uri.Scheme == "https" ? 443 : 80);
        var available = await CanConnectAsync(uri.Host, port, cancellationToken);
        return available
            ? new LlmStatus(true, true, provider, "Available", endpoint)
            : new LlmStatus(true, false, provider, $"No listener on {uri.Host}:{port}", endpoint);
    }

    private async Task<LlmStatus> ProbeOllamaStatusAsync(Uri endpointUri, string? endpoint, CancellationToken cancellationToken)
    {
        try
        {
            using var client = new HttpClient { Timeout = TimeSpan.FromSeconds(Math.Max(2, _config.LlmFallback.TimeoutSeconds)) };
            var baseUri = new Uri(endpointUri.GetLeftPart(UriPartial.Authority));
            var tagsUri = new Uri(baseUri, "/api/tags");

            using var response = await client.GetAsync(tagsUri, cancellationToken);
            if (!response.IsSuccessStatusCode)
            {
                return new LlmStatus(true, false, "ollama", $"HTTP {(int)response.StatusCode}", endpoint);
            }

            var configuredModel = (_config.LlmFallback.Model ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(configuredModel))
            {
                return new LlmStatus(true, true, "ollama", "Available (model not configured)", endpoint);
            }

            var installedModels = await ReadOllamaModelNamesAsync(response, cancellationToken);
            if (installedModels.Count == 0)
            {
                return new LlmStatus(true, false, "ollama", "Available but model list unreadable", endpoint);
            }

            var found = installedModels.Any(name => string.Equals(name, configuredModel, StringComparison.OrdinalIgnoreCase));
            if (found)
            {
                return new LlmStatus(true, true, "ollama", $"Available; model '{configuredModel}' found", endpoint);
            }

            var preview = string.Join(", ", installedModels.Take(5));
            if (installedModels.Count > 5)
            {
                preview += ", ...";
            }

            var message = string.IsNullOrWhiteSpace(preview)
                ? $"Available but model '{configuredModel}' not found"
                : $"Available but model '{configuredModel}' not found. Installed: {preview}";
            return new LlmStatus(true, false, "ollama", Compact(message, 220), endpoint);
        }
        catch (Exception ex)
        {
            return new LlmStatus(true, false, "ollama", Compact(ex.Message, 140), endpoint);
        }
    }

    private static async Task<List<string>> ReadOllamaModelNamesAsync(HttpResponseMessage response, CancellationToken cancellationToken)
    {
        try
        {
            await using var stream = await response.Content.ReadAsStreamAsync(cancellationToken);
            using var document = await JsonDocument.ParseAsync(stream, cancellationToken: cancellationToken);
            if (!document.RootElement.TryGetProperty("models", out var models) || models.ValueKind != JsonValueKind.Array)
            {
                return new List<string>();
            }

            var names = new List<string>();
            foreach (var item in models.EnumerateArray())
            {
                if (item.TryGetProperty("name", out var nameProp))
                {
                    var value = nameProp.GetString();
                    if (!string.IsNullOrWhiteSpace(value))
                    {
                        names.Add(value.Trim());
                    }
                }
            }

            return names;
        }
        catch
        {
            return new List<string>();
        }
    }

    private static async Task<bool> CanConnectAsync(string host, int port, CancellationToken cancellationToken)
    {
        try
        {
            using var client = new System.Net.Sockets.TcpClient();
            var connectTask = client.ConnectAsync(host, port);
            var timeoutTask = Task.Delay(600, cancellationToken);
            var completed = await Task.WhenAny(connectTask, timeoutTask);
            if (completed != connectTask)
            {
                return false;
            }
            await connectTask;
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static string NormalizeProvider(string? provider)
    {
        return string.IsNullOrWhiteSpace(provider) ? "local" : provider.Trim().ToLowerInvariant();
    }

    private static string Compact(string? value, int maxChars)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        var normalized = value.Replace("\r", " ").Replace("\n", " ").Trim();
        return normalized.Length <= maxChars ? normalized : normalized[..maxChars] + "...";
    }
}

internal sealed class RestartRequirementTracker
{
    public bool OcrRestartRequired { get; set; }
}

internal sealed class ConfigFileStore
{
    private readonly string _path;
    private readonly object _lock = new();

    public ConfigFileStore(string path)
    {
        _path = path;
    }

    public Task SaveAsync(AgentConfig config)
    {
        AgentConfigSanitizer.Normalize(config);
        var options = new JsonSerializerOptions
        {
            WriteIndented = true
        };
        var json = JsonSerializer.Serialize(config, options);
        lock (_lock)
        {
            File.WriteAllText(_path, json);
        }
        return Task.CompletedTask;
    }
}

internal static class ConfigValidators
{
    internal static bool IsAllowedProvider(string provider)
    {
        var normalized = provider.Trim().ToLowerInvariant();
        return normalized is "ollama" or "openai" or "llama.cpp";
    }

    internal static bool IsEndpointAllowed(Uri uri, bool allowNonLoopbackEndpoint)
    {
        return uri.IsLoopback || allowNonLoopbackEndpoint;
    }
}

internal static class AppVersionHelper
{
    internal static string Resolve()
    {
        var assembly = Assembly.GetExecutingAssembly();
        var informational = assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion;
        if (!string.IsNullOrWhiteSpace(informational))
        {
            return informational;
        }

        var fileVersion = assembly.GetCustomAttribute<AssemblyFileVersionAttribute>()?.Version;
        if (!string.IsNullOrWhiteSpace(fileVersion))
        {
            return fileVersion;
        }

        return assembly.GetName().Version?.ToString() ?? "unknown";
    }
}

internal static class CommandLineParser
{
    internal static bool TryParseCommand(string commandLine, out string fileName, out string args)
    {
        fileName = string.Empty;
        args = string.Empty;
        if (string.IsNullOrWhiteSpace(commandLine))
        {
            return false;
        }

        var tokens = SplitCommandLine(commandLine);
        if (tokens.Count == 0)
        {
            return false;
        }

        fileName = tokens[0];
        if (tokens.Count > 1)
        {
            args = string.Join(' ', tokens.Skip(1));
        }
        return true;
    }

    private static List<string> SplitCommandLine(string commandLine)
    {
        var results = new List<string>();
        var current = new System.Text.StringBuilder();
        var inQuotes = false;
        char quoteChar = '\0';

        for (var i = 0; i < commandLine.Length; i++)
        {
            var ch = commandLine[i];
            if (inQuotes)
            {
                if (ch == quoteChar)
                {
                    inQuotes = false;
                }
                else
                {
                    current.Append(ch);
                }
                continue;
            }

            if (ch == '"' || ch == '\'')
            {
                inQuotes = true;
                quoteChar = ch;
                continue;
            }

            if (char.IsWhiteSpace(ch))
            {
                if (current.Length > 0)
                {
                    results.Add(current.ToString());
                    current.Clear();
                }
                continue;
            }

            current.Append(ch);
        }

        if (current.Length > 0)
        {
            results.Add(current.ToString());
        }

        return results;
    }
}
