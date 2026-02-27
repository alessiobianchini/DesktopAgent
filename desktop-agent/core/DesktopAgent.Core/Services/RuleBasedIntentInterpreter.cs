using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using DesktopAgent.Core.Abstractions;
using DesktopAgent.Core.Config;
using DesktopAgent.Core.Models;
using DesktopAgent.Proto;

namespace DesktopAgent.Core.Services;

public sealed class RuleBasedIntentInterpreter : IIntentInterpreter
{
    private static readonly string[] OpenVerbs = { "open", "run", "start", "launch", "apri", "avvia", "esegui", "lancia" };
    private static readonly string[] TypeVerbs = { "type", "scrivi", "digita" };
    private static readonly string[] ClickVerbs = { "click", "clicca", "clic" };
    private static readonly string[] FindVerbs = { "find", "cerca", "trova" };
    private static readonly string[] PressVerbs = { "press", "premi" };

    private static readonly Regex OpenRegex = BuildVerbRegex(OpenVerbs);
    private static readonly Regex TypeRegex = BuildVerbRegex(TypeVerbs);
    private static readonly Regex ClickRegex = BuildVerbRegex(ClickVerbs);
    private static readonly Regex FindRegex = BuildVerbRegex(FindVerbs);
    private static readonly Regex PressRegex = BuildVerbRegex(PressVerbs);
    private static readonly Regex OpenAndNewFileRegex = BuildOpenAndNewFileRegex();
    private static readonly Regex NewFileRegex = BuildNewFileRegex();
    private static readonly Regex SaveRegex = BuildSimpleRegex("save", "salva", "save file", "salva file", "salva documento", "save document");
    private static readonly Regex SelectLineRegex = BuildSimpleRegex("select line", "seleziona riga", "seleziona linea");
    private static readonly Regex DeleteLineRegex = BuildSimpleRegex("delete line", "elimina riga", "cancella riga");
    private static readonly Regex ReplaceRegex = BuildReplaceRegex();
    private static readonly Regex SaveAsRegex = BuildSaveAsRegex();
    private static readonly Regex SaveAsOnlyRegex = BuildSimpleRegex("save as", "salva come");
    private static readonly Regex NewTabRegex = BuildSimpleRegex("new tab", "nuova scheda");
    private static readonly Regex CloseTabRegex = BuildSimpleRegex("close tab", "chiudi scheda", "chiudi la scheda");
    private static readonly Regex CloseWindowRegex = BuildSimpleRegex("close window", "chiudi finestra", "chiudi la finestra");
    private static readonly Regex MinimizeWindowRegex = BuildSimpleRegex("minimize window", "minimize", "riduci finestra", "minimizza finestra", "minimizza");
    private static readonly Regex MaximizeWindowRegex = BuildSimpleRegex("maximize window", "maximize", "ingrandisci finestra", "massimizza finestra", "massimizza");
    private static readonly Regex RestoreWindowRegex = BuildSimpleRegex("restore window", "restore", "ripristina finestra", "ripristina");
    private static readonly Regex SwitchWindowRegex = BuildSimpleRegex("switch window", "next window", "change window", "cambia finestra", "passa finestra");
    private static readonly Regex FocusAppRegex = BuildFocusAppRegex();
    private static readonly Regex ScrollRegex = BuildScrollRegex();
    private static readonly Regex PageUpRegex = BuildSimpleRegex("page up", "pagina su");
    private static readonly Regex PageDownRegex = BuildSimpleRegex("page down", "pagina giu", "pagina giù");
    private static readonly Regex HomeRegex = BuildSimpleRegex("home", "inizio");
    private static readonly Regex EndRegex = BuildSimpleRegex("end", "fine");
    private static readonly Regex DoubleClickRegex = BuildDoubleClickRegex();
    private static readonly Regex RightClickRegex = BuildRightClickRegex();
    private static readonly Regex DragRegex = BuildDragRegex();
    private static readonly Regex WaitUntilRegex = BuildWaitUntilRegex();
    private static readonly Regex RetryRegex = BuildRetryRegex();
    private static readonly Regex IfFoundThenRegex = BuildIfFoundThenRegex();
    private static readonly Regex VolumeRegex = BuildVolumeRegex();
    private static readonly Regex BrightnessRegex = BuildBrightnessRegex();
    private static readonly Regex LockScreenRegex = BuildSimpleRegex("lock screen", "lock workstation", "blocca schermo", "blocca pc");
    private static readonly Regex BrowserBackRegex = BuildSimpleRegex("back", "go back", "navigate back", "indietro", "vai indietro", "torna indietro");
    private static readonly Regex BrowserForwardRegex = BuildSimpleRegex("forward", "go forward", "navigate forward", "avanti", "vai avanti");
    private static readonly Regex RefreshRegex = BuildSimpleRegex("refresh", "reload", "ricarica", "aggiorna", "ricarica pagina", "reload page");
    private static readonly Regex FindInPageRegex = BuildFindInPageRegex();
    private static readonly Regex CopyRegex = BuildSimpleRegex("copy", "copia");
    private static readonly Regex PasteRegex = BuildSimpleRegex("paste", "incolla");
    private static readonly Regex UndoRegex = BuildSimpleRegex("undo", "annulla");
    private static readonly Regex RedoRegex = BuildSimpleRegex("redo", "ripeti");
    private static readonly Regex SelectAllRegex = BuildSimpleRegex("select all", "seleziona tutto");
    private static readonly Regex OpenUrlRegex = BuildOpenUrlRegex();
    private static readonly Regex NavigateQueryRegex = BuildNavigateQueryRegex();
    private static readonly Regex OpenUrlOnBrowserRegex = BuildOpenUrlOnBrowserRegex();
    private static readonly Regex SearchOnBrowserRegex = BuildSearchOnBrowserRegex();
    private static readonly Regex SearchWebRegex = BuildSearchWebRegex();
    private static readonly Regex FileWriteRegex = BuildFileWriteRegex();
    private static readonly Regex FileAppendRegex = BuildFileAppendRegex();
    private static readonly Regex FileReadRegex = BuildFileReadRegex();
    private static readonly Regex FileListRegex = BuildFileListRegex();
    private static readonly Regex NotifyRegex = BuildNotifyRegex();
    private static readonly Regex ClipboardHistoryRegex = BuildSimpleRegex("clipboard history", "show clipboard history", "cronologia clipboard");
    private static readonly Regex MouseDurationRegex = BuildMouseDurationRegex();

    private static readonly HashSet<string> KeyTokens = new(StringComparer.OrdinalIgnoreCase)
    {
        "enter", "return", "invio",
        "esc", "escape",
        "tab",
        "shift", "ctrl", "control", "alt",
        "win", "windows", "cmd",
        "backspace", "delete", "del",
        "home", "end", "pageup", "pagedown", "pgup", "pgdn",
        "up", "down", "left", "right",
        "freccia", "su", "giu", "sinistra", "destra"
    };

    private static readonly HashSet<string> KeyFillerTokens = new(StringComparer.OrdinalIgnoreCase)
    {
        "key", "tasto", "chiave"
    };

    private readonly IAppResolver _appResolver;
    private readonly AgentConfig _config;

    public RuleBasedIntentInterpreter(IAppResolver appResolver, AgentConfig config)
    {
        _appResolver = appResolver;
        _config = config;
    }

    public ActionPlan Interpret(string intent)
    {
        var plan = new ActionPlan { Intent = intent };
        var trimmed = (intent ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(trimmed))
        {
            plan.Steps.Add(new PlanStep { Type = ActionType.ReadText, Note = "Default to read text" });
            return plan;
        }

        if (TryMatch(OpenAndNewFileRegex, trimmed, out var app))
        {
            var appTarget = app;
            if (_appResolver.TryResolveApp(app, out var resolved))
            {
                appTarget = resolved;
            }
            plan.Steps.Add(new PlanStep { Type = ActionType.OpenApp, AppIdOrPath = appTarget });
            plan.Steps.Add(new PlanStep { Type = ActionType.KeyCombo, Keys = GetNewFileKeys(), ExpectedAppId = appTarget });
            return plan;
        }

        var segments = SplitSegments(trimmed);
        if (segments.Count > 1)
        {
            var unknownSegments = new List<string>();
            string? currentApp = null;
            string? lastOpenedApp = null;
            foreach (var segment in segments)
            {
                var (steps, recognized) = InterpretSingle(segment);
                if (steps.Count > 0)
                {
                    var openStep = steps.FirstOrDefault(step => step.Type == ActionType.OpenApp && !string.IsNullOrWhiteSpace(step.AppIdOrPath));
                    if (openStep != null)
                    {
                        currentApp = openStep.AppIdOrPath;
                        lastOpenedApp = currentApp;
                        ApplyExpectedAppBinding(steps, currentApp);
                    }
                    else if (!string.IsNullOrWhiteSpace(currentApp))
                    {
                        var needsOpen = lastOpenedApp == null
                                        || !string.Equals(lastOpenedApp, currentApp, StringComparison.OrdinalIgnoreCase);
                        if (needsOpen)
                        {
                            steps.Insert(0, new PlanStep { Type = ActionType.OpenApp, AppIdOrPath = currentApp });
                            lastOpenedApp = currentApp;
                        }
                        ApplyExpectedAppBinding(steps, currentApp);
                    }
                    plan.Steps.AddRange(steps);
                }
                if (!recognized)
                {
                    unknownSegments.Add(segment);
                }
            }

            if (plan.Steps.Count == 0)
            {
                plan.Steps.Add(new PlanStep { Type = ActionType.ReadText, Note = "Default to read text" });
                return plan;
            }

            if (unknownSegments.Count > 0)
            {
                plan.Steps.Add(new PlanStep
                {
                    Type = ActionType.ReadText,
                    Note = $"Unrecognized segments: {string.Join(" | ", unknownSegments)}"
                });
            }

            return plan;
        }

        var (singleSteps, singleRecognized) = InterpretSingle(trimmed);
        if (singleSteps.Count > 0)
        {
            plan.Steps.AddRange(singleSteps);
            return plan;
        }

        if (!singleRecognized)
        {
            plan.Steps.Add(new PlanStep { Type = ActionType.ReadText, Note = "Default to read text" });
        }
        return plan;
    }

    private static Regex BuildVerbRegex(IEnumerable<string> verbs)
    {
        var pattern = $"^(?:{string.Join("|", verbs.Select(Regex.Escape))})\\s+(.+)$";
        return new Regex(pattern, RegexOptions.IgnoreCase | RegexOptions.Compiled);
    }

    private static Regex BuildOpenAndNewFileRegex()
    {
        var verbs = string.Join("|", OpenVerbs.Select(Regex.Escape));
        var connector = "(?:and|e|then|poi|quindi)";
        var newFile = "(?:create|crea|new|nuovo)(?:\\s+(?:a|un|una))?\\s*(?:new|nuovo)?\\s*(?:text|testo|txt)?\\s*file";
        var pattern = $"^(?:{verbs})\\s+(.+?)\\s+{connector}\\s+{newFile}$";
        return new Regex(pattern, RegexOptions.IgnoreCase | RegexOptions.Compiled);
    }

    private static Regex BuildNewFileRegex()
    {
        var newFile = "^(?:create|crea|new|nuovo)(?:\\s+(?:a|un|una))?\\s*(?:new|nuovo)?\\s*(?:text|testo|txt)?\\s*file$";
        return new Regex(newFile, RegexOptions.IgnoreCase | RegexOptions.Compiled);
    }

    private static Regex BuildSaveAsRegex()
    {
        var pattern = "^(?:save as|salva come)\\s+(?<file>.+?)(?:\\s+(?:in|nel|nella)(?:\\s+(?:cartella|folder))?\\s+(?<folder>.+))?$";
        return new Regex(pattern, RegexOptions.IgnoreCase | RegexOptions.Compiled);
    }

    private static Regex BuildReplaceRegex()
    {
        var pattern = "^(?:replace|sostituisci|find and replace|trova e sostituisci)\\s+['\\\"]?(?<find>.+?)['\\\"]?\\s+(?:with|con)\\s+['\\\"]?(?<replace>.+?)['\\\"]?$";
        return new Regex(pattern, RegexOptions.IgnoreCase | RegexOptions.Compiled);
    }

    private static Regex BuildFocusAppRegex()
    {
        var pattern = "^(?:focus|bring to front|porta in primo piano|metti in primo piano)\\s+(?<app>.+)$";
        return new Regex(pattern, RegexOptions.IgnoreCase | RegexOptions.Compiled);
    }

    private static Regex BuildScrollRegex()
    {
        var pattern = "^(?:scroll)\\s+(?<dir>up|down)(?:\\s+(?<amount>\\d+))?$|^(?:scorri)\\s+(?<dirit>su|giu|giù)(?:\\s+(?<amountit>\\d+))?$";
        return new Regex(pattern, RegexOptions.IgnoreCase | RegexOptions.Compiled);
    }

    private static Regex BuildDoubleClickRegex()
    {
        var pattern = "^(?:double\\s*click|doppio\\s+clic)\\s+(?<target>.+)$";
        return new Regex(pattern, RegexOptions.IgnoreCase | RegexOptions.Compiled);
    }

    private static Regex BuildRightClickRegex()
    {
        var pattern = "^(?:right\\s*click|clic\\s+destro)\\s+(?<target>.+)$";
        return new Regex(pattern, RegexOptions.IgnoreCase | RegexOptions.Compiled);
    }

    private static Regex BuildDragRegex()
    {
        var pattern = "^(?:drag|trascina)\\s+(?<from>.+?)\\s+(?:to|in|su)\\s+(?<to>.+)$";
        return new Regex(pattern, RegexOptions.IgnoreCase | RegexOptions.Compiled);
    }

    private static Regex BuildWaitUntilRegex()
    {
        var pattern = "^(?:wait\\s+until|aspetta\\s+finche|aspetta\\s+finché)\\s+(?<text>.+?)(?:\\s+(?:for|per)\\s+(?<seconds>\\d+)\\s*(?:s|sec|secondi?|seconds?))?$";
        return new Regex(pattern, RegexOptions.IgnoreCase | RegexOptions.Compiled);
    }

    private static Regex BuildRetryRegex()
    {
        var pattern = "^(?:retry|riprova)\\s+(?<count>\\d+)\\s+(?:times?|volte)?\\s*(?<cmd>.+)$";
        return new Regex(pattern, RegexOptions.IgnoreCase | RegexOptions.Compiled);
    }

    private static Regex BuildIfFoundThenRegex()
    {
        var pattern = "^(?:if\\s+found|se\\s+trovi)\\s+(?<find>.+?)\\s+(?:then|allora)\\s+(?<then>.+)$";
        return new Regex(pattern, RegexOptions.IgnoreCase | RegexOptions.Compiled);
    }

    private static Regex BuildVolumeRegex()
    {
        var pattern = "^(?:volume|audio)\\s+(?<action>up|down|mute)(?:\\s+(?<amount>\\d+))?$|^(?:volume|audio)\\s+(?<actionit>su|giu|giù|muto)(?:\\s+(?<amountit>\\d+))?$";
        return new Regex(pattern, RegexOptions.IgnoreCase | RegexOptions.Compiled);
    }

    private static Regex BuildBrightnessRegex()
    {
        var pattern = "^(?:brightness|luminosita|luminosità)\\s+(?<action>up|down|su|giu|giù)(?:\\s+(?<amount>\\d+))?$";
        return new Regex(pattern, RegexOptions.IgnoreCase | RegexOptions.Compiled);
    }

    private static Regex BuildOpenUrlRegex()
    {
        var pattern = "^(?:open\\s+url|open\\s+website|open\\s+site|apri\\s+url|apri\\s+sito|go\\s+to|navigate\\s+to|vai\\s+su|vai\\s+a|naviga\\s+a)\\s+(?<url>.+)$";
        return new Regex(pattern, RegexOptions.IgnoreCase | RegexOptions.Compiled);
    }

    private static Regex BuildNavigateQueryRegex()
    {
        var pattern = "^(?:go\\s+to|navigate\\s+to|vai\\s+su|vai\\s+a|naviga\\s+a)\\s+(?<query>.+)$";
        return new Regex(pattern, RegexOptions.IgnoreCase | RegexOptions.Compiled);
    }

    private static Regex BuildOpenUrlOnBrowserRegex()
    {
        var pattern = "^(?:open\\s+url|open\\s+website|open\\s+site|apri\\s+url|apri\\s+sito|go\\s+to|navigate\\s+to|vai\\s+su|vai\\s+a|naviga\\s+a)\\s+(?<url>.+?)\\s+(?:on|in|su|nel|nella)\\s+(?<browser>chrome|google\\s+chrome|edge|microsoft\\s+edge|firefox|brave|safari|browser|web\\s+browser)$";
        return new Regex(pattern, RegexOptions.IgnoreCase | RegexOptions.Compiled);
    }

    private static Regex BuildFindInPageRegex()
    {
        var pattern = "^(?:find\\s+in\\s+page|search\\s+in\\s+page|cerca\\s+nella\\s+pagina|trova\\s+nella\\s+pagina)\\s+(?<text>.+)$";
        return new Regex(pattern, RegexOptions.IgnoreCase | RegexOptions.Compiled);
    }

    private static Regex BuildSearchOnBrowserRegex()
    {
        var pattern = "^(?:search|cerca|trova)(?:\\s+(?:for|di))?\\s+(?<query>.+?)\\s+(?:on|in|su|nel|nella)\\s+(?<browser>chrome|google\\s+chrome|edge|microsoft\\s+edge|firefox|brave|safari|browser|web\\s+browser)$";
        return new Regex(pattern, RegexOptions.IgnoreCase | RegexOptions.Compiled);
    }

    private static Regex BuildSearchWebRegex()
    {
        var pattern = "^(?:search|cerca|trova)(?:\\s+(?:for|di))?\\s+(?<query>.+)$";
        return new Regex(pattern, RegexOptions.IgnoreCase | RegexOptions.Compiled);
    }

    private static Regex BuildFileWriteRegex()
    {
        var pattern = "^file\\s+(?:write|scrivi)\\s+(?<path>\"[^\"]+\"|'[^']+'|\\S+)\\s+(?<text>.+)$";
        return new Regex(pattern, RegexOptions.IgnoreCase | RegexOptions.Compiled);
    }

    private static Regex BuildFileAppendRegex()
    {
        var pattern = "^file\\s+(?:append|aggiungi)\\s+(?<path>\"[^\"]+\"|'[^']+'|\\S+)\\s+(?<text>.+)$";
        return new Regex(pattern, RegexOptions.IgnoreCase | RegexOptions.Compiled);
    }

    private static Regex BuildFileReadRegex()
    {
        var pattern = "^file\\s+(?:read|leggi)\\s+(?<path>\"[^\"]+\"|'[^']+'|\\S+)$";
        return new Regex(pattern, RegexOptions.IgnoreCase | RegexOptions.Compiled);
    }

    private static Regex BuildFileListRegex()
    {
        var pattern = "^file\\s+(?:list|ls|elenca)\\s*(?<path>\"[^\"]+\"|'[^']+'|\\S+)?$";
        return new Regex(pattern, RegexOptions.IgnoreCase | RegexOptions.Compiled);
    }

    private static Regex BuildNotifyRegex()
    {
        var pattern = "^(?:notify|notification|notifica)\\s+(?<text>.+)$";
        return new Regex(pattern, RegexOptions.IgnoreCase | RegexOptions.Compiled);
    }

    private static Regex BuildMouseDurationRegex()
    {
        var pattern = "(?<value>\\d+(?:[\\.,]\\d+)?)\\s*(?<unit>seconds?|secondi?|sec(?:onds?)?|secs?|s|minutes?|minuti?|mins?|munutes?|minites?|m|hours?|ore|hrs?|h)";
        return new Regex(pattern, RegexOptions.IgnoreCase | RegexOptions.Compiled);
    }

    private static Regex BuildSimpleRegex(params string[] phrases)
    {
        var pattern = $"^(?:{string.Join("|", phrases.Select(Regex.Escape))})$";
        return new Regex(pattern, RegexOptions.IgnoreCase | RegexOptions.Compiled);
    }

    private static List<string> SplitSegments(string input)
    {
        var tokens = TokenizeWithQuotes(input);
        var segments = new List<string>();
        var current = new List<string>();

        for (var i = 0; i < tokens.Count; i++)
        {
            var token = tokens[i];
            var lower = token.ToLowerInvariant();
            var next = i + 1 < tokens.Count ? tokens[i + 1].ToLowerInvariant() : string.Empty;

            if (IsConnector(lower, next))
            {
                if (lower == "and" && next == "then")
                {
                    i++;
                }
                if (lower == "e" && next == "poi")
                {
                    i++;
                }

                if (current.Count > 0)
                {
                    segments.Add(string.Join(' ', current));
                    current.Clear();
                }
                continue;
            }

            current.Add(token);
        }

        if (current.Count > 0)
        {
            segments.Add(string.Join(' ', current));
        }

        return segments.Where(segment => !string.IsNullOrWhiteSpace(segment)).ToList();
    }

    private static bool IsConnector(string token, string next)
    {
        if (token is "and" or "then" or "after" or "e" or "poi" or "quindi" or "dopo" or "successivamente")
        {
            return true;
        }

        if (token == "and" && next == "then")
        {
            return true;
        }

        if (token == "e" && next == "poi")
        {
            return true;
        }

        return false;
    }

    private static List<string> TokenizeWithQuotes(string input)
    {
        var tokens = new List<string>();
        var current = new List<char>();
        var inQuotes = false;
        char quoteChar = '\0';

        foreach (var ch in input)
        {
            if ((ch == '"' || ch == '\'') && (!inQuotes || ch == quoteChar))
            {
                if (inQuotes)
                {
                    inQuotes = false;
                    quoteChar = '\0';
                }
                else
                {
                    inQuotes = true;
                    quoteChar = ch;
                }
                continue;
            }

            if (char.IsWhiteSpace(ch) && !inQuotes)
            {
                FlushToken(tokens, current);
                continue;
            }

            current.Add(ch);
        }

        FlushToken(tokens, current);
        return tokens;
    }

    private static void FlushToken(List<string> tokens, List<char> current)
    {
        if (current.Count == 0)
        {
            return;
        }

        tokens.Add(new string(current.ToArray()));
        current.Clear();
    }

    private (List<PlanStep> Steps, bool Recognized) InterpretSingle(string input)
    {
        if (TryParseMouseJiggle(input, out var duration))
        {
            return (new List<PlanStep>
            {
                new()
                {
                    Type = ActionType.MouseJiggle,
                    WaitFor = duration
                }
            }, true);
        }

        if (RetryRegex.IsMatch(input))
        {
            var match = RetryRegex.Match(input);
            var cmd = match.Groups["cmd"].Value.Trim();
            var count = int.TryParse(match.Groups["count"].Value, out var parsedCount) ? Math.Clamp(parsedCount, 1, 10) : 1;
            var (retriedSteps, retriedRecognized) = InterpretSingle(cmd);
            if (retriedRecognized && retriedSteps.Count > 0)
            {
                var repeated = new List<PlanStep>();
                for (var i = 0; i < count; i++)
                {
                    repeated.AddRange(CloneSteps(retriedSteps));
                }

                return (repeated, true);
            }
        }

        if (IfFoundThenRegex.IsMatch(input))
        {
            var match = IfFoundThenRegex.Match(input);
            var findText = CleanQuoted(match.Groups["find"].Value);
            var thenText = match.Groups["then"].Value.Trim();
            if (!string.IsNullOrWhiteSpace(findText))
            {
                var (thenSteps, thenRecognized) = InterpretSingle(thenText);
                if (thenRecognized && thenSteps.Count > 0)
                {
                    var steps = new List<PlanStep>
                    {
                        new()
                        {
                            Type = ActionType.Find,
                            Selector = new Selector { NameContains = findText },
                            Note = "if-found-guard"
                        }
                    };

                    steps.AddRange(CloneSteps(thenSteps).Select(step =>
                    {
                        step.Note = string.IsNullOrWhiteSpace(step.Note)
                            ? "if-found-dependent"
                            : $"{step.Note};if-found-dependent";
                        return step;
                    }));
                    return (steps, true);
                }
            }
        }

        if (WaitUntilRegex.IsMatch(input))
        {
            var match = WaitUntilRegex.Match(input);
            var waitText = CleanQuoted(match.Groups["text"].Value);
            if (!string.IsNullOrWhiteSpace(waitText))
            {
                var timeout = TimeSpan.FromSeconds(30);
                if (int.TryParse(match.Groups["seconds"].Value, out var seconds))
                {
                    timeout = TimeSpan.FromSeconds(Math.Clamp(seconds, 1, 600));
                }

                return (new List<PlanStep>
                {
                    new()
                    {
                        Type = ActionType.WaitForText,
                        Text = waitText,
                        WaitFor = timeout
                    }
                }, true);
            }
        }

        if (LockScreenRegex.IsMatch(input))
        {
            return (new List<PlanStep> { new() { Type = ActionType.LockScreen } }, true);
        }

        if (VolumeRegex.IsMatch(input))
        {
            var match = VolumeRegex.Match(input);
            var action = match.Groups["action"].Success ? match.Groups["action"].Value : match.Groups["actionit"].Value;
            var amountText = match.Groups["amount"].Success ? match.Groups["amount"].Value : match.Groups["amountit"].Value;
            var amount = int.TryParse(amountText, out var amountValue) ? Math.Clamp(amountValue, 1, 20) : 1;
            return action.ToLowerInvariant() switch
            {
                "up" or "su" => (new List<PlanStep> { new() { Type = ActionType.VolumeUp, Text = amount.ToString() } }, true),
                "down" or "giu" or "giù" => (new List<PlanStep> { new() { Type = ActionType.VolumeDown, Text = amount.ToString() } }, true),
                _ => (new List<PlanStep> { new() { Type = ActionType.VolumeMute } }, true)
            };
        }

        if (BrightnessRegex.IsMatch(input))
        {
            var match = BrightnessRegex.Match(input);
            var action = match.Groups["action"].Value.ToLowerInvariant();
            var amount = int.TryParse(match.Groups["amount"].Value, out var amountValue) ? Math.Clamp(amountValue, 1, 50) : 10;
            if (action is "up" or "su")
            {
                return (new List<PlanStep> { new() { Type = ActionType.BrightnessUp, Text = amount.ToString() } }, true);
            }

            return (new List<PlanStep> { new() { Type = ActionType.BrightnessDown, Text = amount.ToString() } }, true);
        }

        if (DoubleClickRegex.IsMatch(input))
        {
            var match = DoubleClickRegex.Match(input);
            var target = CleanQuoted(match.Groups["target"].Value);
            if (!string.IsNullOrWhiteSpace(target))
            {
                return (new List<PlanStep>
                {
                    new()
                    {
                        Type = ActionType.DoubleClick,
                        Selector = new Selector { NameContains = target }
                    }
                }, true);
            }
        }

        if (RightClickRegex.IsMatch(input))
        {
            var match = RightClickRegex.Match(input);
            var target = CleanQuoted(match.Groups["target"].Value);
            if (!string.IsNullOrWhiteSpace(target))
            {
                return (new List<PlanStep>
                {
                    new()
                    {
                        Type = ActionType.RightClick,
                        Selector = new Selector { NameContains = target }
                    }
                }, true);
            }
        }

        if (DragRegex.IsMatch(input))
        {
            var match = DragRegex.Match(input);
            var from = CleanQuoted(match.Groups["from"].Value);
            var to = CleanQuoted(match.Groups["to"].Value);
            if (!string.IsNullOrWhiteSpace(from) && !string.IsNullOrWhiteSpace(to))
            {
                return (new List<PlanStep>
                {
                    new()
                    {
                        Type = ActionType.Drag,
                        Selector = new Selector { NameContains = from },
                        Target = to
                    }
                }, true);
            }
        }

        if (ClipboardHistoryRegex.IsMatch(input))
        {
            return (new List<PlanStep> { new() { Type = ActionType.ClipboardHistory } }, true);
        }

        if (NotifyRegex.IsMatch(input))
        {
            var match = NotifyRegex.Match(input);
            return (new List<PlanStep> { new() { Type = ActionType.Notify, Text = match.Groups["text"].Value.Trim() } }, true);
        }

        if (FileWriteRegex.IsMatch(input))
        {
            var match = FileWriteRegex.Match(input);
            return (new List<PlanStep>
            {
                new()
                {
                    Type = ActionType.FileWrite,
                    Target = CleanQuoted(match.Groups["path"].Value),
                    Text = CleanQuoted(match.Groups["text"].Value)
                }
            }, true);
        }

        if (FileAppendRegex.IsMatch(input))
        {
            var match = FileAppendRegex.Match(input);
            return (new List<PlanStep>
            {
                new()
                {
                    Type = ActionType.FileAppend,
                    Target = CleanQuoted(match.Groups["path"].Value),
                    Text = CleanQuoted(match.Groups["text"].Value)
                }
            }, true);
        }

        if (FileReadRegex.IsMatch(input))
        {
            var match = FileReadRegex.Match(input);
            return (new List<PlanStep> { new() { Type = ActionType.FileRead, Target = CleanQuoted(match.Groups["path"].Value) } }, true);
        }

        if (FileListRegex.IsMatch(input))
        {
            var match = FileListRegex.Match(input);
            var path = CleanQuoted(match.Groups["path"].Value);
            if (string.IsNullOrWhiteSpace(path))
            {
                path = ".";
            }

            return (new List<PlanStep> { new() { Type = ActionType.FileList, Target = path } }, true);
        }

        if (SearchOnBrowserRegex.IsMatch(input))
        {
            var match = SearchOnBrowserRegex.Match(input);
            var query = CleanQuoted(match.Groups["query"].Value);
            var browser = CleanQuoted(match.Groups["browser"].Value);
            if (!string.IsNullOrWhiteSpace(query))
            {
                var appTarget = browser;
                if (_appResolver.TryResolveApp(browser, out var resolvedBrowser))
                {
                    appTarget = resolvedBrowser;
                }

                return (new List<PlanStep>
                {
                    new() { Type = ActionType.OpenApp, AppIdOrPath = appTarget },
                    new() { Type = ActionType.OpenUrl, Target = BuildSearchUrl(query), AppIdOrPath = appTarget }
                }, true);
            }
        }

        if (SearchWebRegex.IsMatch(input))
        {
            var match = SearchWebRegex.Match(input);
            var query = CleanQuoted(match.Groups["query"].Value);
            if (!string.IsNullOrWhiteSpace(query))
            {
                return (new List<PlanStep> { new() { Type = ActionType.OpenUrl, Target = BuildSearchUrl(query) } }, true);
            }
        }

        if (OpenUrlOnBrowserRegex.IsMatch(input))
        {
            var match = OpenUrlOnBrowserRegex.Match(input);
            var browser = CleanQuoted(match.Groups["browser"].Value);
            if (TryNormalizeUrl(match.Groups["url"].Value, out var browserUrl))
            {
                var appTarget = browser;
                if (_appResolver.TryResolveApp(browser, out var resolvedBrowser))
                {
                    appTarget = resolvedBrowser;
                }

                return (new List<PlanStep>
                {
                    new() { Type = ActionType.OpenApp, AppIdOrPath = appTarget },
                    new() { Type = ActionType.OpenUrl, Target = browserUrl, AppIdOrPath = appTarget }
                }, true);
            }
        }

        if (OpenUrlRegex.IsMatch(input))
        {
            var match = OpenUrlRegex.Match(input);
            if (TryNormalizeUrl(match.Groups["url"].Value, out var url))
            {
                return (new List<PlanStep> { new() { Type = ActionType.OpenUrl, Target = url } }, true);
            }
        }

        if (NavigateQueryRegex.IsMatch(input))
        {
            var match = NavigateQueryRegex.Match(input);
            var query = CleanQuoted(match.Groups["query"].Value);
            if (!string.IsNullOrWhiteSpace(query) && !TryNormalizeUrl(query, out _))
            {
                return (new List<PlanStep> { new() { Type = ActionType.OpenUrl, Target = BuildSearchUrl(query) } }, true);
            }
        }

        if (TryNormalizeUrl(input, out var directUrl))
        {
            return (new List<PlanStep> { new() { Type = ActionType.OpenUrl, Target = directUrl } }, true);
        }

        if (FindInPageRegex.IsMatch(input))
        {
            var match = FindInPageRegex.Match(input);
            var searchText = CleanQuoted(match.Groups["text"].Value);
            if (!string.IsNullOrWhiteSpace(searchText))
            {
                return (new List<PlanStep>
                {
                    new() { Type = ActionType.KeyCombo, Keys = GetFindInPageKeys() },
                    new() { Type = ActionType.TypeText, Text = searchText }
                }, true);
            }
        }

        if (SaveAsRegex.IsMatch(input))
        {
            var match = SaveAsRegex.Match(input);
            var file = CleanQuoted(match.Groups["file"].Value);
            var folder = CleanQuoted(match.Groups["folder"].Value);
            var target = BuildSaveAsTarget(file, folder);
            var steps = new List<PlanStep> { new() { Type = ActionType.KeyCombo, Keys = GetSaveAsKeys() } };
            if (!string.IsNullOrWhiteSpace(target))
            {
                steps.Add(new PlanStep { Type = ActionType.TypeText, Text = target });
                steps.Add(new PlanStep { Type = ActionType.KeyCombo, Keys = new List<string> { "enter" } });
            }
            return (steps, true);
        }

        if (SaveAsOnlyRegex.IsMatch(input))
        {
            return (new List<PlanStep> { new() { Type = ActionType.KeyCombo, Keys = GetSaveAsKeys() } }, true);
        }

        if (SelectLineRegex.IsMatch(input))
        {
            return (new List<PlanStep>
            {
                new() { Type = ActionType.KeyCombo, Keys = new List<string> { "home" } },
                new() { Type = ActionType.KeyCombo, Keys = new List<string> { "shift", "end" } }
            }, true);
        }

        if (DeleteLineRegex.IsMatch(input))
        {
            return (new List<PlanStep>
            {
                new() { Type = ActionType.KeyCombo, Keys = new List<string> { "home" } },
                new() { Type = ActionType.KeyCombo, Keys = new List<string> { "shift", "end" } },
                new() { Type = ActionType.KeyCombo, Keys = new List<string> { "delete" } }
            }, true);
        }

        if (ReplaceRegex.IsMatch(input))
        {
            var match = ReplaceRegex.Match(input);
            var from = CleanQuoted(match.Groups["find"].Value);
            var to = CleanQuoted(match.Groups["replace"].Value);
            if (!string.IsNullOrWhiteSpace(from) && !string.IsNullOrWhiteSpace(to))
            {
                return (new List<PlanStep>
                {
                    new() { Type = ActionType.KeyCombo, Keys = new List<string> { "ctrl", "h" } },
                    new() { Type = ActionType.TypeText, Text = from },
                    new() { Type = ActionType.KeyCombo, Keys = new List<string> { "tab" } },
                    new() { Type = ActionType.TypeText, Text = to },
                    new() { Type = ActionType.KeyCombo, Keys = new List<string> { "enter" } }
                }, true);
            }
        }

        if (MinimizeWindowRegex.IsMatch(input))
        {
            return (new List<PlanStep> { new() { Type = ActionType.KeyCombo, Keys = GetMinimizeWindowKeys() } }, true);
        }

        if (MaximizeWindowRegex.IsMatch(input))
        {
            return (new List<PlanStep> { new() { Type = ActionType.KeyCombo, Keys = GetMaximizeWindowKeys() } }, true);
        }

        if (RestoreWindowRegex.IsMatch(input))
        {
            return (new List<PlanStep> { new() { Type = ActionType.KeyCombo, Keys = GetRestoreWindowKeys() } }, true);
        }

        if (SwitchWindowRegex.IsMatch(input))
        {
            return (new List<PlanStep> { new() { Type = ActionType.KeyCombo, Keys = GetSwitchWindowKeys() } }, true);
        }

        if (FocusAppRegex.IsMatch(input))
        {
            var match = FocusAppRegex.Match(input);
            var appToFocus = CleanQuoted(match.Groups["app"].Value);
            if (!string.IsNullOrWhiteSpace(appToFocus))
            {
                var appTarget = appToFocus;
                if (_appResolver.TryResolveApp(appToFocus, out var resolved))
                {
                    appTarget = resolved;
                }

                return (new List<PlanStep> { new() { Type = ActionType.OpenApp, AppIdOrPath = appTarget } }, true);
            }
        }

        if (ScrollRegex.IsMatch(input))
        {
            var match = ScrollRegex.Match(input);
            var direction = match.Groups["dir"].Success ? match.Groups["dir"].Value : match.Groups["dirit"].Value;
            var amountText = match.Groups["amount"].Success ? match.Groups["amount"].Value : match.Groups["amountit"].Value;
            var amount = int.TryParse(amountText, out var parsedAmount) ? Math.Clamp(parsedAmount, 1, 20) : 3;
            var key = direction.Equals("up", StringComparison.OrdinalIgnoreCase) || direction.Equals("su", StringComparison.OrdinalIgnoreCase)
                ? "pageup"
                : "pagedown";

            var steps = new List<PlanStep>();
            for (var i = 0; i < amount; i++)
            {
                steps.Add(new PlanStep { Type = ActionType.KeyCombo, Keys = new List<string> { key } });
            }

            return (steps, true);
        }

        if (PageUpRegex.IsMatch(input))
        {
            return (new List<PlanStep> { new() { Type = ActionType.KeyCombo, Keys = new List<string> { "pageup" } } }, true);
        }

        if (PageDownRegex.IsMatch(input))
        {
            return (new List<PlanStep> { new() { Type = ActionType.KeyCombo, Keys = new List<string> { "pagedown" } } }, true);
        }

        if (HomeRegex.IsMatch(input))
        {
            return (new List<PlanStep> { new() { Type = ActionType.KeyCombo, Keys = new List<string> { "home" } } }, true);
        }

        if (EndRegex.IsMatch(input))
        {
            return (new List<PlanStep> { new() { Type = ActionType.KeyCombo, Keys = new List<string> { "end" } } }, true);
        }

        if (TryMatchActionInApp(input, out var actionPart, out var appPart))
        {
            var (actionSteps, actionRecognized) = InterpretSingle(actionPart);
            if (actionRecognized && actionSteps.Count > 0)
            {
                var appTarget = appPart;
                if (_appResolver.TryResolveApp(appPart, out var resolved))
                {
                    appTarget = resolved;
                }

                var filtered = actionSteps.Where(step => step.Type != ActionType.OpenApp).ToList();
                filtered.Insert(0, new PlanStep { Type = ActionType.OpenApp, AppIdOrPath = appTarget });
                ApplyExpectedAppBinding(filtered, appTarget);
                return (filtered, true);
            }
        }

        if (TryMatch(OpenRegex, input, out var appToOpen))
        {
            var appTarget = appToOpen;
            if (_appResolver.TryResolveApp(appToOpen, out var resolved))
            {
                appTarget = resolved;
            }
            return (new List<PlanStep> { new() { Type = ActionType.OpenApp, AppIdOrPath = appTarget } }, true);
        }

        if (TryMatch(TypeRegex, input, out var text))
        {
            return (new List<PlanStep> { new() { Type = ActionType.TypeText, Text = text } }, true);
        }

        if (TryMatch(ClickRegex, input, out var clickTarget))
        {
            return (new List<PlanStep> { new()
            {
                Type = ActionType.Click,
                Selector = new Selector { NameContains = clickTarget }
            } }, true);
        }

        if (TryMatch(FindRegex, input, out var findValue))
        {
            return (new List<PlanStep> { new()
            {
                Type = ActionType.Find,
                Selector = new Selector { NameContains = findValue }
            } }, true);
        }

        if (TryMatch(PressRegex, input, out var keys))
        {
            if (TryParseKeyCombo(keys, out var combo))
            {
                return (new List<PlanStep> { new() { Type = ActionType.KeyCombo, Keys = combo } }, true);
            }

            return (new List<PlanStep> { new()
            {
                Type = ActionType.Click,
                Selector = new Selector { NameContains = keys }
            } }, true);
        }

        if (NewFileRegex.IsMatch(input))
        {
            return (new List<PlanStep> { new() { Type = ActionType.KeyCombo, Keys = GetNewFileKeys() } }, true);
        }

        if (SaveRegex.IsMatch(input))
        {
            return (new List<PlanStep> { new() { Type = ActionType.KeyCombo, Keys = GetSaveKeys() } }, true);
        }

        if (NewTabRegex.IsMatch(input))
        {
            return (new List<PlanStep> { new() { Type = ActionType.KeyCombo, Keys = GetNewTabKeys() } }, true);
        }

        if (CloseTabRegex.IsMatch(input))
        {
            return (new List<PlanStep> { new() { Type = ActionType.KeyCombo, Keys = GetCloseTabKeys() } }, true);
        }

        if (CloseWindowRegex.IsMatch(input))
        {
            return (new List<PlanStep> { new() { Type = ActionType.KeyCombo, Keys = GetCloseWindowKeys() } }, true);
        }

        if (BrowserBackRegex.IsMatch(input))
        {
            return (new List<PlanStep> { new() { Type = ActionType.KeyCombo, Keys = GetBrowserBackKeys() } }, true);
        }

        if (BrowserForwardRegex.IsMatch(input))
        {
            return (new List<PlanStep> { new() { Type = ActionType.KeyCombo, Keys = GetBrowserForwardKeys() } }, true);
        }

        if (RefreshRegex.IsMatch(input))
        {
            return (new List<PlanStep> { new() { Type = ActionType.KeyCombo, Keys = GetRefreshKeys() } }, true);
        }

        if (CopyRegex.IsMatch(input))
        {
            return (new List<PlanStep> { new() { Type = ActionType.KeyCombo, Keys = GetCopyKeys() } }, true);
        }

        if (PasteRegex.IsMatch(input))
        {
            return (new List<PlanStep> { new() { Type = ActionType.KeyCombo, Keys = GetPasteKeys() } }, true);
        }

        if (UndoRegex.IsMatch(input))
        {
            return (new List<PlanStep> { new() { Type = ActionType.KeyCombo, Keys = GetUndoKeys() } }, true);
        }

        if (RedoRegex.IsMatch(input))
        {
            return (new List<PlanStep> { new() { Type = ActionType.KeyCombo, Keys = GetRedoKeys() } }, true);
        }

        if (SelectAllRegex.IsMatch(input))
        {
            return (new List<PlanStep> { new() { Type = ActionType.KeyCombo, Keys = GetSelectAllKeys() } }, true);
        }

        if (_appResolver.TryResolveApp(input, out var bareApp))
        {
            return (new List<PlanStep> { new() { Type = ActionType.OpenApp, AppIdOrPath = bareApp } }, true);
        }

        return (new List<PlanStep>(), false);
    }

    private static bool TryMatchActionInApp(string input, out string actionPart, out string appPart)
    {
        actionPart = string.Empty;
        appPart = string.Empty;
        var actionInApp = Regex.Match(input, "^(?<action>.+?)\\s+(?:in|nel|nella)\\s+(?<app>.+)$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        if (actionInApp.Success)
        {
            actionPart = actionInApp.Groups["action"].Value.Trim();
            appPart = actionInApp.Groups["app"].Value.Trim();
            return !string.IsNullOrWhiteSpace(actionPart) && !string.IsNullOrWhiteSpace(appPart);
        }

        var inAppAction = Regex.Match(input, "^(?:in|nel|nella)\\s+(?<app>.+?)\\s+(?<action>.+)$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        if (inAppAction.Success)
        {
            appPart = inAppAction.Groups["app"].Value.Trim();
            actionPart = inAppAction.Groups["action"].Value.Trim();
            return !string.IsNullOrWhiteSpace(actionPart) && !string.IsNullOrWhiteSpace(appPart);
        }

        return false;
    }

    private static string BuildSearchUrl(string query)
    {
        var q = query.Trim();
        return $"https://www.google.com/search?q={Uri.EscapeDataString(q)}";
    }

    private static string CleanQuoted(string input)
    {
        var trimmed = input.Trim();
        if (trimmed.Length >= 2 && ((trimmed.StartsWith('"') && trimmed.EndsWith('"')) || (trimmed.StartsWith('\'') && trimmed.EndsWith('\''))))
        {
            return trimmed[1..^1];
        }

        return trimmed;
    }

    private static string BuildSaveAsTarget(string file, string folder)
    {
        if (string.IsNullOrWhiteSpace(file))
        {
            return string.Empty;
        }

        if (string.IsNullOrWhiteSpace(folder))
        {
            return file;
        }

        if (Path.IsPathRooted(file) || file.Contains(":\\", StringComparison.OrdinalIgnoreCase) || file.Contains('/') || file.Contains('\\'))
        {
            return file;
        }

        var combined = Path.Combine(folder, file);
        return combined;
    }

    private static bool TryNormalizeUrl(string input, out string normalized)
    {
        normalized = input.Trim();
        if (Uri.TryCreate(normalized, UriKind.Absolute, out var absolute) && (absolute.Scheme is "http" or "https"))
        {
            normalized = absolute.ToString();
            return true;
        }

        if (normalized.Contains(' ') || !normalized.Contains('.'))
        {
            return false;
        }

        var withScheme = $"https://{normalized}";
        if (Uri.TryCreate(withScheme, UriKind.Absolute, out var withHttps) && !string.IsNullOrWhiteSpace(withHttps.Host))
        {
            normalized = withHttps.ToString();
            return true;
        }

        return false;
    }

    private static bool TryMatch(Regex regex, string input, out string value)
    {
        var match = regex.Match(input);
        if (match.Success)
        {
            value = match.Groups[1].Value.Trim();
            return true;
        }

        value = string.Empty;
        return false;
    }

    private static bool TryParseKeyCombo(string input, out List<string> keys)
    {
        var normalized = input.Replace("+", " ").Replace(",", " ");
        var parts = normalized.Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Where(p => !KeyFillerTokens.Contains(p))
            .ToList();

        if (parts.Count == 0)
        {
            keys = new List<string>();
            return false;
        }

        var allKeys = parts.All(IsKeyToken);
        if (!allKeys)
        {
            keys = parts;
            return false;
        }

        keys = parts;
        return true;
    }

    private static bool IsKeyToken(string token)
    {
        if (KeyTokens.Contains(token))
        {
            return true;
        }

        if (token.Length == 1 && char.IsLetterOrDigit(token[0]))
        {
            return true;
        }

        return Regex.IsMatch(token, "^f\\d{1,2}$", RegexOptions.IgnoreCase);
    }

    private static bool TryParseMouseJiggle(string input, out TimeSpan duration)
    {
        duration = TimeSpan.FromSeconds(30);
        var normalized = (input ?? string.Empty).Trim().ToLowerInvariant();
        if (string.IsNullOrWhiteSpace(normalized))
        {
            return false;
        }

        var hasMouse = normalized.Contains("mouse");
        if (!hasMouse)
        {
            return false;
        }

        var isJiggleIntent =
            normalized.Contains("jiggle mouse")
            || normalized.Contains("mouse jiggle")
            || normalized.Contains("move mouse")
            || normalized.Contains("move the mouse")
            || normalized.Contains("muovi mouse")
            || normalized.Contains("muovi il mouse")
            || normalized.Contains("muovi randomicamente");

        var hasRandomHint = normalized.Contains("random") || normalized.Contains("casual");
        if (!isJiggleIntent && !hasRandomHint)
        {
            return false;
        }

        var match = MouseDurationRegex.Match(normalized);
        if (match.Success
            && double.TryParse(match.Groups["value"].Value.Replace(',', '.'),
                System.Globalization.NumberStyles.Float,
                System.Globalization.CultureInfo.InvariantCulture,
                out var amount))
        {
            var unit = match.Groups["unit"].Value.ToLowerInvariant();
            duration = unit switch
            {
                "h" or "hr" or "hrs" or "hour" or "hours" or "ora" or "ore" => TimeSpan.FromHours(amount),
                "m" or "min" or "mins" or "minute" or "minutes" or "munute" or "munutes" or "minite" or "minites" or "minuto" or "minuti" => TimeSpan.FromMinutes(amount),
                _ => TimeSpan.FromSeconds(amount)
            };
        }
        else
        {
            var bare = Regex.Match(normalized, "(?:for|per)\\s+(?<value>\\d+(?:[\\.,]\\d+)?)", RegexOptions.IgnoreCase);
            if (bare.Success
                && double.TryParse(bare.Groups["value"].Value.Replace(',', '.'),
                    System.Globalization.NumberStyles.Float,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out var seconds))
            {
                duration = TimeSpan.FromSeconds(seconds);
            }
        }

        if (duration <= TimeSpan.Zero)
        {
            duration = TimeSpan.FromSeconds(1);
        }

        if (duration > TimeSpan.FromHours(8))
        {
            duration = TimeSpan.FromHours(8);
        }

        return true;
    }

    private List<string> GetNewFileKeys()
    {
        return ResolveKeys(_config.NewFileKeyCombo, new[] { "ctrl", "n" }, new[] { "cmd", "n" });
    }

    private List<string> GetSaveKeys()
    {
        return ResolveKeys(_config.SaveKeyCombo, new[] { "ctrl", "s" }, new[] { "cmd", "s" });
    }

    private List<string> GetSaveAsKeys()
    {
        return ResolveKeys(_config.SaveAsKeyCombo, new[] { "ctrl", "shift", "s" }, new[] { "cmd", "shift", "s" });
    }

    private List<string> GetNewTabKeys()
    {
        return ResolveKeys(_config.NewTabKeyCombo, new[] { "ctrl", "t" }, new[] { "cmd", "t" });
    }

    private List<string> GetCloseTabKeys()
    {
        return ResolveKeys(_config.CloseTabKeyCombo, new[] { "ctrl", "w" }, new[] { "cmd", "w" });
    }

    private List<string> GetCloseWindowKeys()
    {
        return ResolveKeys(_config.CloseWindowKeyCombo, new[] { "alt", "f4" }, new[] { "cmd", "q" });
    }

    private List<string> GetMinimizeWindowKeys()
    {
        return ResolveKeys(new List<string>(), new[] { "win", "down" }, new[] { "cmd", "m" });
    }

    private List<string> GetMaximizeWindowKeys()
    {
        return ResolveKeys(new List<string>(), new[] { "win", "up" }, new[] { "ctrl", "cmd", "f" });
    }

    private List<string> GetRestoreWindowKeys()
    {
        return ResolveKeys(new List<string>(), new[] { "alt", "space", "r" }, new[] { "ctrl", "cmd", "f" });
    }

    private List<string> GetSwitchWindowKeys()
    {
        return ResolveKeys(new List<string>(), new[] { "alt", "tab" }, new[] { "cmd", "tab" });
    }

    private List<string> GetCopyKeys()
    {
        return ResolveKeys(_config.CopyKeyCombo, new[] { "ctrl", "c" }, new[] { "cmd", "c" });
    }

    private List<string> GetPasteKeys()
    {
        return ResolveKeys(_config.PasteKeyCombo, new[] { "ctrl", "v" }, new[] { "cmd", "v" });
    }

    private List<string> GetUndoKeys()
    {
        return ResolveKeys(_config.UndoKeyCombo, new[] { "ctrl", "z" }, new[] { "cmd", "z" });
    }

    private List<string> GetRedoKeys()
    {
        return ResolveKeys(_config.RedoKeyCombo, new[] { "ctrl", "y" }, new[] { "cmd", "shift", "z" });
    }

    private List<string> GetSelectAllKeys()
    {
        return ResolveKeys(_config.SelectAllKeyCombo, new[] { "ctrl", "a" }, new[] { "cmd", "a" });
    }

    private List<string> GetBrowserBackKeys()
    {
        return ResolveKeys(_config.BrowserBackKeyCombo, new[] { "alt", "left" }, new[] { "cmd", "[" });
    }

    private List<string> GetBrowserForwardKeys()
    {
        return ResolveKeys(_config.BrowserForwardKeyCombo, new[] { "alt", "right" }, new[] { "cmd", "]" });
    }

    private List<string> GetRefreshKeys()
    {
        return ResolveKeys(_config.RefreshKeyCombo, new[] { "ctrl", "r" }, new[] { "cmd", "r" });
    }

    private List<string> GetFindInPageKeys()
    {
        return ResolveKeys(_config.FindInPageKeyCombo, new[] { "ctrl", "f" }, new[] { "cmd", "f" });
    }

    private List<string> ResolveKeys(List<string> configured, string[] windowsDefault, string[] macDefault)
    {
        var isMac = RuntimeInformation.IsOSPlatform(OSPlatform.OSX);
        var isWindowsDefault = SequenceEquals(configured, windowsDefault);

        if (configured.Count == 0)
        {
            return (isMac ? macDefault : windowsDefault).ToList();
        }

        if (isMac && isWindowsDefault)
        {
            return macDefault.ToList();
        }

        return configured.ToList();
    }

    private static bool SequenceEquals(IReadOnlyList<string> values, IReadOnlyList<string> expected)
    {
        if (values.Count != expected.Count)
        {
            return false;
        }

        for (var i = 0; i < values.Count; i++)
        {
            if (!string.Equals(values[i], expected[i], StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }
        }

        return true;
    }

    private static List<PlanStep> CloneSteps(IEnumerable<PlanStep> steps)
    {
        return steps.Select(step => new PlanStep
        {
            Type = step.Type,
            ExpectedAppId = step.ExpectedAppId,
            ExpectedWindowId = step.ExpectedWindowId,
            Text = step.Text,
            Target = step.Target,
            AppIdOrPath = step.AppIdOrPath,
            ElementId = step.ElementId,
            WaitFor = step.WaitFor,
            Note = step.Note,
            Selector = step.Selector == null
                ? null
                : new Selector
                {
                    Role = step.Selector.Role,
                    NameContains = step.Selector.NameContains,
                    AutomationId = step.Selector.AutomationId,
                    ClassName = step.Selector.ClassName,
                    AncestorNameContains = step.Selector.AncestorNameContains,
                    Index = step.Selector.Index,
                    WindowId = step.Selector.WindowId,
                    BoundsHint = step.Selector.BoundsHint == null
                        ? null
                        : new Rect
                        {
                            X = step.Selector.BoundsHint.X,
                            Y = step.Selector.BoundsHint.Y,
                            Width = step.Selector.BoundsHint.Width,
                            Height = step.Selector.BoundsHint.Height
                        }
                },
            Keys = step.Keys == null ? null : new List<string>(step.Keys),
            Point = step.Point == null
                ? null
                : new Rect
                {
                    X = step.Point.X,
                    Y = step.Point.Y,
                    Width = step.Point.Width,
                    Height = step.Point.Height
                }
        }).ToList();
    }

    private static void ApplyExpectedAppBinding(List<PlanStep> steps, string? appId)
    {
        if (string.IsNullOrWhiteSpace(appId))
        {
            return;
        }

        foreach (var step in steps)
        {
            if (step.Type == ActionType.OpenApp)
            {
                continue;
            }

            // Keep browser continuity across chained intents:
            // "open edge and search ..." should use edge, not system default browser.
            if (step.Type == ActionType.OpenUrl && string.IsNullOrWhiteSpace(step.AppIdOrPath))
            {
                step.AppIdOrPath = appId;
            }

            if (!NeedsContextBinding(step.Type))
            {
                continue;
            }

            if (string.IsNullOrWhiteSpace(step.ExpectedAppId))
            {
                step.ExpectedAppId = appId;
            }
        }
    }

    private static bool NeedsContextBinding(ActionType type)
    {
        return type is ActionType.Find
            or ActionType.Click
            or ActionType.DoubleClick
            or ActionType.RightClick
            or ActionType.Drag
            or ActionType.TypeText
            or ActionType.KeyCombo
            or ActionType.Invoke
            or ActionType.SetValue
            or ActionType.ReadText
            or ActionType.CaptureScreen;
    }
}
