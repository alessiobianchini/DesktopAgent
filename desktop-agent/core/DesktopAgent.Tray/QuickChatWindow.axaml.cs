using System.Diagnostics;
using System.Text.RegularExpressions;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Threading;

namespace DesktopAgent.Tray;

internal partial class QuickChatWindow : Window
{
    private readonly WebApiClient _apiClient;
    private readonly string _webUiUrl;

    private TextBlock? _historyBox;
    private ScrollViewer? _historyScroll;
    private TextBox? _inputBox;
    private TextBlock? _statusText;
    private StackPanel? _confirmPanel;
    private TextBlock? _confirmText;
    private Button? _sendButton;
    private Button? _confirmButton;
    private Button? _cancelButton;
    private Button? _statusButton;
    private Button? _armButton;
    private Button? _disarmButton;
    private Button? _openWebButton;
    private Button? _copyButton;
    private Button? _clearButton;

    private string? _pendingToken;
    private bool _busy;
    private readonly Queue<string> _historyLines = new();
    private const int MaxHistoryLines = 300;

    public QuickChatWindow(WebApiClient apiClient, string webUiUrl)
    {
        _apiClient = apiClient;
        _webUiUrl = webUiUrl;
        InitializeComponent();
        WireControls();
    }

    private void WireControls()
    {
        _historyBox = this.FindControl<TextBlock>("HistoryBox");
        _historyScroll = this.FindControl<ScrollViewer>("HistoryScroll");
        _inputBox = this.FindControl<TextBox>("InputBox");
        _statusText = this.FindControl<TextBlock>("StatusText");
        _confirmPanel = this.FindControl<StackPanel>("ConfirmPanel");
        _confirmText = this.FindControl<TextBlock>("ConfirmText");
        _sendButton = this.FindControl<Button>("SendButton");
        _confirmButton = this.FindControl<Button>("ConfirmButton");
        _cancelButton = this.FindControl<Button>("CancelButton");
        _statusButton = this.FindControl<Button>("StatusButton");
        _armButton = this.FindControl<Button>("ArmButton");
        _disarmButton = this.FindControl<Button>("DisarmButton");
        _openWebButton = this.FindControl<Button>("OpenWebButton");
        _copyButton = this.FindControl<Button>("CopyButton");
        _clearButton = this.FindControl<Button>("ClearButton");

        if (_sendButton != null)
        {
            _sendButton.Click += async (_, _) => await SendFromInputAsync();
        }

        if (_inputBox != null)
        {
            _inputBox.KeyDown += async (_, e) =>
            {
                if (e.Key == Key.Enter)
                {
                    e.Handled = true;
                    await SendFromInputAsync();
                }
            };
        }

        if (_confirmButton != null)
        {
            _confirmButton.Click += async (_, _) => await ConfirmAsync(true);
        }

        if (_cancelButton != null)
        {
            _cancelButton.Click += async (_, _) => await ConfirmAsync(false);
        }

        if (_statusButton != null)
        {
            _statusButton.Click += async (_, _) => await SendMessageAsync("status");
        }

        if (_armButton != null)
        {
            _armButton.Click += async (_, _) => await SendMessageAsync("arm");
        }

        if (_disarmButton != null)
        {
            _disarmButton.Click += async (_, _) => await SendMessageAsync("disarm");
        }

        if (_openWebButton != null)
        {
            _openWebButton.Click += (_, _) => OpenWebUi();
        }

        if (_copyButton != null)
        {
            _copyButton.Click += async (_, _) => await CopyHistoryAsync();
        }

        if (_clearButton != null)
        {
            _clearButton.Click += (_, _) =>
            {
                _historyLines.Clear();
                if (_historyBox != null)
                {
                    _historyBox.Text = string.Empty;
                }
            };
        }

        Opened += OnOpened;
    }

    private async void OnOpened(object? sender, EventArgs e)
    {
        AppendSystem("Quick chat ready.");
        await RefreshStatusLineAsync();
    }

    private async Task SendFromInputAsync()
    {
        if (_inputBox == null)
        {
            return;
        }

        var text = _inputBox.Text?.Trim() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(text))
        {
            return;
        }

        text = SanitizeInputForCommand(text);
        if (string.IsNullOrWhiteSpace(text))
        {
            return;
        }

        _inputBox.Text = string.Empty;
        await SendMessageAsync(text);
    }

    private async Task SendMessageAsync(string message)
    {
        if (_busy)
        {
            return;
        }

        SetBusy(true);
        AppendUser(message);
        try
        {
            var response = await _apiClient.SendChatAsync(message, CancellationToken.None);
            RenderResponse(response);
        }
        catch (Exception ex)
        {
            AppendSystem($"Error: {ex.Message}");
        }
        finally
        {
            SetBusy(false);
            await RefreshStatusLineAsync();
        }
    }

    private async Task ConfirmAsync(bool approve)
    {
        if (_busy || string.IsNullOrWhiteSpace(_pendingToken))
        {
            return;
        }

        SetBusy(true);
        try
        {
            var response = await _apiClient.ConfirmAsync(_pendingToken, approve, CancellationToken.None);
            RenderResponse(response);
        }
        catch (Exception ex)
        {
            AppendSystem($"Error: {ex.Message}");
        }
        finally
        {
            _pendingToken = null;
            ShowConfirm(false, null);
            SetBusy(false);
            await RefreshStatusLineAsync();
        }
    }

    private void RenderResponse(WebChatResponse response)
    {
        var reply = string.IsNullOrWhiteSpace(response.Reply) ? "<no reply>" : response.Reply;
        AppendAgent(reply);

        if (!string.IsNullOrWhiteSpace(response.ModeLabel))
        {
            AppendSystem(response.ModeLabel!);
        }

        if (response.Steps is { Count: > 0 })
        {
            foreach (var step in response.Steps)
            {
                AppendSystem(step);
            }
        }

        if (response.NeedsConfirmation && !string.IsNullOrWhiteSpace(response.Token))
        {
            _pendingToken = response.Token;
            var prompt = string.IsNullOrWhiteSpace(response.ActionLabel)
                ? "Confirmation required."
                : $"{response.ActionLabel} required.";
            ShowConfirm(true, prompt);
        }
        else
        {
            _pendingToken = null;
            ShowConfirm(false, null);
        }
    }

    private async Task RefreshStatusLineAsync()
    {
        try
        {
            var line = await _apiClient.GetStatusLineAsync(CancellationToken.None);
            Dispatcher.UIThread.Post(() =>
            {
                if (_statusText != null)
                {
                    _statusText.Text = line;
                }
            });
        }
        catch
        {
            Dispatcher.UIThread.Post(() =>
            {
                if (_statusText != null)
                {
                    _statusText.Text = "status unavailable";
                }
            });
        }
    }

    private void ShowConfirm(bool show, string? text)
    {
        Dispatcher.UIThread.Post(() =>
        {
            if (_confirmPanel != null)
            {
                _confirmPanel.IsVisible = show;
            }

            if (_confirmText != null && !string.IsNullOrWhiteSpace(text))
            {
                _confirmText.Text = text;
            }
        });
    }

    private void SetBusy(bool busy)
    {
        _busy = busy;
        Dispatcher.UIThread.Post(() =>
        {
            if (_sendButton != null) _sendButton.IsEnabled = !busy;
            if (_confirmButton != null) _confirmButton.IsEnabled = !busy;
            if (_cancelButton != null) _cancelButton.IsEnabled = !busy;
            if (_statusButton != null) _statusButton.IsEnabled = !busy;
            if (_armButton != null) _armButton.IsEnabled = !busy;
            if (_disarmButton != null) _disarmButton.IsEnabled = !busy;
            if (_copyButton != null) _copyButton.IsEnabled = !busy;
            if (_clearButton != null) _clearButton.IsEnabled = !busy;
        });
    }

    private void AppendUser(string text) => AppendLine("YOU", text);
    private void AppendAgent(string text) => AppendLine("AGENT", text);
    private void AppendSystem(string text) => AppendLine("SYSTEM", text);

    private async Task CopyHistoryAsync()
    {
        var value = _historyBox?.Text ?? string.Empty;
        if (string.IsNullOrWhiteSpace(value))
        {
            AppendSystem("Nothing to copy.");
            return;
        }

        try
        {
            var clipboard = TopLevel.GetTopLevel(this)?.Clipboard;
            if (clipboard == null)
            {
                AppendSystem("Clipboard unavailable.");
                return;
            }

            await clipboard.SetTextAsync(value);
            AppendSystem("Conversation copied.");
        }
        catch (Exception ex)
        {
            AppendSystem($"Copy failed: {ex.Message}");
        }
    }

    private void AppendLine(string role, string text)
    {
        Dispatcher.UIThread.Post(() =>
        {
            if (_historyBox == null)
            {
                return;
            }

            var timestamp = DateTime.Now.ToString("HH:mm:ss");
            var cleanText = text.Replace("\r", string.Empty).Replace('\n', ' ').Trim();
            var line = $"[{timestamp}] {role,-6} {cleanText}";
            _historyLines.Enqueue(line);
            while (_historyLines.Count > MaxHistoryLines)
            {
                _historyLines.Dequeue();
            }

            _historyBox.Text = string.Join(Environment.NewLine, _historyLines);
            if (_historyScroll != null)
            {
                _historyScroll.Offset = new Vector(_historyScroll.Offset.X, double.MaxValue);
            }
        });
    }

    private static string SanitizeInputForCommand(string input)
    {
        var value = input.Trim();
        for (var i = 0; i < 3; i++)
        {
            var cleaned = Regex.Replace(value, "^(you|agent|system)\\s*:\\s*", string.Empty, RegexOptions.IgnoreCase).Trim();
            if (string.Equals(cleaned, value, StringComparison.Ordinal))
            {
                break;
            }

            value = cleaned;
        }

        return value;
    }

    private void OpenWebUi()
    {
        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = _webUiUrl,
                UseShellExecute = true
            });
        }
        catch
        {
            // Ignore shell launch errors.
        }
    }
}
