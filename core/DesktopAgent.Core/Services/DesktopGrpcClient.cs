using DesktopAgent.Core.Abstractions;
using DesktopAgent.Proto;
using Grpc.Net.Client;
using Microsoft.Extensions.Logging;
using System.Net.Http;

namespace DesktopAgent.Core.Services;

public sealed class DesktopGrpcClient : IDesktopAdapterClient, IDisposable
{
    private readonly GrpcChannel _channel;
    private readonly DesktopAdapter.DesktopAdapterClient _client;
    private readonly ILogger<DesktopGrpcClient> _logger;

    public DesktopGrpcClient(string endpoint, ILogger<DesktopGrpcClient> logger)
    {
        _logger = logger;
        // Allow h2c (HTTP/2 without TLS) for local adapter connections.
        AppContext.SetSwitch("System.Net.Http.SocketsHttpHandler.Http2UnencryptedSupport", true);
        var httpHandler = new SocketsHttpHandler
        {
            EnableMultipleHttp2Connections = true
        };
        _channel = GrpcChannel.ForAddress(endpoint, new GrpcChannelOptions
        {
            HttpHandler = httpHandler
        });
        _client = new DesktopAdapter.DesktopAdapterClient(_channel);
    }

    public async Task<WindowRef?> GetActiveWindowAsync(CancellationToken cancellationToken)
    {
        try
        {
            var result = await _client.GetActiveWindowAsync(new Empty(), cancellationToken: cancellationToken);
            return string.IsNullOrWhiteSpace(result.Id) ? null : result;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "GetActiveWindow failed");
            return null;
        }
    }

    public async Task<IReadOnlyList<WindowRef>> ListWindowsAsync(CancellationToken cancellationToken)
    {
        try
        {
            var result = await _client.ListWindowsAsync(new Empty(), cancellationToken: cancellationToken);
            return result.Windows;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "ListWindows failed");
            return Array.Empty<WindowRef>();
        }
    }

    public async Task<UiTree?> GetUiTreeAsync(string windowId, CancellationToken cancellationToken)
    {
        try
        {
            var result = await _client.GetUiTreeAsync(new WindowRequest { WindowId = windowId }, cancellationToken: cancellationToken);
            return result.Root == null ? null : result;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "GetUiTree failed");
            return null;
        }
    }

    public async Task<IReadOnlyList<ElementRef>> FindElementsAsync(Selector selector, CancellationToken cancellationToken)
    {
        try
        {
            var result = await _client.FindElementsAsync(new FindRequest { Selector = selector }, cancellationToken: cancellationToken);
            return result.Elements;
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "FindElements failed");
            return Array.Empty<ElementRef>();
        }
    }

    public async Task<ActionResult> InvokeElementAsync(string elementId, CancellationToken cancellationToken)
    {
        try
        {
            return await _client.InvokeElementAsync(new ElementRequest { ElementId = elementId }, cancellationToken: cancellationToken);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "InvokeElement failed");
            return new ActionResult { Success = false, Message = ex.Message };
        }
    }

    public async Task<ActionResult> SetElementValueAsync(string elementId, string value, CancellationToken cancellationToken)
    {
        try
        {
            return await _client.SetElementValueAsync(new SetValueRequest { ElementId = elementId, Value = value }, cancellationToken: cancellationToken);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "SetElementValue failed");
            return new ActionResult { Success = false, Message = ex.Message };
        }
    }

    public async Task<ActionResult> ClickPointAsync(int x, int y, CancellationToken cancellationToken)
    {
        try
        {
            return await _client.ClickPointAsync(new ClickRequest { X = x, Y = y }, cancellationToken: cancellationToken);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "ClickPoint failed");
            return new ActionResult { Success = false, Message = ex.Message };
        }
    }

    public async Task<ActionResult> TypeTextAsync(string text, CancellationToken cancellationToken)
    {
        try
        {
            return await _client.TypeTextAsync(new TypeTextRequest { Text = text }, cancellationToken: cancellationToken);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "TypeText failed");
            return new ActionResult { Success = false, Message = ex.Message };
        }
    }

    public async Task<ActionResult> KeyComboAsync(IEnumerable<string> keys, CancellationToken cancellationToken)
    {
        try
        {
            var request = new KeyComboRequest();
            request.Keys.AddRange(keys);
            return await _client.KeyComboAsync(request, cancellationToken: cancellationToken);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "KeyCombo failed");
            return new ActionResult { Success = false, Message = ex.Message };
        }
    }

    public async Task<ActionResult> OpenAppAsync(string appIdOrPath, CancellationToken cancellationToken)
    {
        try
        {
            return await _client.OpenAppAsync(new OpenAppRequest { AppIdOrPath = appIdOrPath }, cancellationToken: cancellationToken);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "OpenApp failed");
            return new ActionResult { Success = false, Message = ex.Message };
        }
    }

    public async Task<ScreenshotResponse> CaptureScreenAsync(ScreenshotRequest request, CancellationToken cancellationToken)
    {
        try
        {
            return await _client.CaptureScreenAsync(request, cancellationToken: cancellationToken);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "CaptureScreen failed");
            return new ScreenshotResponse();
        }
    }

    public async Task<ClipboardResponse> GetClipboardAsync(CancellationToken cancellationToken)
    {
        try
        {
            return await _client.GetClipboardAsync(new Empty(), cancellationToken: cancellationToken);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "GetClipboard failed");
            return new ClipboardResponse { Text = string.Empty };
        }
    }

    public async Task<ActionResult> SetClipboardAsync(string text, CancellationToken cancellationToken)
    {
        try
        {
            return await _client.SetClipboardAsync(new SetClipboardRequest { Text = text }, cancellationToken: cancellationToken);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "SetClipboard failed");
            return new ActionResult { Success = false, Message = ex.Message };
        }
    }

    public async Task<Status> ArmAsync(bool requireUserPresence, CancellationToken cancellationToken)
    {
        try
        {
            return await _client.ArmAsync(new ArmRequest { RequireUserPresence = requireUserPresence }, cancellationToken: cancellationToken);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Arm failed");
            return new Status { Armed = false, RequireUserPresence = false, Message = ex.Message };
        }
    }

    public async Task<Status> DisarmAsync(CancellationToken cancellationToken)
    {
        try
        {
            return await _client.DisarmAsync(new Empty(), cancellationToken: cancellationToken);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Disarm failed");
            return new Status { Armed = false, RequireUserPresence = false, Message = ex.Message };
        }
    }

    public async Task<Status> GetStatusAsync(CancellationToken cancellationToken)
    {
        try
        {
            return await _client.GetStatusAsync(new Empty(), cancellationToken: cancellationToken);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "GetStatus failed");
            return new Status { Armed = false, RequireUserPresence = false, Message = ex.Message };
        }
    }

    public void Dispose()
    {
        _channel.Dispose();
    }
}
