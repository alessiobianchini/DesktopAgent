using System.Net;
using DesktopAgent.Adapter.Windows;
using Microsoft.AspNetCore.Server.Kestrel.Core;

var builder = WebApplication.CreateBuilder(args);

var port = Environment.GetEnvironmentVariable("DESKTOP_AGENT_PORT") ?? "50051";
builder.WebHost.ConfigureKestrel(options =>
{
    options.ListenAnyIP(int.Parse(port), listenOptions =>
    {
        listenOptions.Protocols = HttpProtocols.Http2;
    });
});

builder.Services.AddGrpc();
builder.Services.AddSingleton<AdapterState>();

var app = builder.Build();

app.MapGrpcService<DesktopAdapterService>();
app.MapGet("/", () => "DesktopAgent.Adapter.Windows gRPC server");

app.Run();
