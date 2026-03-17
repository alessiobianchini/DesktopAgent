using DesktopAgent.Core.Config;
using DesktopAgent.Core.Services;
using Microsoft.Extensions.Logging.Abstractions;
using Xunit;

namespace DesktopAgent.Tests;

public sealed class OcrEngineFactoryTests
{
    [Fact]
    public void Create_WhenOcrDisabled_ReturnsStub()
    {
        var config = new AgentConfig { OcrEnabled = false };
        var engine = OcrEngineFactory.Create(config, NullLoggerFactory.Instance);
        Assert.Equal("stub", engine.Name);
    }

    [Fact]
    public void Create_WhenAutoAndLlmEnabled_ReturnsChain()
    {
        var config = new AgentConfig
        {
            OcrEnabled = true,
            Ocr = new OcrConfig { Engine = "auto", TesseractPath = "tesseract" },
            LlmFallbackEnabled = true,
            LlmFallback = new LlmFallbackConfig
            {
                Provider = "ollama",
                Endpoint = "http://localhost:11434/api/generate",
                Model = "llava"
            }
        };

        var engine = OcrEngineFactory.Create(config, NullLoggerFactory.Instance);
        Assert.Contains("chain", engine.Name, StringComparison.OrdinalIgnoreCase);
    }
}

