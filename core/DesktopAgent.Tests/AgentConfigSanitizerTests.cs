using DesktopAgent.Core.Config;
using Xunit;

namespace DesktopAgent.Tests;

public sealed class AgentConfigSanitizerTests
{
    [Theory]
    [InlineData("wasapi", "wasapi")]
    [InlineData("dshow", "dshow")]
    [InlineData("avfoundation", "avfoundation")]
    [InlineData("pipewire", "pipewire")]
    [InlineData("pulse", "pulse")]
    [InlineData("alsa", "alsa")]
    [InlineData("AUTO", "auto")]
    [InlineData("unknown", "auto")]
    public void Normalize_AudioBackendPreference_IsSanitized(string input, string expected)
    {
        var config = new AgentConfig
        {
            ScreenRecordingAudioBackendPreference = input
        };

        AgentConfigSanitizer.Normalize(config);

        Assert.Equal(expected, config.ScreenRecordingAudioBackendPreference);
    }

    [Theory]
    [InlineData("tesseract", "tesseract")]
    [InlineData("ai", "ai")]
    [InlineData("AUTO", "auto")]
    [InlineData("other", "auto")]
    public void Normalize_OcrEngine_IsSanitized(string input, string expected)
    {
        var config = new AgentConfig
        {
            Ocr = new OcrConfig
            {
                Engine = input,
                TesseractPath = ""
            }
        };

        AgentConfigSanitizer.Normalize(config);

        Assert.Equal(expected, config.Ocr.Engine);
        Assert.Equal("tesseract", config.Ocr.TesseractPath);
    }
}
