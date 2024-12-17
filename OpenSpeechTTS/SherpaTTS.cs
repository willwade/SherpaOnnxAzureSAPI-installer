using System;
using System.Linq;
using SherpaOnnx;

public class SherpaTTS
{
    private OfflineTts _tts;

    public SherpaTTS(string modelPath, string tokensPath, string lexiconPath = "")
    {
        // Correct property-based initialization
        var vitsConfig = new OfflineTtsVitsModelConfig
        {
            Model = modelPath,
            Lexicon = lexiconPath,
            Tokens = tokensPath,
            NoiseScale = 0.667f,
            NoiseScaleW = 0.8f,
            LengthScale = 1.0f
        };

        var modelConfig = new OfflineTtsModelConfig { Vits = vitsConfig };
        var config = new OfflineTtsConfig { Model = modelConfig };

        _tts = new OfflineTts(config);
    }

    public byte[] GenerateAudio(string text, float speed = 1.0f, int speakerId = 0)
    {
    var audioData = _tts.Generate(text, speed, speakerId);

    if (audioData == null || audioData.Samples.Length == 0)
        throw new Exception("Failed to generate audio.");

    // Convert float samples to 16-bit PCM byte array
    return audioData.Samples
        .SelectMany(sample => BitConverter.GetBytes((short)(sample * short.MaxValue)))
        .ToArray();
    }
}
