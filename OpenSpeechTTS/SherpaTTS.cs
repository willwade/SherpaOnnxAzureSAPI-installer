using System;
using System.Linq;
using SherpaOnnx; // Reference the Sherpa ONNX C# API

public class SherpaTTS
{
    private OfflineTts _tts;

    public SherpaTTS(string modelPath, string tokensPath, string lexiconPath = "")
    {
        // Configure Sherpa ONNX with the provided model
        var vitsConfig = new OfflineTtsVitsModelConfig(
            modelPath,
            lexiconPath,
            tokensPath,
            dataDir: "",
            dictDir: "",
            noiseScale: 0.667,
            noiseScaleW: 0.8,
            lengthScale: 1.0
        );

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

    
    public byte[] GenerateAudio(string text)
    {
        var audioData = _tts.Generate(text, speed: 1.0f, speakerId: 0);
        if (audioData == null || audioData.Samples.Length == 0)
            throw new Exception("Failed to generate audio.");

        return audioData.Samples.SelectMany(sample =>
            BitConverter.GetBytes((short)(sample * short.MaxValue))).ToArray();
    }


}
