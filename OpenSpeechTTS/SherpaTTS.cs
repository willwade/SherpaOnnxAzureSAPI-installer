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

        var modelConfig = new OfflineTtsModelConfig { Vits = vitsConfig };
        var config = new OfflineTtsConfig(modelConfig, maxNumSentences: 2);

        _tts = new OfflineTts(config);
    }

    public float[] GenerateAudio(string text)
    {
        var audioData = _tts.Generate(text);

        if (audioData == null || audioData.Samples.Length == 0)
            throw new Exception("Failed to generate audio.");

        // Normalize audio data
        float maxAmplitude = Math.Abs(audioData.Samples.Max());
        return audioData.Samples.Select(sample => sample / maxAmplitude).ToArray();
    }

}
