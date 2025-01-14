using System;
using SherpaOnnx;

namespace SherpaNative
{
    public class SherpaWrapper
    {
        private OfflineTts _tts;

        public SherpaWrapper(string modelPath, string tokensPath)
        {
            var config = new OfflineTtsConfig();
            config.Model.Vits.Model = modelPath;
            config.Model.Vits.Tokens = tokensPath;
            config.Model.Vits.NoiseScale = 0.667f;
            config.Model.Vits.NoiseScaleW = 0.8f;
            config.Model.Vits.LengthScale = 1.0f;
            config.Model.NumThreads = 1;
            config.Model.Debug = 0;
            config.Model.Provider = "cpu";

            _tts = new OfflineTts(config);
        }

        public byte[] GenerateWaveform(string text)
        {
            var audio = _tts.Generate(text, 1.0f, 0);
            var samples = audio.Samples;
            var bytes = new byte[samples.Length * 2];
            Buffer.BlockCopy(samples, 0, bytes, 0, bytes.Length);
            return bytes;
        }

        public int SampleRate => _tts.SampleRate;

        public void Dispose()
        {
            _tts?.Dispose();
        }
    }
}
