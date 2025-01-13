using System;
using System.IO;
using SherpaOnnx;

namespace OpenSpeechTTS
{
    public class SherpaTTS : IDisposable
    {
        private readonly OfflineTts _tts;

        public SherpaTTS(string modelPath, string tokensPath, string lexiconPath = "", string dataDirPath = "")
        {
            try
            {
                var config = new OfflineTtsConfig();
                config.Model.Vits.Model = modelPath;
                config.Model.Vits.Tokens = tokensPath;
                config.Model.Vits.Lexicon = lexiconPath;
                config.Model.Vits.DataDir = dataDirPath;
                config.Model.Vits.NoiseScale = 0.667f;
                config.Model.Vits.NoiseScaleW = 0.8f;
                config.Model.Vits.LengthScale = 1.0f;
                config.Model.NumThreads = 1;
                config.Model.Debug = 0;
                config.Model.Provider = "cpu";

                _tts = new OfflineTts(config);
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to initialize SherpaTTS: {ex.Message}", ex);
            }
        }

        public byte[] GenerateAudio(string text, float speed = 1.0f, int speakerId = 0)
        {
            try
            {
                var audio = _tts.Generate(text, speed, speakerId);
                
                // Convert float samples to 16-bit PCM
                var pcmSamples = new byte[audio.Samples.Length * 2];
                for (int i = 0; i < audio.Samples.Length; i++)
                {
                    var sample = (short)(audio.Samples[i] * short.MaxValue);
                    var bytes = BitConverter.GetBytes(sample);
                    pcmSamples[i * 2] = bytes[0];
                    pcmSamples[i * 2 + 1] = bytes[1];
                }

                return pcmSamples;
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to generate audio: {ex.Message}", ex);
            }
        }

        public int SampleRate => _tts.SampleRate;

        public void Dispose()
        {
            _tts?.Dispose();
        }
    }
}
