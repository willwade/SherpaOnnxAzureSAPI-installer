using NAudio.Wave;
using System;
using System.IO;
using System.Runtime.InteropServices;
using SherpaOnnx;

namespace OpenSpeechTTS
{
    public class SherpaTTS : IDisposable
    {
        private readonly OfflineTts _tts;
        private bool _disposed;

        public SherpaTTS(string modelPath, string tokensPath, string lexiconPath, string dataDirPath)
        {
            if (!File.Exists(modelPath))
                throw new FileNotFoundException($"Model file not found: {modelPath}");

            if (!File.Exists(tokensPath))
                throw new FileNotFoundException($"Tokens file not found: {tokensPath}");

            if (!Directory.Exists(dataDirPath))
                throw new DirectoryNotFoundException($"Data directory not found: {dataDirPath}");

            // Log the paths for debugging
            try
            {
                File.AppendAllText("C:\\OpenSpeech\\sherpa_init.log", 
                    $"{DateTime.Now}: Initializing SherpaTTS\nModel: {modelPath}\nTokens: {tokensPath}\nData Dir: {dataDirPath}\n\n");
            }
            catch { }

            // Initialize the Sherpa ONNX TTS engine
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

        public void SpeakToWaveStream(string text, Stream stream)
        {
            if (_disposed)
                throw new ObjectDisposedException(nameof(SherpaTTS));

            try
            {
                // Generate audio using Sherpa ONNX
                var audio = _tts.Generate(text, 1.0f, 0);
                var samples = audio.Samples;
                
                // Convert float samples to bytes (16-bit PCM)
                byte[] bytes = new byte[samples.Length * 2];
                for (int i = 0; i < samples.Length; i++)
                {
                    short pcm = (short)(samples[i] * short.MaxValue);
                    bytes[i * 2] = (byte)(pcm & 0xFF);
                    bytes[i * 2 + 1] = (byte)((pcm >> 8) & 0xFF);
                }

                using (var writer = new BinaryWriter(stream))
                {
                    writer.Write(0x46464952); // "RIFF"
                    writer.Write(36 + bytes.Length);
                    writer.Write(0x45564157); // "WAVE"
                    writer.Write(0x20746D66); // "fmt "
                    writer.Write(16);
                    writer.Write((short)1); // PCM
                    writer.Write((short)1); // Mono
                    writer.Write(_tts.SampleRate); // Sample rate
                    writer.Write(_tts.SampleRate * 2); // Bytes per second
                    writer.Write((short)2); // Block align
                    writer.Write((short)16); // Bits per sample
                    writer.Write(0x61746164); // "data"
                    writer.Write(bytes.Length);
                    writer.Write(bytes);
                }
            }
            catch (Exception ex)
            {
                // Log the error for debugging
                try
                {
                    File.AppendAllText("C:\\OpenSpeech\\sherpa_error.log", 
                        $"{DateTime.Now}: Error in SpeakToWaveStream: {ex.Message}\nText: {text}\n{ex.StackTrace}\n\n");
                }
                catch { }
                
                throw;
            }
        }

        // Add a method to generate audio bytes directly
        public byte[] GenerateAudio(string text)
        {
            if (_disposed)
                throw new ObjectDisposedException(nameof(SherpaTTS));

            try
            {
                // Generate audio using Sherpa ONNX
                var audio = _tts.Generate(text, 1.0f, 0);
                var samples = audio.Samples;
                
                // Convert float samples to bytes (16-bit PCM)
                byte[] bytes = new byte[samples.Length * 2];
                for (int i = 0; i < samples.Length; i++)
                {
                    short pcm = (short)(samples[i] * short.MaxValue);
                    bytes[i * 2] = (byte)(pcm & 0xFF);
                    bytes[i * 2 + 1] = (byte)((pcm >> 8) & 0xFF);
                }
                
                // Create a WAV file in memory
                using (var ms = new MemoryStream())
                {
                    using (var writer = new BinaryWriter(ms))
                    {
                        writer.Write(0x46464952); // "RIFF"
                        writer.Write(36 + bytes.Length);
                        writer.Write(0x45564157); // "WAVE"
                        writer.Write(0x20746D66); // "fmt "
                        writer.Write(16);
                        writer.Write((short)1); // PCM
                        writer.Write((short)1); // Mono
                        writer.Write(_tts.SampleRate); // Sample rate
                        writer.Write(_tts.SampleRate * 2); // Bytes per second
                        writer.Write((short)2); // Block align
                        writer.Write((short)16); // Bits per sample
                        writer.Write(0x61746164); // "data"
                        writer.Write(bytes.Length);
                        writer.Write(bytes);
                    }
                    
                    return ms.ToArray();
                }
            }
            catch (Exception ex)
            {
                // Log the error for debugging
                try
                {
                    File.AppendAllText("C:\\OpenSpeech\\sherpa_error.log", 
                        $"{DateTime.Now}: Error in GenerateAudio: {ex.Message}\nText: {text}\n{ex.StackTrace}\n\n");
                }
                catch { }
                
                throw;
            }
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                _tts?.Dispose();
                _disposed = true;
            }
        }
    }
}
