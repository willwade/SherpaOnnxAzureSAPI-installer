using NAudio.Wave;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Reflection;
// Note: SherpaOnnx namespace will be loaded dynamically to avoid strong-name issues

namespace OpenSpeechTTS
{
    public class SherpaTTS : IDisposable
    {
        private readonly OfflineTts _tts;
        private bool _disposed;
        private readonly string _logDir = "C:\\OpenSpeech";

        public SherpaTTS(string modelPath, string tokensPath, string lexiconPath, string dataDirPath)
        {
            try
            {
                // Create log directory if it doesn't exist
                if (!Directory.Exists(_logDir))
                {
                    Directory.CreateDirectory(_logDir);
                }

                LogMessage("Initializing SherpaTTS");
                LogMessage($"Model Path: {modelPath}");
                LogMessage($"Tokens Path: {tokensPath}");
                LogMessage($"Data Directory: {dataDirPath}");

                // Check for assembly version
                var assembly = Assembly.GetAssembly(typeof(OfflineTts));
                LogMessage($"Using sherpa-onnx assembly version: {assembly.GetName().Version}");

                if (!File.Exists(modelPath))
                {
                    LogError($"Model file not found: {modelPath}");
                    throw new FileNotFoundException($"Model file not found: {modelPath}");
                }

                if (!File.Exists(tokensPath))
                {
                    LogError($"Tokens file not found: {tokensPath}");
                    throw new FileNotFoundException($"Tokens file not found: {tokensPath}");
                }

                if (!Directory.Exists(dataDirPath))
                {
                    LogError($"Data directory not found: {dataDirPath}");
                    throw new DirectoryNotFoundException($"Data directory not found: {dataDirPath}");
                }

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

                LogMessage("Creating OfflineTts instance with config");
                _tts = new OfflineTts(config);
                LogMessage("SherpaTTS initialized successfully");
            }
            catch (Exception ex)
            {
                LogError("Error initializing SherpaTTS", ex);
                throw;
            }
        }

        public void SpeakToWaveStream(string text, Stream stream)
        {
            if (_disposed)
                throw new ObjectDisposedException(nameof(SherpaTTS));

            try
            {
                LogMessage($"Generating speech for text: '{text}'");

                // Generate audio using Sherpa ONNX
                var audio = _tts.Generate(text, 1.0f, 0);
                var samples = audio.Samples;

                LogMessage($"Generated {samples.Length} audio samples at {_tts.SampleRate}Hz");

                // Convert float samples to bytes (16-bit PCM)
                byte[] bytes = new byte[samples.Length * 2];
                for (int i = 0; i < samples.Length; i++)
                {
                    short pcm = (short)(samples[i] * short.MaxValue);
                    bytes[i * 2] = (byte)(pcm & 0xFF);
                    bytes[i * 2 + 1] = (byte)((pcm >> 8) & 0xFF);
                }

                LogMessage($"Writing {bytes.Length} bytes to WAV stream");

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

                LogMessage("Successfully wrote WAV data to stream");
            }
            catch (Exception ex)
            {
                LogError($"Error in SpeakToWaveStream", ex);
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
                LogMessage($"Generating audio bytes for text: '{text}'");

                // Generate audio using Sherpa ONNX
                var audio = _tts.Generate(text, 1.0f, 0);
                var samples = audio.Samples;

                LogMessage($"Generated {samples.Length} audio samples at {_tts.SampleRate}Hz");

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

                    LogMessage($"Successfully created WAV data ({ms.Length} bytes)");
                    return ms.ToArray();
                }
            }
            catch (Exception ex)
            {
                LogError($"Error in GenerateAudio", ex);
                throw;
            }
        }

        private void LogMessage(string message)
        {
            try
            {
                File.AppendAllText(Path.Combine(_logDir, "sherpa_debug.log"),
                    $"{DateTime.Now}: {message}\n");
            }
            catch { }
        }

        private void LogError(string message, Exception ex = null)
        {
            try
            {
                string errorLog = Path.Combine(_logDir, "sherpa_error.log");
                string errorMessage = $"{DateTime.Now}: {message}\n";

                if (ex != null)
                {
                    errorMessage += $"Exception: {ex.GetType().Name}\n";
                    errorMessage += $"Message: {ex.Message}\n";
                    errorMessage += $"Stack Trace: {ex.StackTrace}\n";

                    if (ex.InnerException != null)
                    {
                        errorMessage += $"Inner Exception: {ex.InnerException.Message}\n";
                        errorMessage += $"Inner Stack Trace: {ex.InnerException.StackTrace}\n";
                    }
                }

                errorMessage += "\n";

                File.AppendAllText(errorLog, errorMessage);
            }
            catch { }
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                try
                {
                    LogMessage("Disposing SherpaTTS");
                    _tts?.Dispose();
                    LogMessage("SherpaTTS disposed successfully");
                }
                catch (Exception ex)
                {
                    LogError("Error disposing SherpaTTS", ex);
                }
                finally
                {
                    _disposed = true;
                }
            }
        }
    }
}
