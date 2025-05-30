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
        private readonly object _tts; // OfflineTts _tts; // TEMPORARY: Changed to object for testing
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

                LogMessage("TEMPORARY: Initializing SherpaTTS in mock mode");
                LogMessage($"Model Path: {modelPath}");
                LogMessage($"Tokens Path: {tokensPath}");
                LogMessage($"Data Directory: {dataDirPath}");

                // TEMPORARY: Comment out Sherpa ONNX initialization for testing
                /*
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
                */
                _tts = null; // TEMPORARY: Set to null for testing
                LogMessage("SherpaTTS initialized successfully (MOCK MODE)");
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
                LogMessage($"TEMPORARY: Generating MOCK speech for text: '{text}'");

                // TEMPORARY: Generate mock audio data instead of using Sherpa ONNX
                byte[] audioData = GenerateMockAudioData(text);
                stream.Write(audioData, 0, audioData.Length);

                LogMessage("Successfully wrote MOCK WAV data to stream");
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
                LogMessage($"TEMPORARY: Generating MOCK audio bytes for text: '{text}'");

                // TEMPORARY: Return mock audio data instead of using Sherpa ONNX
                return GenerateMockAudioData(text);
            }
            catch (Exception ex)
            {
                LogError($"Error in GenerateAudio", ex);
                throw;
            }
        }

        // TEMPORARY: Generate mock audio data for testing
        private byte[] GenerateMockAudioData(string text)
        {
            try
            {
                // Create a simple WAV file with 1 second of silence per 10 characters
                int durationMs = Math.Max(1000, text.Length * 100); // At least 1 second
                uint sampleRate = 22050;
                int samples = (int)(sampleRate * durationMs / 1000);

                using (var ms = new MemoryStream())
                {
                    using (var writer = new BinaryWriter(ms))
                    {
                        // WAV header
                        writer.Write(0x46464952); // "RIFF"
                        writer.Write(36 + samples * 2);
                        writer.Write(0x45564157); // "WAVE"
                        writer.Write(0x20746D66); // "fmt "
                        writer.Write(16);
                        writer.Write((short)1); // PCM
                        writer.Write((short)1); // Mono
                        writer.Write(sampleRate); // Sample rate
                        writer.Write(sampleRate * 2); // Bytes per second
                        writer.Write((short)2); // Block align
                        writer.Write((short)16); // Bits per sample
                        writer.Write(0x61746164); // "data"
                        writer.Write(samples * 2);

                        // Generate simple tone instead of silence for testing
                        for (int i = 0; i < samples; i++)
                        {
                            // Generate a simple 440Hz tone (A note)
                            double time = (double)i / sampleRate;
                            double amplitude = Math.Sin(2 * Math.PI * 440 * time) * 0.1; // Low volume
                            short sample = (short)(amplitude * short.MaxValue);
                            writer.Write(sample);
                        }
                    }

                    LogMessage($"Successfully created MOCK WAV data ({ms.Length} bytes)");
                    return ms.ToArray();
                }
            }
            catch (Exception ex)
            {
                LogError($"Error generating mock audio data: {ex.Message}", ex);
                // Return minimal WAV file on error
                return new byte[] { 0x52, 0x49, 0x46, 0x46, 0x24, 0x00, 0x00, 0x00, 0x57, 0x41, 0x56, 0x45 };
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
                    LogMessage("Disposing SherpaTTS (MOCK MODE)");
                    // TEMPORARY: _tts?.Dispose(); // Commented out for testing
                    LogMessage("SherpaTTS disposed successfully (MOCK MODE)");
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
