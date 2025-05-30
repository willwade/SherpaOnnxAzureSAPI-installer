using NAudio.Wave;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Reflection;

namespace OpenSpeechTTS
{
    public class SherpaTTS : IDisposable
    {
        private object _tts; // Will hold OfflineTts instance
        private bool _disposed;
        private readonly string _logDir = "C:\\OpenSpeech";
        private bool _useRealTts = false; // Flag to control real vs mock TTS

        public SherpaTTS(string modelPath, string tokensPath, string lexiconPath, string dataDirPath)
        {
            try
            {
                // Create log directory if it doesn't exist
                if (!Directory.Exists(_logDir))
                {
                    Directory.CreateDirectory(_logDir);
                }

                LogMessage("Initializing SherpaTTS...");
                LogMessage($"Model Path: {modelPath}");
                LogMessage($"Tokens Path: {tokensPath}");
                LogMessage($"Data Directory: {dataDirPath}");

                // Try to initialize real SherpaOnnx TTS
                if (TryInitializeRealTts(modelPath, tokensPath))
                {
                    _useRealTts = true;
                    LogMessage("SherpaTTS initialized successfully with REAL TTS");
                }
                else
                {
                    _useRealTts = false;
                    LogMessage("SherpaTTS initialized in MOCK MODE (real TTS failed)");
                }
            }
            catch (Exception ex)
            {
                LogError("Error initializing SherpaTTS", ex);
                _useRealTts = false;
                LogMessage("Falling back to MOCK MODE due to initialization error");
            }
        }

        private bool TryInitializeRealTts(string modelPath, string tokensPath)
        {
            try
            {
                LogMessage("Attempting to initialize real SherpaOnnx TTS...");

                // Check if model files exist
                if (!File.Exists(modelPath))
                {
                    LogError($"Model file not found: {modelPath}");
                    return false;
                }

                if (!File.Exists(tokensPath))
                {
                    LogError($"Tokens file not found: {tokensPath}");
                    return false;
                }

                // Try to load SherpaOnnx assembly
                LogMessage("Loading SherpaOnnx types...");

                // Use reflection to avoid compile-time dependency issues
                var sherpaAssembly = Assembly.LoadFrom("sherpa-onnx.dll");
                var offlineTtsConfigType = sherpaAssembly.GetType("SherpaOnnx.OfflineTtsConfig");
                var offlineTtsType = sherpaAssembly.GetType("SherpaOnnx.OfflineTts");

                if (offlineTtsConfigType == null || offlineTtsType == null)
                {
                    LogError("Could not find SherpaOnnx types in assembly");
                    return false;
                }

                LogMessage("Creating TTS configuration...");

                // Create configuration using reflection
                var config = Activator.CreateInstance(offlineTtsConfigType);

                // Set configuration properties using reflection
                var modelProperty = offlineTtsConfigType.GetProperty("Model");
                var modelValue = modelProperty.GetValue(config);
                var vitsProperty = modelValue.GetType().GetProperty("Vits");
                var vitsValue = vitsProperty.GetValue(modelValue);

                // Set VITS model properties
                vitsValue.GetType().GetProperty("Model").SetValue(vitsValue, modelPath);
                vitsValue.GetType().GetProperty("Tokens").SetValue(vitsValue, tokensPath);
                vitsValue.GetType().GetProperty("NoiseScale").SetValue(vitsValue, 0.667f);
                vitsValue.GetType().GetProperty("NoiseScaleW").SetValue(vitsValue, 0.8f);
                vitsValue.GetType().GetProperty("LengthScale").SetValue(vitsValue, 1.0f);

                // Set other config properties
                modelValue.GetType().GetProperty("NumThreads").SetValue(modelValue, 1);
                modelValue.GetType().GetProperty("Debug").SetValue(modelValue, 0);
                modelValue.GetType().GetProperty("Provider").SetValue(modelValue, "cpu");

                LogMessage("Creating OfflineTts instance...");

                // Create TTS instance
                _tts = Activator.CreateInstance(offlineTtsType, config);

                LogMessage("Real SherpaOnnx TTS initialized successfully!");
                return true;
            }
            catch (Exception ex)
            {
                LogError($"Failed to initialize real TTS: {ex.Message}", ex);
                return false;
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
                if (_useRealTts && _tts != null)
                {
                    LogMessage($"Generating REAL audio bytes for text: '{text}'");
                    return GenerateRealAudioData(text);
                }
                else
                {
                    LogMessage($"Generating MOCK audio bytes for text: '{text}'");
                    return GenerateMockAudioData(text);
                }
            }
            catch (Exception ex)
            {
                LogError($"Error in GenerateAudio", ex);

                // Fall back to mock audio on error
                if (_useRealTts)
                {
                    LogMessage("Falling back to mock audio due to real TTS error");
                    return GenerateMockAudioData(text);
                }
                throw;
            }
        }

        private byte[] GenerateRealAudioData(string text)
        {
            try
            {
                LogMessage("Using real SherpaOnnx TTS to generate audio...");

                // Use reflection to call the Generate method
                var generateMethod = _tts.GetType().GetMethod("Generate", new[] { typeof(string), typeof(float), typeof(int) });
                if (generateMethod == null)
                {
                    LogError("Could not find Generate method on OfflineTts");
                    return GenerateMockAudioData(text);
                }

                // Call Generate(text, speed=1.0f, speakerId=0)
                var audioResult = generateMethod.Invoke(_tts, new object[] { text, 1.0f, 0 });

                // Get the samples from the result
                var samplesProperty = audioResult.GetType().GetProperty("Samples");
                var sampleRateProperty = audioResult.GetType().GetProperty("SampleRate");

                if (samplesProperty == null || sampleRateProperty == null)
                {
                    LogError("Could not find Samples or SampleRate properties on audio result");
                    return GenerateMockAudioData(text);
                }

                var samples = (float[])samplesProperty.GetValue(audioResult);
                var sampleRate = (int)sampleRateProperty.GetValue(audioResult);

                LogMessage($"Generated {samples.Length} samples at {sampleRate}Hz");

                // Convert float samples to WAV format
                return ConvertSamplesToWav(samples, sampleRate);
            }
            catch (Exception ex)
            {
                LogError($"Error generating real audio: {ex.Message}", ex);
                return GenerateMockAudioData(text);
            }
        }

        private byte[] ConvertSamplesToWav(float[] samples, int sampleRate)
        {
            try
            {
                using (var ms = new MemoryStream())
                {
                    using (var writer = new BinaryWriter(ms))
                    {
                        // WAV header
                        writer.Write(System.Text.Encoding.ASCII.GetBytes("RIFF"));
                        var dataSize = samples.Length * 2; // 16-bit samples
                        var fileSize = 36 + dataSize;
                        writer.Write((uint)fileSize);
                        writer.Write(System.Text.Encoding.ASCII.GetBytes("WAVE"));
                        writer.Write(System.Text.Encoding.ASCII.GetBytes("fmt "));
                        writer.Write((uint)16); // fmt chunk size
                        writer.Write((ushort)1); // PCM format
                        writer.Write((ushort)1); // mono
                        writer.Write((uint)sampleRate);
                        writer.Write((uint)(sampleRate * 2)); // byte rate
                        writer.Write((ushort)2); // block align
                        writer.Write((ushort)16); // bits per sample
                        writer.Write(System.Text.Encoding.ASCII.GetBytes("data"));
                        writer.Write((uint)dataSize);

                        // Convert float samples to 16-bit PCM
                        foreach (var sample in samples)
                        {
                            var pcmSample = (short)(sample * 32767);
                            writer.Write(pcmSample);
                        }
                    }

                    LogMessage($"Converted to WAV format: {ms.Length} bytes");
                    return ms.ToArray();
                }
            }
            catch (Exception ex)
            {
                LogError($"Error converting samples to WAV: {ex.Message}", ex);
                return GenerateMockAudioData("Error converting audio");
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
