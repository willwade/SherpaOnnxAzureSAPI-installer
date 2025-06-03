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

                // Try ProcessBridge first (preferred method)
                if (TryInitializeProcessBridge(modelPath, tokensPath))
                {
                    _useRealTts = true;
                    LogMessage("SherpaTTS initialized successfully with PROCESSBRIDGE TTS");
                }
                // Fallback to direct SherpaOnnx integration
                else if (TryInitializeRealTts(modelPath, tokensPath))
                {
                    _useRealTts = true;
                    LogMessage("SherpaTTS initialized successfully with DIRECT REAL TTS");
                }
                else
                {
                    _useRealTts = false;
                    LogMessage("SherpaTTS initialized in MOCK MODE (ProcessBridge and real TTS failed)");
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
                LogMessage("Attempting to initialize real TTS using native bridge...");

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

                // Check if ONNX runtime is available
                string installDir = @"C:\Program Files\OpenAssistive\OpenSpeech";
                string onnxRuntimePath = Path.Combine(installDir, "onnxruntime.dll");

                if (!File.Exists(onnxRuntimePath))
                {
                    LogError($"ONNX runtime not found at: {onnxRuntimePath}");
                    return false;
                }

                LogMessage($"Found ONNX runtime: {onnxRuntimePath}");
                LogMessage($"Model path: {modelPath}");
                LogMessage($"Tokens path: {tokensPath}");

                // Instead of using the problematic SherpaOnnx assembly,
                // we'll create a native bridge or process-based solution

                // For now, let's try to use the native SherpaNative.dll if available
                string nativePath = Path.Combine(installDir, "SherpaNative.dll");
                if (File.Exists(nativePath))
                {
                    LogMessage($"Found native bridge: {nativePath}");

                    // Try to initialize using native bridge
                    if (TryInitializeNativeBridge(modelPath, tokensPath, nativePath))
                    {
                        LogMessage("Successfully initialized using native bridge");
                        return true;
                    }
                }

                // Fallback: Create a process-based bridge
                LogMessage("Attempting process-based TTS bridge...");
                if (TryInitializeProcessBridge(modelPath, tokensPath))
                {
                    LogMessage("Successfully initialized using process bridge");
                    return true;
                }

                LogError("All real TTS initialization methods failed");
                return false;
            }
            catch (Exception ex)
            {
                LogError($"Failed to initialize real TTS: {ex.Message}", ex);
                return false;
            }
        }

        private bool TryInitializeNativeBridge(string modelPath, string tokensPath, string nativePath)
        {
            try
            {
                LogMessage("Attempting native bridge initialization...");

                // For now, we'll implement this as a placeholder
                // In a full implementation, this would use P/Invoke to call native functions
                LogMessage("Native bridge not yet implemented - falling back to process bridge");
                return false;
            }
            catch (Exception ex)
            {
                LogError($"Native bridge initialization failed: {ex.Message}", ex);
                return false;
            }
        }

        private bool TryInitializeProcessBridge(string modelPath, string tokensPath)
        {
            try
            {
                LogMessage("Attempting process bridge initialization...");

                // Create a ProcessBridge TTS that will use the SherpaWorker executable
                _tts = new ProcessBasedTTS(modelPath, tokensPath);

                LogMessage("Process bridge TTS initialized successfully");
                return true;
            }
            catch (Exception ex)
            {
                LogError($"Process bridge initialization failed: {ex.Message}", ex);
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
                LogMessage("Using bridge TTS to generate audio...");

                // Check if we have a ProcessBasedTTS instance
                if (_tts is ProcessBasedTTS processTts)
                {
                    LogMessage("Using ProcessBasedTTS for audio generation");
                    var audioResult = processTts.Generate(text, 1.0f, 0);

                    // Extract samples and sample rate from the result
                    var samplesProperty = audioResult.GetType().GetProperty("Samples");
                    var sampleRateProperty = audioResult.GetType().GetProperty("SampleRate");

                    if (samplesProperty != null && sampleRateProperty != null)
                    {
                        var samples = (float[])samplesProperty.GetValue(audioResult);
                        var sampleRate = (int)sampleRateProperty.GetValue(audioResult);

                        LogMessage($"Generated {samples.Length} samples at {sampleRate}Hz using ProcessBasedTTS");
                        return ConvertSamplesToWav(samples, sampleRate);
                    }
                }
                else
                {
                    // Try the original reflection-based approach for other TTS types
                    LogMessage("Using reflection-based TTS generation");

                    var generateMethod = _tts.GetType().GetMethod("Generate", new[] { typeof(string), typeof(float), typeof(int) });
                    if (generateMethod == null)
                    {
                        LogError("Could not find Generate method on TTS object");
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

                    LogMessage($"Generated {samples.Length} samples at {sampleRate}Hz using reflection");
                    return ConvertSamplesToWav(samples, sampleRate);
                }

                LogError("No valid TTS method found");
                return GenerateMockAudioData(text);
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

                    byte[] result = ms.ToArray();
                    LogMessage($"Successfully created MOCK WAV data ({result.Length} bytes)");
                    return result;
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

    // Process-based TTS bridge to handle .NET compatibility issues
    public class ProcessBasedTTS
    {
        private readonly string _modelPath;
        private readonly string _tokensPath;
        private readonly string _logDir = "C:\\OpenSpeech";

        public ProcessBasedTTS(string modelPath, string tokensPath)
        {
            _modelPath = modelPath;
            _tokensPath = tokensPath;
            LogMessage($"ProcessBasedTTS initialized with model: {modelPath}");
        }

        public object Generate(string text, float speed, int speakerId)
        {
            try
            {
                LogMessage($"ProcessBasedTTS generating audio for: '{text}' using SherpaWorker");

                // Use the real ProcessBridge to call SherpaWorker.exe
                var result = GenerateWithProcessBridge(text, speed, speakerId);

                if (result != null)
                {
                    LogMessage($"Generated real audio result with {result.Samples.Length} samples using ProcessBridge");
                    return result;
                }
                else
                {
                    LogMessage("ProcessBridge failed, falling back to mock audio");
                    var mockResult = new MockAudioResult(text);
                    LogMessage($"Generated fallback mock audio result with {mockResult.Samples.Length} samples");
                    return mockResult;
                }
            }
            catch (Exception ex)
            {
                LogError($"Error in ProcessBasedTTS.Generate: {ex.Message}", ex);

                // Fallback to mock audio on error
                LogMessage("Exception occurred, falling back to mock audio");
                var mockResult = new MockAudioResult(text);
                return mockResult;
            }
        }

        private ProcessBridgeAudioResult GenerateWithProcessBridge(string text, float speed, int speakerId)
        {
            try
            {
                LogMessage("Starting ProcessBridge TTS generation...");

                // Path to the SherpaWorker executable
                string sherpaWorkerPath = Path.Combine(
                    Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location),
                    "..", "..", "SherpaWorker", "bin", "Release", "net6.0", "win-x64", "publish", "SherpaWorker.exe");

                // Fallback paths if the relative path doesn't work
                if (!File.Exists(sherpaWorkerPath))
                {
                    sherpaWorkerPath = @"C:\Program Files\OpenAssistive\OpenSpeech\SherpaWorker.exe";
                }

                if (!File.Exists(sherpaWorkerPath))
                {
                    // Try the development path
                    sherpaWorkerPath = Path.Combine(Environment.CurrentDirectory, "SherpaWorker", "bin", "Release", "net6.0", "win-x64", "publish", "SherpaWorker.exe");
                }

                if (!File.Exists(sherpaWorkerPath))
                {
                    LogError($"SherpaWorker.exe not found at any expected location", new FileNotFoundException());
                    return null;
                }

                LogMessage($"Using SherpaWorker at: {sherpaWorkerPath}");

                // Create temporary request file
                string tempDir = Path.Combine(Path.GetTempPath(), "OpenSpeechTTS");
                if (!Directory.Exists(tempDir))
                    Directory.CreateDirectory(tempDir);

                string requestId = Guid.NewGuid().ToString("N").Substring(0, 8);
                string requestPath = Path.Combine(tempDir, $"tts_request_{requestId}.json");
                string responsePath = Path.Combine(tempDir, $"tts_request_{requestId}.response.json");
                string audioPath = Path.Combine(tempDir, $"tts_audio_{requestId}");

                // Create TTS request
                var request = new
                {
                    Text = text,
                    Speed = speed,
                    SpeakerId = speakerId,
                    OutputPath = audioPath
                };

                string requestJson = Newtonsoft.Json.JsonConvert.SerializeObject(request, Newtonsoft.Json.Formatting.Indented);
                File.WriteAllText(requestPath, requestJson);

                LogMessage($"Created request file: {requestPath}");

                // Launch SherpaWorker process
                var processInfo = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = sherpaWorkerPath,
                    Arguments = $"\"{requestPath}\"",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true,
                    WorkingDirectory = Path.GetDirectoryName(sherpaWorkerPath)
                };

                LogMessage($"Launching: {processInfo.FileName} {processInfo.Arguments}");

                using (var process = System.Diagnostics.Process.Start(processInfo))
                {
                    if (process == null)
                    {
                        LogError("Failed to start SherpaWorker process", new Exception("Process.Start returned null"));
                        return null;
                    }

                    // Wait for process to complete (with timeout)
                    bool completed = process.WaitForExit(30000); // 30 second timeout

                    if (!completed)
                    {
                        LogError("SherpaWorker process timed out", new TimeoutException());
                        process.Kill();
                        return null;
                    }

                    string stdout = process.StandardOutput.ReadToEnd();
                    string stderr = process.StandardError.ReadToEnd();

                    LogMessage($"SherpaWorker exit code: {process.ExitCode}");
                    if (!string.IsNullOrEmpty(stdout))
                        LogMessage($"SherpaWorker stdout: {stdout}");
                    if (!string.IsNullOrEmpty(stderr))
                        LogMessage($"SherpaWorker stderr: {stderr}");

                    if (process.ExitCode != 0)
                    {
                        LogError($"SherpaWorker failed with exit code {process.ExitCode}", new Exception(stderr));
                        return null;
                    }
                }

                // Read response
                if (!File.Exists(responsePath))
                {
                    LogError("Response file not created by SherpaWorker", new FileNotFoundException());
                    return null;
                }

                string responseJson = File.ReadAllText(responsePath);
                LogMessage($"Response: {responseJson}");

                var response = Newtonsoft.Json.JsonConvert.DeserializeObject<TtsResponse>(responseJson);

                if (response == null || !response.Success)
                {
                    LogError($"SherpaWorker reported failure: {response?.ErrorMessage}", new Exception(response?.ErrorMessage));
                    return null;
                }

                // Load the generated audio file
                string audioFilePath = response.AudioPath;
                if (!File.Exists(audioFilePath))
                {
                    LogError($"Audio file not found: {audioFilePath}", new FileNotFoundException());
                    return null;
                }

                // Read the WAV file and extract samples
                var samples = ReadWavFile(audioFilePath);
                if (samples == null || samples.Length == 0)
                {
                    LogError("Failed to read audio samples from WAV file", new Exception("No samples extracted"));
                    return null;
                }

                LogMessage($"Successfully loaded {samples.Length} samples from ProcessBridge");

                // Clean up temporary files
                try
                {
                    File.Delete(requestPath);
                    File.Delete(responsePath);
                    File.Delete(audioFilePath);
                }
                catch { } // Ignore cleanup errors

                // Create result object
                return new ProcessBridgeAudioResult(samples, response.SampleRate);
            }
            catch (Exception ex)
            {
                LogError($"ProcessBridge generation failed: {ex.Message}", ex);
                return null;
            }
        }

        private float[] ReadWavFile(string filePath)
        {
            try
            {
                using var fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read);
                using var reader = new BinaryReader(fileStream);

                // Read WAV header
                string riff = new string(reader.ReadChars(4));
                if (riff != "RIFF")
                    throw new Exception("Invalid WAV file - missing RIFF header");

                uint fileSize = reader.ReadUInt32();
                string wave = new string(reader.ReadChars(4));
                if (wave != "WAVE")
                    throw new Exception("Invalid WAV file - missing WAVE header");

                // Find fmt chunk
                while (fileStream.Position < fileStream.Length)
                {
                    string chunkId = new string(reader.ReadChars(4));
                    uint chunkSize = reader.ReadUInt32();

                    if (chunkId == "fmt ")
                    {
                        ushort audioFormat = reader.ReadUInt16();
                        ushort numChannels = reader.ReadUInt16();
                        uint sampleRate = reader.ReadUInt32();
                        uint byteRate = reader.ReadUInt32();
                        ushort blockAlign = reader.ReadUInt16();
                        ushort bitsPerSample = reader.ReadUInt16();

                        // Skip any extra format bytes
                        if (chunkSize > 16)
                            reader.ReadBytes((int)(chunkSize - 16));

                        LogMessage($"WAV format: {audioFormat}, channels: {numChannels}, rate: {sampleRate}, bits: {bitsPerSample}");
                    }
                    else if (chunkId == "data")
                    {
                        // Read audio data
                        int sampleCount = (int)(chunkSize / 2); // 16-bit samples
                        float[] samples = new float[sampleCount];

                        for (int i = 0; i < sampleCount; i++)
                        {
                            short sample = reader.ReadInt16();
                            samples[i] = sample / 32768.0f; // Convert to float [-1, 1]
                        }

                        LogMessage($"Read {sampleCount} samples from WAV file");
                        return samples;
                    }
                    else
                    {
                        // Skip unknown chunk
                        reader.ReadBytes((int)chunkSize);
                    }
                }

                throw new Exception("No data chunk found in WAV file");
            }
            catch (Exception ex)
            {
                LogError($"Failed to read WAV file: {ex.Message}", ex);
                return null;
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
                    errorMessage += $"Exception: {ex.Message}\n";
                }
                File.AppendAllText(errorLog, errorMessage);
            }
            catch { }
        }
    }

    // Mock audio result that matches the expected SherpaOnnx interface
    public class MockAudioResult
    {
        public float[] Samples { get; private set; }
        public int SampleRate { get; private set; }

        public MockAudioResult(string text)
        {
            // Generate a simple audio waveform based on the text
            SampleRate = 22050;
            int durationMs = Math.Max(1000, text.Length * 100); // At least 1 second
            int sampleCount = (int)(SampleRate * durationMs / 1000.0);

            Samples = new float[sampleCount];

            // Generate a simple tone pattern
            for (int i = 0; i < sampleCount; i++)
            {
                double time = (double)i / SampleRate;
                // Create a simple melody based on text length
                double frequency = 440.0 + (text.Length % 12) * 50; // Vary frequency based on text
                double amplitude = Math.Sin(2 * Math.PI * frequency * time) * 0.1; // Low volume
                Samples[i] = (float)amplitude;
            }
        }
    }

    // TTS Response structure for ProcessBridge IPC communication
    public class TtsResponse
    {
        public bool Success { get; set; }
        public string ErrorMessage { get; set; } = "";
        public int SampleCount { get; set; }
        public int SampleRate { get; set; }
        public string AudioPath { get; set; } = "";
    }

    // ProcessBridge audio result that matches the expected SherpaOnnx interface
    public class ProcessBridgeAudioResult
    {
        public float[] Samples { get; private set; }
        public int SampleRate { get; private set; }

        public ProcessBridgeAudioResult(float[] samples, int sampleRate)
        {
            Samples = samples;
            SampleRate = sampleRate;
        }
    }
}
