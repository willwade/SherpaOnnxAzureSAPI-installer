using System;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace OpenSpeechTTS
{
    [Guid("3d8f5c5d-9d6b-4b92-a12b-1a6dff80b6b2")]
    [ComVisible(true)]
    public class Sapi5VoiceImpl : ISpTTSEngine
    {
        private SherpaTTS _tts;
        private bool _initialized;

        public Sapi5VoiceImpl()
        {
            try
            {
                // Get the voice token from the registry
                string voiceToken = null;
                using (var key = Registry.ClassesRoot.OpenSubKey(@"CLSID\{3d8f5c5d-9d6b-4b92-a12b-1a6dff80b6b2}\Token"))
                {
                    if (key != null)
                    {
                        voiceToken = (string)key.GetValue("");
                    }
                }

                if (string.IsNullOrEmpty(voiceToken))
                {
                    throw new Exception("Voice token not found in registry");
                }

                // Log which voice token we're trying to load
                LogMessage($"Loading voice token: {voiceToken}");

                string registryPath = $@"SOFTWARE\Microsoft\Speech\Voices\Tokens\{voiceToken}";
                
                var voiceKey = Registry.LocalMachine.OpenSubKey(registryPath);
                if (voiceKey == null)
                    throw new Exception($"Voice registry key not found: {registryPath}");

                var attributesKey = voiceKey.OpenSubKey("Attributes");
                if (attributesKey == null)
                    throw new Exception("Voice attributes not found in registry");

                // Read the model paths from the registry for THIS specific voice
                var modelPath = (string)attributesKey.GetValue("Model Path");
                var tokensPath = (string)attributesKey.GetValue("Tokens Path");
                var dataDir = (string)attributesKey.GetValue("Data Directory");

                // Log the paths we found
                LogMessage($"Registry values read for voice '{voiceToken}':");
                LogMessage($"  ModelPath: {modelPath}");
                LogMessage($"  TokensPath: {tokensPath}");
                LogMessage($"  DataDirPath: {dataDir}");

                if (string.IsNullOrEmpty(modelPath))
                    throw new Exception($"ModelPath not found in registry for voice: {voiceToken}");
                if (string.IsNullOrEmpty(tokensPath))
                    throw new Exception($"TokensPath not found in registry for voice: {voiceToken}");

                // Verify that the files exist
                if (!File.Exists(modelPath))
                    throw new Exception($"Model file does not exist: {modelPath}");
                if (!File.Exists(tokensPath))
                    throw new Exception($"Tokens file does not exist: {tokensPath}");

                // If dataDir is not specified, use the directory containing the model
                string dataDirPath = dataDir;
                if (string.IsNullOrEmpty(dataDirPath))
                {
                    dataDirPath = Path.GetDirectoryName(modelPath);
                    if (string.IsNullOrEmpty(dataDirPath))
                    {
                        dataDirPath = Path.GetDirectoryName(tokensPath);
                    }
                }

                LogMessage("All files exist, creating SherpaTTS instance...");
                
                _tts = new SherpaTTS(modelPath, tokensPath, "", dataDirPath);
                _initialized = true;
                
                LogMessage($"Successfully initialized SherpaTTS for voice: {voiceToken}");
            }
            catch (Exception ex)
            {
                // Log the error to a file for debugging
                LogError($"Error in Sapi5VoiceImpl constructor: {ex.Message}", ex);
                throw;
            }
        }

        public void Speak(string text, uint flags, IntPtr reserved)
        {
            if (!_initialized)
                throw new Exception("TTS engine not initialized");

            try
            {
                // Log the speak request for debugging
                LogMessage($"Speaking text: {text}");

                // Generate audio data
                var memoryStream = new MemoryStream();
                _tts.SpeakToWaveStream(text, memoryStream);
                var buffer = memoryStream.ToArray();
                
                // Log the audio generation result
                LogMessage($"Generated {buffer.Length} bytes of audio data");
                
                // Copy the buffer to the reserved memory location if provided
                if (reserved != IntPtr.Zero)
                {
                    Marshal.Copy(buffer, 0, reserved, buffer.Length);
                }
                else
                {
                    // If no reserved memory is provided, we can't output the audio
                    // This is a common issue with SAPI5 integration
                    LogMessage("Warning: No reserved memory provided for audio output");
                }
            }
            catch (Exception ex)
            {
                // Log the error for debugging
                LogError($"Error in Speak: {ex.Message}", ex);
                throw;
            }
        }

        public void GetOutputFormat(ref Guid targetFormatId, ref WaveFormatEx targetFormat, out Guid actualFormatId, out WaveFormatEx actualFormat)
        {
            // Initialize output format
            actualFormat = new WaveFormatEx
            {
                wFormatTag = 1, // PCM
                nChannels = 1, // Mono
                nSamplesPerSec = 22050, // Sample rate
                wBitsPerSample = 16,
                nBlockAlign = 2, // (nChannels * wBitsPerSample) / 8
                nAvgBytesPerSec = 22050 * 2, // nSamplesPerSec * nBlockAlign
                cbSize = 0
            };

            // Use the same format ID as the target
            actualFormatId = targetFormatId;
        }

        private void LogMessage(string message)
        {
            try
            {
                string logDir = "C:\\OpenSpeech";
                if (!Directory.Exists(logDir))
                {
                    Directory.CreateDirectory(logDir);
                }
                
                File.AppendAllText(Path.Combine(logDir, "sapi_debug.log"), 
                    $"{DateTime.Now}: {message}\n");
            }
            catch { }
        }

        private void LogError(string message, Exception ex = null)
        {
            try
            {
                string logDir = "C:\\OpenSpeech";
                if (!Directory.Exists(logDir))
                {
                    Directory.CreateDirectory(logDir);
                }
                
                string errorLog = Path.Combine(logDir, "sapi_error.log");
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
    }
}
