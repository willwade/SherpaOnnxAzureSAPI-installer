using System;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace OpenSpeechTTS
{
    [Guid("3d8f5c5e-9d6b-4b92-a12b-1a6dff80b6b3")]  // Different GUID from Sherpa implementation
    [ComVisible(true)]
    public class AzureSapi5VoiceImpl : ISpTTSEngine
    {
        private AzureTTS _tts;
        private bool _initialized;

        public AzureSapi5VoiceImpl()
        {
            try
            {
                // Get the voice token from the registry
                string voiceToken = null;
                using (var key = Registry.ClassesRoot.OpenSubKey(@"CLSID\{3d8f5c5e-9d6b-4b92-a12b-1a6dff80b6b3}\Token"))
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
                LogMessage($"Loading Azure voice token: {voiceToken}");

                string registryPath = $@"SOFTWARE\Microsoft\Speech\Voices\Tokens\{voiceToken}";
                
                var voiceKey = Registry.LocalMachine.OpenSubKey(registryPath);
                if (voiceKey == null)
                    throw new Exception($"Voice registry key not found: {registryPath}");

                var attributesKey = voiceKey.OpenSubKey("Attributes");
                if (attributesKey == null)
                    throw new Exception("Voice attributes not found in registry");

                // Read the Azure-specific settings from the registry
                var subscriptionKey = (string)attributesKey.GetValue("SubscriptionKey");
                var region = (string)attributesKey.GetValue("Region");
                var voiceName = (string)attributesKey.GetValue("VoiceName");

                // Log the settings we found
                LogMessage($"Registry values read for Azure voice '{voiceToken}':");
                LogMessage($"  Region: {region}");
                LogMessage($"  VoiceName: {voiceName}");
                LogMessage($"  SubscriptionKey: {(string.IsNullOrEmpty(subscriptionKey) ? "Not found" : "Found (not shown for security)")}");

                if (string.IsNullOrEmpty(subscriptionKey))
                    throw new Exception($"SubscriptionKey not found in registry for voice: {voiceToken}");
                if (string.IsNullOrEmpty(region))
                    throw new Exception($"Region not found in registry for voice: {voiceToken}");
                if (string.IsNullOrEmpty(voiceName))
                    throw new Exception($"VoiceName not found in registry for voice: {voiceToken}");

                LogMessage("All settings found, creating AzureTTS instance...");
                
                _tts = new AzureTTS(subscriptionKey, region, voiceName);
                _initialized = true;
                
                LogMessage($"Successfully initialized AzureTTS for voice: {voiceToken}");
            }
            catch (Exception ex)
            {
                // Log the error to a file for debugging
                LogError($"Error in AzureSapi5VoiceImpl constructor: {ex.Message}", ex);
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
            // Initialize output format - Azure TTS uses 24kHz by default
            actualFormat = new WaveFormatEx
            {
                wFormatTag = 1, // PCM
                nChannels = 1, // Mono
                nSamplesPerSec = 24000, // Sample rate
                wBitsPerSample = 16,
                nBlockAlign = 2, // (nChannels * wBitsPerSample) / 8
                nAvgBytesPerSec = 24000 * 2, // nSamplesPerSec * nBlockAlign
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
                
                File.AppendAllText(Path.Combine(logDir, "azure_sapi_debug.log"), 
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
                
                string errorLog = Path.Combine(logDir, "azure_sapi_error.log");
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
