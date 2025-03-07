using System;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace OpenSpeechTTS
{
    [Guid("3d8f5c5e-9d6b-4b92-a12b-1a6dff80b6b3")]
    [ComVisible(true)]
    public class AzureSapi5VoiceImpl : ISpTTSEngine
    {
        private AzureTTS _azureTts;
        private bool _initialized;
        private string _locale;

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

                string registryPath = $@"SOFTWARE\Microsoft\Speech\Voices\Tokens\{voiceToken}";
                
                var voiceKey = Registry.LocalMachine.OpenSubKey(registryPath);
                if (voiceKey == null)
                    throw new Exception($"Voice registry key not found: {registryPath}");

                var attributesKey = voiceKey.OpenSubKey("Attributes");
                if (attributesKey == null)
                    throw new Exception("Voice attributes not found in registry");

                // Get Azure TTS parameters
                var subscriptionKey = (string)attributesKey.GetValue("SubscriptionKey");
                var region = (string)attributesKey.GetValue("Region");
                var voiceName = (string)attributesKey.GetValue("VoiceName");
                var selectedStyle = (string)attributesKey.GetValue("SelectedStyle");
                var selectedRole = (string)attributesKey.GetValue("SelectedRole");
                
                // Get locale from language LCID
                var lcid = (string)attributesKey.GetValue("Language");
                _locale = GetLocaleFromLcid(lcid);

                if (string.IsNullOrEmpty(subscriptionKey))
                    throw new Exception("SubscriptionKey not found in registry");
                if (string.IsNullOrEmpty(region))
                    throw new Exception("Region not found in registry");
                if (string.IsNullOrEmpty(voiceName))
                    throw new Exception("VoiceName not found in registry");

                _azureTts = new AzureTTS(subscriptionKey, region, voiceName, selectedStyle, selectedRole);
                _initialized = true;
                
                // Log successful initialization
                try
                {
                    File.AppendAllText("C:\\OpenSpeech\\sapi_init.log", 
                        $"{DateTime.Now}: Successfully initialized Azure TTS engine\n" +
                        $"Voice: {voiceToken}\n" +
                        $"Azure Voice: {voiceName}\n" +
                        $"Region: {region}\n" +
                        $"Locale: {_locale}\n\n");
                }
                catch { }
            }
            catch (Exception ex)
            {
                // Log the error to a file for debugging
                try
                {
                    File.AppendAllText("C:\\OpenSpeech\\sapi_error.log", 
                        $"{DateTime.Now}: Error in AzureSapi5VoiceImpl constructor: {ex.Message}\n{ex.StackTrace}\n\n");
                }
                catch { }
                
                throw new Exception($"Error in AzureSapi5VoiceImpl constructor: {ex.Message}", ex);
            }
        }

        private string GetLocaleFromLcid(string lcid)
        {
            // Convert LCID from hex to decimal
            if (int.TryParse(lcid, System.Globalization.NumberStyles.HexNumber, null, out int lcidValue))
            {
                try
                {
                    var culture = System.Globalization.CultureInfo.GetCultureInfo(lcidValue);
                    return culture.Name;
                }
                catch
                {
                    // Default to en-US if conversion fails
                    return "en-US";
                }
            }
            
            // Default to en-US if parsing fails
            return "en-US";
        }

        public void Speak(string text, uint flags, IntPtr reserved)
        {
            if (!_initialized)
                throw new Exception("TTS engine not initialized");

            try
            {
                // Log the speak request for debugging
                try
                {
                    File.AppendAllText("C:\\OpenSpeech\\sapi_speak.log", 
                        $"{DateTime.Now}: Speaking text with Azure TTS: {text}\nFlags: {flags}\nReserved: {reserved}\n\n");
                }
                catch { }

                // Generate audio data
                var buffer = _azureTts.GenerateAudio(text, _locale);
                
                // Log the audio generation result
                try
                {
                    File.AppendAllText("C:\\OpenSpeech\\sapi_speak.log", 
                        $"Generated {buffer.Length} bytes of audio data\n\n");
                }
                catch { }
                
                // Copy the buffer to the reserved memory location if provided
                if (reserved != IntPtr.Zero)
                {
                    Marshal.Copy(buffer, 0, reserved, buffer.Length);
                }
                else
                {
                    // If no reserved memory is provided, we can't output the audio
                    // This is a common issue with SAPI5 integration
                    try
                    {
                        File.AppendAllText("C:\\OpenSpeech\\sapi_speak.log", 
                            $"Warning: No reserved memory provided for audio output\n\n");
                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                // Log the error for debugging
                try
                {
                    File.AppendAllText("C:\\OpenSpeech\\sapi_error.log", 
                        $"{DateTime.Now}: Error in Azure Speak: {ex.Message}\nText: {text}\n{ex.StackTrace}\n\n");
                }
                catch { }
                
                throw new Exception($"Error in Azure Speak: {ex.Message}", ex);
            }
        }

        public void GetOutputFormat(ref Guid targetFormatId, ref WaveFormatEx targetFormat, out Guid actualFormatId, out WaveFormatEx actualFormat)
        {
            // Initialize output format for Azure TTS (24kHz)
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
    }
}
