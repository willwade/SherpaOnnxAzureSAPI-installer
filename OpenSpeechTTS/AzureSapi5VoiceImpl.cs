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

        // Updated SAPI5 Speak method implementation for Azure
        public int Speak(uint dwSpeakFlags, ref Guid rguidFormatId, ref WaveFormatEx pWaveFormatEx,
                        ref SpTTSFragList pTextFragList, IntPtr pOutputSite)
        {
            if (!_initialized)
                return unchecked((int)0x80004005); // E_FAIL

            try
            {
                // Extract text from fragment list
                string text = ExtractTextFromFragList(ref pTextFragList);
                if (string.IsNullOrEmpty(text))
                    return 0; // S_OK

                // Log the speak request for debugging
                try
                {
                    File.AppendAllText("C:\\OpenSpeech\\sapi_speak.log",
                        $"{DateTime.Now}: Azure TTS Speaking: {text}\n");
                }
                catch { }

                // Get the output site interface
                if (pOutputSite == IntPtr.Zero)
                    return unchecked((int)0x80070057); // E_INVALIDARG

                var outputSite = Marshal.GetObjectForIUnknown(pOutputSite) as ISpTTSEngineSite;
                if (outputSite == null)
                    return unchecked((int)0x80004005); // E_FAIL

                // Generate audio using Azure TTS
                var buffer = _azureTts.GenerateAudio(text, _locale);

                // Write audio data to output site
                uint bytesWritten = 0;
                int hr = outputSite.Write(Marshal.UnsafeAddrOfPinnedArrayElement(buffer, 0),
                                         (uint)buffer.Length, out bytesWritten);

                return hr;
            }
            catch (Exception ex)
            {
                // Log the error for debugging
                try
                {
                    File.AppendAllText("C:\\OpenSpeech\\sapi_error.log",
                        $"{DateTime.Now}: Error in Azure Speak: {ex.Message}\n{ex.StackTrace}\n\n");
                }
                catch { }

                return unchecked((int)0x80004005); // E_FAIL
            }
        }

        public int GetOutputFormat(ref Guid pTargetFormatId, ref WaveFormatEx pTargetWaveFormatEx,
                                  out Guid pOutputFormatId, out IntPtr ppCoMemOutputWaveFormatEx)
        {
            try
            {
                // Set our preferred format
                pOutputFormatId = new Guid("C31ADBAE-527F-4ff5-A230-F62BB61FF70C"); // SPDFID_WaveFormatEx

                // Allocate memory for WaveFormatEx
                int waveFormatSize = Marshal.SizeOf<WaveFormatEx>();
                ppCoMemOutputWaveFormatEx = Marshal.AllocCoTaskMem(waveFormatSize);

                // Initialize output format for Azure TTS (24kHz)
                var format = new WaveFormatEx
                {
                    wFormatTag = 1, // PCM
                    nChannels = 1, // Mono
                    nSamplesPerSec = 24000, // Sample rate
                    wBitsPerSample = 16,
                    nBlockAlign = 2, // (nChannels * wBitsPerSample) / 8
                    nAvgBytesPerSec = 24000 * 2, // nSamplesPerSec * nBlockAlign
                    cbSize = 0
                };

                Marshal.StructureToPtr(format, ppCoMemOutputWaveFormatEx, false);
                return 0; // S_OK
            }
            catch
            {
                pOutputFormatId = Guid.Empty;
                ppCoMemOutputWaveFormatEx = IntPtr.Zero;
                return unchecked((int)0x80004005); // E_FAIL
            }
        }

        // Helper method to extract text from SAPI fragment list
        private string ExtractTextFromFragList(ref SpTTSFragList fragList)
        {
            try
            {
                if (fragList.pTextStart == IntPtr.Zero || fragList.ulTextLen == 0)
                    return string.Empty;

                return Marshal.PtrToStringUni(fragList.pTextStart, (int)fragList.ulTextLen);
            }
            catch
            {
                return string.Empty;
            }
        }
    }
}
