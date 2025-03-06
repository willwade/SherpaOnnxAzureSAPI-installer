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

                string registryPath = $@"SOFTWARE\Microsoft\Speech\Voices\Tokens\{voiceToken}";
                
                var voiceKey = Registry.LocalMachine.OpenSubKey(registryPath);
                if (voiceKey == null)
                    throw new Exception($"Voice registry key not found: {registryPath}");

                var attributesKey = voiceKey.OpenSubKey("Attributes");
                if (attributesKey == null)
                    throw new Exception("Voice attributes not found in registry");

                var modelPath = (string)attributesKey.GetValue("Model Path");
                var tokensPath = (string)attributesKey.GetValue("Tokens Path");

                if (string.IsNullOrEmpty(modelPath))
                    throw new Exception("ModelPath not found in registry");
                if (string.IsNullOrEmpty(tokensPath))
                    throw new Exception("TokensPath not found in registry");

                // Get the directory containing the model as the data directory
                string dataDirPath = Path.GetDirectoryName(modelPath);
                if (string.IsNullOrEmpty(dataDirPath))
                {
                    dataDirPath = Path.GetDirectoryName(tokensPath);
                }

                _tts = new SherpaTTS(modelPath, tokensPath, "", dataDirPath);
                _initialized = true;
            }
            catch (Exception ex)
            {
                // Log the error to a file for debugging
                try
                {
                    File.AppendAllText("C:\\OpenSpeech\\sapi_error.log", 
                        $"{DateTime.Now}: Error in Sapi5VoiceImpl constructor: {ex.Message}\n{ex.StackTrace}\n\n");
                }
                catch { }
                
                throw new Exception($"Error in Sapi5VoiceImpl constructor: {ex.Message}", ex);
            }
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
                        $"{DateTime.Now}: Speaking text: {text}\nFlags: {flags}\nReserved: {reserved}\n\n");
                }
                catch { }

                // Generate audio data
                var memoryStream = new MemoryStream();
                _tts.SpeakToWaveStream(text, memoryStream);
                var buffer = memoryStream.ToArray();
                
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
                        $"{DateTime.Now}: Error in Speak: {ex.Message}\nText: {text}\n{ex.StackTrace}\n\n");
                }
                catch { }
                
                throw new Exception($"Error in Speak: {ex.Message}", ex);
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
    }
}
