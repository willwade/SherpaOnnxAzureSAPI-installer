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

        private const string RegistryBasePath = @"SOFTWARE\Microsoft\Speech\Voices\Tokens\OpenSpeechAmyVoice";

        public Sapi5VoiceImpl()
        {
            try
            {
                var voiceKey = Registry.LocalMachine.OpenSubKey(RegistryBasePath);
                if (voiceKey == null)
                    throw new Exception("Voice registry key not found");

                var modelPath = (string)voiceKey.GetValue("ModelPath");
                var tokensPath = (string)voiceKey.GetValue("TokensPath");
                var dataDirPath = (string)voiceKey.GetValue("DataDirPath");

                if (string.IsNullOrEmpty(modelPath))
                    throw new Exception("ModelPath not found in registry");
                if (string.IsNullOrEmpty(tokensPath))
                    throw new Exception("TokensPath not found in registry");
                if (string.IsNullOrEmpty(dataDirPath))
                    throw new Exception("DataDirPath not found in registry");

                _tts = new SherpaTTS(modelPath, tokensPath, "", dataDirPath);
                _initialized = true;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error in Sapi5VoiceImpl constructor: {ex.Message}", ex);
            }
        }

        public void Speak(string text, uint flags, IntPtr reserved)
        {
            if (!_initialized)
                throw new Exception("TTS engine not initialized");

            try
            {
                var memoryStream = new MemoryStream();
                _tts.SpeakToWaveStream(text, memoryStream);
                var buffer = memoryStream.ToArray();
                
                // Copy the buffer to the reserved memory location if provided
                if (reserved != IntPtr.Zero)
                {
                    Marshal.Copy(buffer, 0, reserved, buffer.Length);
                }
            }
            catch (Exception ex)
            {
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
