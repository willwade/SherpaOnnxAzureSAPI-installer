using System;
using System.IO;
using System.Runtime.InteropServices;
using Installer.Core.Base;
using Installer.Core.Interfaces;
using Microsoft.Win32;

namespace Installer.Engines.Azure
{
    /// <summary>
    /// SAPI voice implementation for Azure TTS
    /// </summary>
    [Guid("3d8f5c5e-9d6b-4b92-a12b-1a6dff80b6b3")]
    [ComVisible(true)]
    public class AzureSapi5VoiceImpl : SapiVoiceImplBase
    {
        private OpenSpeechTTS.AzureTTS _azureTts;
        private string _locale;
        
        /// <summary>
        /// Creates a new instance of the AzureSapi5VoiceImpl class
        /// </summary>
        public AzureSapi5VoiceImpl()
        {
            try
            {
                // Get the voice token from the registry
                string voiceToken = GetVoiceToken(new Guid("3d8f5c5e-9d6b-4b92-a12b-1a6dff80b6b3"));
                
                // Get voice attributes
                var attributesKey = GetVoiceAttributes(voiceToken);
                
                // Get Azure TTS parameters
                var subscriptionKey = (string)attributesKey.GetValue("SubscriptionKey");
                var region = (string)attributesKey.GetValue("Region");
                var voiceName = (string)attributesKey.GetValue("VoiceName");
                var selectedStyle = (string)attributesKey.GetValue("SelectedStyle");
                var selectedRole = (string)attributesKey.GetValue("SelectedRole");
                _locale = (string)attributesKey.GetValue("Locale");
                
                if (string.IsNullOrEmpty(subscriptionKey))
                    throw new Exception("SubscriptionKey not found in registry");
                if (string.IsNullOrEmpty(region))
                    throw new Exception("Region not found in registry");
                if (string.IsNullOrEmpty(voiceName))
                    throw new Exception("VoiceName not found in registry");
                
                // Create Azure TTS instance
                _azureTts = new OpenSpeechTTS.AzureTTS(subscriptionKey, region, voiceName, selectedStyle, selectedRole);
                _initialized = true;
                
                LogMessage($"Initialized AzureSapi5VoiceImpl with voice: {voiceName}");
            }
            catch (Exception ex)
            {
                LogError("Error initializing AzureSapi5VoiceImpl", ex);
                throw;
            }
        }
        
        /// <summary>
        /// Speaks the provided text
        /// </summary>
        /// <param name="text">Text to speak</param>
        /// <param name="flags">SAPI flags</param>
        /// <param name="reserved">Reserved parameter</param>
        public override void Speak(string text, uint flags, IntPtr reserved)
        {
            try
            {
                if (!_initialized || _azureTts == null)
                {
                    LogError("AzureSapi5VoiceImpl not initialized");
                    return;
                }
                
                LogMessage($"Speaking text: {text}");
                
                // Generate and play audio
                _azureTts.SpeakAsync(text).Wait();
            }
            catch (Exception ex)
            {
                LogError("Error speaking text", ex);
            }
        }
        
        /// <summary>
        /// Gets the output format for the voice
        /// </summary>
        /// <param name="targetFormatId">Target format ID</param>
        /// <param name="targetFormat">Target format</param>
        /// <param name="actualFormatId">Actual format ID</param>
        /// <param name="actualFormat">Actual format</param>
        public override void GetOutputFormat(ref Guid targetFormatId, ref WaveFormatEx targetFormat, out Guid actualFormatId, out WaveFormatEx actualFormat)
        {
            // Azure TTS uses 24kHz 16-bit mono PCM
            actualFormatId = targetFormatId;
            actualFormat = new WaveFormatEx
            {
                wFormatTag = 1, // PCM
                nChannels = 1, // Mono
                nSamplesPerSec = 24000, // 24kHz
                wBitsPerSample = 16, // 16-bit
                nBlockAlign = 2, // 2 bytes per sample (16-bit mono)
                nAvgBytesPerSec = 48000, // 24000 * 2
                cbSize = 0
            };
        }
        
        /// <summary>
        /// Finalizes the instance
        /// </summary>
        ~AzureSapi5VoiceImpl()
        {
            try
            {
                _azureTts?.Dispose();
                _azureTts = null;
            }
            catch
            {
                // Ignore errors during finalization
            }
        }
    }
} 