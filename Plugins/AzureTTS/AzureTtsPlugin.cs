using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using OpenSpeech.TTS.Plugins;

namespace OpenSpeech.TTS.Plugins.AzureTTS
{
    /// <summary>
    /// Azure TTS engine plugin
    /// </summary>
    public class AzureTtsPlugin : TtsEngineBase
    {
        /// <summary>
        /// Gets the name of the TTS engine
        /// </summary>
        public override string EngineName => "AzureTTS";
        
        /// <summary>
        /// Gets the version of the TTS engine
        /// </summary>
        public override string EngineVersion => "1.0.0";
        
        /// <summary>
        /// Gets a description of the TTS engine
        /// </summary>
        public override string EngineDescription => "Azure Cognitive Services Text-to-Speech";
        
        /// <summary>
        /// Gets a value indicating whether the engine is properly configured
        /// </summary>
        public override bool IsConfigured
        {
            get
            {
                // Check if the configuration is valid
                var config = new Dictionary<string, string>();
                return ValidateConfiguration(config);
            }
        }
        
        /// <summary>
        /// Validates the engine configuration
        /// </summary>
        /// <param name="configuration">The engine configuration</param>
        /// <returns>True if the configuration is valid, otherwise false</returns>
        public override bool ValidateConfiguration(Dictionary<string, string> configuration)
        {
            string subscriptionKey = GetConfigValue<string>(configuration, "subscriptionKey", string.Empty);
            string region = GetConfigValue<string>(configuration, "region", string.Empty);
            
            if (string.IsNullOrEmpty(subscriptionKey) || string.IsNullOrEmpty(region))
            {
                return false;
            }
            
            return true;
        }
        
        /// <summary>
        /// Gets the available voices for the TTS engine
        /// </summary>
        /// <param name="configuration">The engine configuration</param>
        /// <returns>A collection of voice information</returns>
        public override async Task<IEnumerable<TtsVoiceInfo>> GetAvailableVoicesAsync(Dictionary<string, string> configuration)
        {
            string subscriptionKey = GetConfigValue<string>(configuration, "subscriptionKey", string.Empty);
            string region = GetConfigValue<string>(configuration, "region", string.Empty);
            
            if (string.IsNullOrEmpty(subscriptionKey) || string.IsNullOrEmpty(region))
            {
                return new List<TtsVoiceInfo>();
            }
            
            try
            {
                // In a real implementation, this would call the Azure TTS API to get the list of voices
                // For this example, we'll return a static list
                var voices = new List<TtsVoiceInfo>
                {
                    new TtsVoiceInfo
                    {
                        Id = "en-US-JennyNeural",
                        Name = "Jenny",
                        Gender = "Female",
                        Language = "English",
                        Locale = "en-US",
                        SupportedStyles = new List<string> { "cheerful", "sad", "angry" }
                    },
                    new TtsVoiceInfo
                    {
                        Id = "en-US-GuyNeural",
                        Name = "Guy",
                        Gender = "Male",
                        Language = "English",
                        Locale = "en-US",
                        SupportedStyles = new List<string> { "cheerful", "sad" }
                    },
                    new TtsVoiceInfo
                    {
                        Id = "en-GB-SoniaNeural",
                        Name = "Sonia",
                        Gender = "Female",
                        Language = "English",
                        Locale = "en-GB",
                        SupportedStyles = new List<string> { "cheerful", "sad" }
                    }
                };
                
                return voices;
            }
            catch (Exception ex)
            {
                // Log the error
                Console.WriteLine($"Error getting Azure TTS voices: {ex.Message}");
                return new List<TtsVoiceInfo>();
            }
        }
        
        /// <summary>
        /// Synthesizes speech from text
        /// </summary>
        /// <param name="text">The text to synthesize</param>
        /// <param name="voiceId">The ID of the voice to use</param>
        /// <param name="configuration">The engine configuration</param>
        /// <returns>The synthesized audio data</returns>
        public override async Task<byte[]> SynthesizeSpeechAsync(string text, string voiceId, Dictionary<string, string> configuration)
        {
            string subscriptionKey = GetConfigValue<string>(configuration, "subscriptionKey", string.Empty);
            string region = GetConfigValue<string>(configuration, "region", string.Empty);
            
            if (string.IsNullOrEmpty(subscriptionKey) || string.IsNullOrEmpty(region))
            {
                throw new InvalidOperationException("Azure TTS is not properly configured");
            }
            
            try
            {
                // Use the OpenSpeechTTS library to generate audio
                using (var tts = new global::OpenSpeechTTS.AzureTTS(subscriptionKey, region, voiceId))
                {
                    return await tts.SynthesizeSpeechAsync(text);
                }
            }
            catch (Exception ex)
            {
                // Log the error
                Console.WriteLine($"Error synthesizing speech with Azure TTS: {ex.Message}");
                throw;
            }
        }
    }
} 