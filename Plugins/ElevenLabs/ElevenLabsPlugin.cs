using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using OpenSpeech.TTS.Plugins;
using Installer.Engines.ElevenLabs;

namespace Plugins.ElevenLabs
{
    /// <summary>
    /// Plugin adapter for ElevenLabs TTS engine
    /// </summary>
    public class ElevenLabsPlugin : TtsEngineBase
    {
        private readonly ElevenLabsEngine _engine;
        
        /// <summary>
        /// Initializes a new instance of the ElevenLabsPlugin class
        /// </summary>
        public ElevenLabsPlugin()
        {
            _engine = new ElevenLabsEngine();
        }
        
        /// <summary>
        /// Gets the name of the TTS engine
        /// </summary>
        public override string EngineName => _engine.EngineName;
        
        /// <summary>
        /// Gets the version of the TTS engine
        /// </summary>
        public override string EngineVersion => _engine.EngineVersion;
        
        /// <summary>
        /// Gets a description of the TTS engine
        /// </summary>
        public override string EngineDescription => _engine.EngineDescription;
        
        /// <summary>
        /// Gets a value indicating whether the engine is properly configured
        /// </summary>
        public override bool IsConfigured => !_engine.RequiresAuthentication || base.IsConfigured;
        
        /// <summary>
        /// Validates the engine configuration
        /// </summary>
        /// <param name="configuration">The engine configuration</param>
        /// <returns>True if the configuration is valid, otherwise false</returns>
        public override bool ValidateConfiguration(Dictionary<string, string> configuration)
        {
            // Check if API key is provided
            if (!configuration.TryGetValue("apiKey", out string apiKey) || string.IsNullOrEmpty(apiKey))
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
            try
            {
                var voices = await _engine.GetAvailableVoicesAsync(configuration);
                
                // Convert from the existing TtsVoiceInfo to our plugin's TtsVoiceInfo
                var result = new List<TtsVoiceInfo>();
                
                foreach (var voice in voices)
                {
                    result.Add(new TtsVoiceInfo
                    {
                        Id = voice.Id,
                        Name = voice.Name,
                        Gender = voice.Gender,
                        Language = voice.Locale.Split('-')[0],
                        Locale = voice.Locale,
                        SupportedStyles = voice.SupportedStyles ?? new List<string>(),
                        AdditionalInfo = new Dictionary<string, string>
                        {
                            { "EngineName", EngineName }
                        }
                    });
                }
                
                return result;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error getting ElevenLabs voices: {ex.Message}");
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
            return await _engine.SynthesizeSpeechAsync(text, voiceId, configuration);
        }
    }
} 