using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace OpenSpeech.TTS.Plugins
{
    /// <summary>
    /// Interface for TTS engine plugins
    /// </summary>
    public interface ITtsEngine
    {
        /// <summary>
        /// Gets the name of the TTS engine
        /// </summary>
        string EngineName { get; }
        
        /// <summary>
        /// Gets the version of the TTS engine
        /// </summary>
        string EngineVersion { get; }
        
        /// <summary>
        /// Gets a description of the TTS engine
        /// </summary>
        string EngineDescription { get; }
        
        /// <summary>
        /// Gets a value indicating whether the engine is properly configured
        /// </summary>
        bool IsConfigured { get; }
        
        /// <summary>
        /// Validates the engine configuration
        /// </summary>
        /// <param name="configuration">The engine configuration</param>
        /// <returns>True if the configuration is valid, otherwise false</returns>
        bool ValidateConfiguration(Dictionary<string, string> configuration);
        
        /// <summary>
        /// Gets the available voices for the TTS engine
        /// </summary>
        /// <param name="configuration">The engine configuration</param>
        /// <returns>A collection of voice information</returns>
        Task<IEnumerable<TtsVoiceInfo>> GetAvailableVoicesAsync(Dictionary<string, string> configuration);
        
        /// <summary>
        /// Synthesizes speech from text
        /// </summary>
        /// <param name="text">The text to synthesize</param>
        /// <param name="voiceId">The ID of the voice to use</param>
        /// <param name="configuration">The engine configuration</param>
        /// <returns>The synthesized audio data</returns>
        Task<byte[]> SynthesizeSpeechAsync(string text, string voiceId, Dictionary<string, string> configuration);
    }
    
    /// <summary>
    /// Information about a TTS voice
    /// </summary>
    public class TtsVoiceInfo
    {
        /// <summary>
        /// Gets or sets the voice ID
        /// </summary>
        public string Id { get; set; }
        
        /// <summary>
        /// Gets or sets the voice name
        /// </summary>
        public string Name { get; set; }
        
        /// <summary>
        /// Gets or sets the voice gender
        /// </summary>
        public string Gender { get; set; }
        
        /// <summary>
        /// Gets or sets the voice language
        /// </summary>
        public string Language { get; set; }
        
        /// <summary>
        /// Gets or sets the voice locale
        /// </summary>
        public string Locale { get; set; }
        
        /// <summary>
        /// Gets or sets the voice styles
        /// </summary>
        public List<string> SupportedStyles { get; set; } = new List<string>();
        
        /// <summary>
        /// Gets or sets additional voice information
        /// </summary>
        public Dictionary<string, string> AdditionalInfo { get; set; } = new Dictionary<string, string>();
    }
} 