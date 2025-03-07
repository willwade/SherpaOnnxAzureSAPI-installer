using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace OpenSpeech.TTS.Plugins
{
    /// <summary>
    /// Base class for TTS engines
    /// </summary>
    public abstract class TtsEngineBase : ITtsEngine
    {
        /// <summary>
        /// Gets the name of the TTS engine
        /// </summary>
        public abstract string EngineName { get; }
        
        /// <summary>
        /// Gets the version of the TTS engine
        /// </summary>
        public abstract string EngineVersion { get; }
        
        /// <summary>
        /// Gets a description of the TTS engine
        /// </summary>
        public abstract string EngineDescription { get; }
        
        /// <summary>
        /// Gets a value indicating whether the engine is properly configured
        /// </summary>
        public virtual bool IsConfigured => true;
        
        /// <summary>
        /// Validates the engine configuration
        /// </summary>
        /// <param name="configuration">The engine configuration</param>
        /// <returns>True if the configuration is valid, otherwise false</returns>
        public abstract bool ValidateConfiguration(Dictionary<string, string> configuration);
        
        /// <summary>
        /// Gets the available voices for the TTS engine
        /// </summary>
        /// <param name="configuration">The engine configuration</param>
        /// <returns>A collection of voice information</returns>
        public abstract Task<IEnumerable<TtsVoiceInfo>> GetAvailableVoicesAsync(Dictionary<string, string> configuration);
        
        /// <summary>
        /// Synthesizes speech from text
        /// </summary>
        /// <param name="text">The text to synthesize</param>
        /// <param name="voiceId">The ID of the voice to use</param>
        /// <param name="configuration">The engine configuration</param>
        /// <returns>The synthesized audio data</returns>
        public abstract Task<byte[]> SynthesizeSpeechAsync(string text, string voiceId, Dictionary<string, string> configuration);
        
        /// <summary>
        /// Gets a configuration value
        /// </summary>
        /// <typeparam name="T">The type of the configuration value</typeparam>
        /// <param name="configuration">The configuration dictionary</param>
        /// <param name="key">The configuration key</param>
        /// <param name="defaultValue">The default value to return if the key is not found</param>
        /// <returns>The configuration value</returns>
        protected T GetConfigValue<T>(Dictionary<string, string> configuration, string key, T defaultValue)
        {
            if (configuration == null || !configuration.ContainsKey(key))
            {
                return defaultValue;
            }
            
            try
            {
                string value = configuration[key];
                
                if (typeof(T) == typeof(string))
                {
                    return (T)(object)value;
                }
                else if (typeof(T) == typeof(int))
                {
                    return (T)(object)int.Parse(value);
                }
                else if (typeof(T) == typeof(double))
                {
                    return (T)(object)double.Parse(value);
                }
                else if (typeof(T) == typeof(bool))
                {
                    return (T)(object)bool.Parse(value);
                }
                else
                {
                    throw new NotSupportedException($"Type {typeof(T).Name} is not supported for configuration values");
                }
            }
            catch
            {
                return defaultValue;
            }
        }
    }
} 