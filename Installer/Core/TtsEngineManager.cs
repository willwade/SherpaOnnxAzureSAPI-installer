using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using OpenSpeech.TTS.Plugins;

namespace Installer.Core
{
    /// <summary>
    /// Manages TTS engines
    /// </summary>
    public class TtsEngineManager
    {
        private readonly ConfigurationManager _configManager;
        private readonly Dictionary<string, ITtsEngine> _engines = new Dictionary<string, ITtsEngine>();
        
        /// <summary>
        /// Initializes a new instance of the TtsEngineManager class
        /// </summary>
        /// <param name="configManager">The configuration manager</param>
        public TtsEngineManager(ConfigurationManager configManager)
        {
            _configManager = configManager;
        }
        
        /// <summary>
        /// Registers a TTS engine
        /// </summary>
        /// <param name="engine">The TTS engine to register</param>
        public void RegisterEngine(ITtsEngine engine)
        {
            if (engine == null)
            {
                throw new ArgumentNullException(nameof(engine));
            }
            
            if (string.IsNullOrEmpty(engine.EngineName))
            {
                throw new ArgumentException("Engine name cannot be null or empty", nameof(engine));
            }
            
            if (_engines.ContainsKey(engine.EngineName))
            {
                // Replace the existing engine
                _engines[engine.EngineName] = engine;
            }
            else
            {
                // Add the new engine
                _engines.Add(engine.EngineName, engine);
            }
        }
        
        /// <summary>
        /// Gets a TTS engine by name
        /// </summary>
        /// <param name="engineName">The name of the engine</param>
        /// <returns>The TTS engine, or null if not found</returns>
        public ITtsEngine GetEngine(string engineName)
        {
            if (string.IsNullOrEmpty(engineName))
            {
                return null;
            }
            
            if (_engines.TryGetValue(engineName, out ITtsEngine engine))
            {
                return engine;
            }
            
            return null;
        }
        
        /// <summary>
        /// Gets all registered TTS engines
        /// </summary>
        /// <returns>A collection of TTS engines</returns>
        public IEnumerable<ITtsEngine> GetAllEngines()
        {
            return _engines.Values;
        }
        
        /// <summary>
        /// Gets the names of all registered TTS engines
        /// </summary>
        /// <returns>A collection of engine names</returns>
        public IEnumerable<string> GetEngineNames()
        {
            return _engines.Keys;
        }
        
        /// <summary>
        /// Gets the default TTS engine
        /// </summary>
        /// <returns>The default TTS engine, or null if no engines are registered</returns>
        public ITtsEngine GetDefaultEngine()
        {
            string defaultEngineName = _configManager.GetDefaultEngine();
            
            if (string.IsNullOrEmpty(defaultEngineName))
            {
                // If no default engine is specified, use the first enabled engine
                foreach (var engine in _engines.Values)
                {
                    if (_configManager.IsEngineEnabled(engine.EngineName))
                    {
                        return engine;
                    }
                }
                
                // If no engines are enabled, use the first engine
                return _engines.Values.FirstOrDefault();
            }
            
            // Try to get the default engine
            if (_engines.TryGetValue(defaultEngineName, out ITtsEngine engine))
            {
                return engine;
            }
            
            // If the default engine is not found, use the first engine
            return _engines.Values.FirstOrDefault();
        }
        
        /// <summary>
        /// Gets all available voices from all enabled engines
        /// </summary>
        /// <returns>A collection of voice information</returns>
        public async Task<IEnumerable<TtsVoiceInfo>> GetAllVoicesAsync()
        {
            var voices = new List<TtsVoiceInfo>();
            
            foreach (var engine in _engines.Values)
            {
                if (_configManager.IsEngineEnabled(engine.EngineName))
                {
                    try
                    {
                        var engineConfig = _configManager.GetEngineConfiguration(engine.EngineName);
                        
                        if (engine.ValidateConfiguration(engineConfig))
                        {
                            var engineVoices = await engine.GetAvailableVoicesAsync(engineConfig);
                            
                            foreach (var voice in engineVoices)
                            {
                                // Add engine name to additional info
                                if (voice.AdditionalInfo == null)
                                {
                                    voice.AdditionalInfo = new Dictionary<string, string>();
                                }
                                
                                voice.AdditionalInfo["EngineName"] = engine.EngineName;
                                
                                voices.Add(voice);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error getting voices for engine {engine.EngineName}: {ex.Message}");
                    }
                }
            }
            
            return voices;
        }
        
        /// <summary>
        /// Synthesizes speech using the specified engine
        /// </summary>
        /// <param name="engineName">The name of the engine to use</param>
        /// <param name="text">The text to synthesize</param>
        /// <param name="voiceId">The ID of the voice to use</param>
        /// <returns>The synthesized audio data</returns>
        public async Task<byte[]> SynthesizeSpeechAsync(string engineName, string text, string voiceId)
        {
            if (string.IsNullOrEmpty(engineName))
            {
                throw new ArgumentException("Engine name cannot be null or empty", nameof(engineName));
            }
            
            if (string.IsNullOrEmpty(text))
            {
                throw new ArgumentException("Text cannot be null or empty", nameof(text));
            }
            
            if (string.IsNullOrEmpty(voiceId))
            {
                throw new ArgumentException("Voice ID cannot be null or empty", nameof(voiceId));
            }
            
            if (!_engines.TryGetValue(engineName, out ITtsEngine engine))
            {
                throw new ArgumentException($"Engine not found: {engineName}", nameof(engineName));
            }
            
            if (!_configManager.IsEngineEnabled(engineName))
            {
                throw new InvalidOperationException($"Engine is not enabled: {engineName}");
            }
            
            var engineConfig = _configManager.GetEngineConfiguration(engineName);
            
            if (!engine.ValidateConfiguration(engineConfig))
            {
                throw new InvalidOperationException($"Engine configuration is invalid: {engineName}");
            }
            
            return await engine.SynthesizeSpeechAsync(text, voiceId, engineConfig);
        }
    }
} 