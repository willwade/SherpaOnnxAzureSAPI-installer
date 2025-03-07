using System;
using System.Collections.Generic;
using System.Linq;
using Installer.Core.Interfaces;

namespace Installer.Core.Managers
{
    /// <summary>
    /// Manages TTS engines
    /// </summary>
    public class TtsEngineManager
    {
        private readonly Dictionary<string, ITtsEngine> _engines = new Dictionary<string, ITtsEngine>();
        private readonly ConfigurationManager _configManager;
        
        /// <summary>
        /// Creates a new instance of the TtsEngineManager class
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
                throw new ArgumentException("Engine name cannot be null or empty");
            }
            
            if (_engines.ContainsKey(engine.EngineName))
            {
                throw new ArgumentException($"Engine with name '{engine.EngineName}' is already registered");
            }
            
            _engines[engine.EngineName] = engine;
            engine.Initialize();
        }
        
        /// <summary>
        /// Unregisters a TTS engine
        /// </summary>
        /// <param name="engineName">The name of the engine to unregister</param>
        public void UnregisterEngine(string engineName)
        {
            if (string.IsNullOrEmpty(engineName))
            {
                throw new ArgumentNullException(nameof(engineName));
            }
            
            if (!_engines.ContainsKey(engineName))
            {
                throw new ArgumentException($"Engine with name '{engineName}' is not registered");
            }
            
            _engines[engineName].Shutdown();
            _engines.Remove(engineName);
        }
        
        /// <summary>
        /// Gets a TTS engine by name
        /// </summary>
        /// <param name="engineName">The name of the engine to get</param>
        /// <returns>The TTS engine</returns>
        public ITtsEngine GetEngine(string engineName)
        {
            if (string.IsNullOrEmpty(engineName))
            {
                throw new ArgumentNullException(nameof(engineName));
            }
            
            if (!_engines.ContainsKey(engineName))
            {
                throw new ArgumentException($"Engine with name '{engineName}' is not registered");
            }
            
            return _engines[engineName];
        }
        
        /// <summary>
        /// Gets all registered TTS engines
        /// </summary>
        /// <returns>The registered TTS engines</returns>
        public IEnumerable<ITtsEngine> GetAllEngines()
        {
            return _engines.Values;
        }
        
        /// <summary>
        /// Gets the names of all registered TTS engines
        /// </summary>
        /// <returns>The names of the registered TTS engines</returns>
        public IEnumerable<string> GetEngineNames()
        {
            return _engines.Keys;
        }
        
        /// <summary>
        /// Gets the default TTS engine
        /// </summary>
        /// <returns>The default TTS engine</returns>
        public ITtsEngine GetDefaultEngine()
        {
            var config = _configManager.LoadConfiguration();
            string defaultEngineName = config.DefaultEngine;
            
            if (string.IsNullOrEmpty(defaultEngineName) || !_engines.ContainsKey(defaultEngineName))
            {
                // If the default engine is not set or not registered, use the first registered engine
                defaultEngineName = _engines.Keys.FirstOrDefault();
                
                if (string.IsNullOrEmpty(defaultEngineName))
                {
                    throw new InvalidOperationException("No engines are registered");
                }
            }
            
            return _engines[defaultEngineName];
        }
        
        /// <summary>
        /// Sets the default TTS engine
        /// </summary>
        /// <param name="engineName">The name of the engine to set as default</param>
        public void SetDefaultEngine(string engineName)
        {
            if (string.IsNullOrEmpty(engineName))
            {
                throw new ArgumentNullException(nameof(engineName));
            }
            
            if (!_engines.ContainsKey(engineName))
            {
                throw new ArgumentException($"Engine with name '{engineName}' is not registered");
            }
            
            var config = _configManager.LoadConfiguration();
            config.DefaultEngine = engineName;
            _configManager.SaveConfiguration(config);
        }
        
        /// <summary>
        /// Checks if an engine is registered
        /// </summary>
        /// <param name="engineName">The name of the engine to check</param>
        /// <returns>True if the engine is registered, false otherwise</returns>
        public bool IsEngineRegistered(string engineName)
        {
            return !string.IsNullOrEmpty(engineName) && _engines.ContainsKey(engineName);
        }
        
        /// <summary>
        /// Shuts down all registered engines
        /// </summary>
        public void ShutdownAllEngines()
        {
            foreach (var engine in _engines.Values)
            {
                try
                {
                    engine.Shutdown();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error shutting down engine '{engine.EngineName}': {ex.Message}");
                }
            }
            
            _engines.Clear();
        }
    }
} 