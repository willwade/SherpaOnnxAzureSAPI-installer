using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using OpenSpeech.TTS.Plugins;

namespace Installer.Core
{
    /// <summary>
    /// Loads TTS engine plugins
    /// </summary>
    public class PluginLoader
    {
        private readonly string _pluginDirectory;
        private readonly TtsEngineManager _engineManager;
        
        /// <summary>
        /// Initializes a new instance of the PluginLoader class
        /// </summary>
        /// <param name="pluginDirectory">The directory containing plugins</param>
        /// <param name="engineManager">The engine manager</param>
        public PluginLoader(string pluginDirectory, TtsEngineManager engineManager)
        {
            _pluginDirectory = pluginDirectory;
            _engineManager = engineManager;
        }
        
        /// <summary>
        /// Loads all plugins from the plugin directory
        /// </summary>
        public void LoadAllPlugins()
        {
            if (!Directory.Exists(_pluginDirectory))
            {
                Console.WriteLine($"Plugin directory not found: {_pluginDirectory}");
                return;
            }
            
            // Get all DLL files in the plugin directory
            string[] dllFiles = Directory.GetFiles(_pluginDirectory, "*.dll");
            
            foreach (string dllFile in dllFiles)
            {
                try
                {
                    LoadPlugin(dllFile);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error loading plugin {dllFile}: {ex.Message}");
                }
            }
        }
        
        /// <summary>
        /// Loads a plugin from a DLL file
        /// </summary>
        /// <param name="dllPath">The path to the DLL file</param>
        public void LoadPlugin(string dllPath)
        {
            try
            {
                // Load the assembly
                Assembly assembly = Assembly.LoadFrom(dllPath);
                
                // Find all types that implement ITtsEngine
                foreach (Type type in assembly.GetTypes())
                {
                    if (typeof(ITtsEngine).IsAssignableFrom(type) && !type.IsAbstract)
                    {
                        // Create an instance of the engine
                        ITtsEngine engine = (ITtsEngine)Activator.CreateInstance(type);
                        
                        // Register the engine
                        _engineManager.RegisterEngine(engine);
                        
                        Console.WriteLine($"Loaded plugin: {engine.EngineName} ({engine.EngineVersion})");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading plugin {dllPath}: {ex.Message}");
                throw;
            }
        }
    }
} 