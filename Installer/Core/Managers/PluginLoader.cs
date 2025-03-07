using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Installer.Core.Interfaces;

namespace Installer.Core.Managers
{
    /// <summary>
    /// Loads TTS engine plugins
    /// </summary>
    public class PluginLoader
    {
        private readonly string _pluginDirectory;
        private readonly TtsEngineManager _engineManager;
        
        /// <summary>
        /// Creates a new instance of the PluginLoader class
        /// </summary>
        /// <param name="pluginDirectory">The directory containing the plugins</param>
        /// <param name="engineManager">The engine manager</param>
        public PluginLoader(string pluginDirectory, TtsEngineManager engineManager)
        {
            _pluginDirectory = pluginDirectory;
            _engineManager = engineManager;
        }
        
        /// <summary>
        /// Discovers and loads TTS engines from plugins
        /// </summary>
        /// <returns>The discovered engines</returns>
        public IEnumerable<ITtsEngine> DiscoverEngines()
        {
            var engines = new List<ITtsEngine>();
            
            // Load external engines from plugin directory
            if (Directory.Exists(_pluginDirectory))
            {
                foreach (var file in Directory.GetFiles(_pluginDirectory, "*.dll"))
                {
                    try
                    {
                        Console.WriteLine($"Loading plugin: {file}");
                        var assembly = Assembly.LoadFrom(file);
                        
                        foreach (var type in assembly.GetTypes()
                            .Where(t => typeof(ITtsEngine).IsAssignableFrom(t) && !t.IsAbstract))
                        {
                            try
                            {
                                var engine = (ITtsEngine)Activator.CreateInstance(type);
                                engines.Add(engine);
                                Console.WriteLine($"Discovered engine: {engine.EngineName} ({engine.EngineVersion})");
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Error creating engine instance from type {type.FullName}: {ex.Message}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error loading plugin {file}: {ex.Message}");
                    }
                }
            }
            else
            {
                Console.WriteLine($"Plugin directory not found: {_pluginDirectory}");
            }
            
            return engines;
        }
        
        /// <summary>
        /// Loads all discovered engines into the engine manager
        /// </summary>
        public void LoadAllEngines()
        {
            var engines = DiscoverEngines();
            
            foreach (var engine in engines)
            {
                try
                {
                    _engineManager.RegisterEngine(engine);
                    Console.WriteLine($"Registered engine: {engine.EngineName}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error registering engine {engine.EngineName}: {ex.Message}");
                }
            }
        }
        
        /// <summary>
        /// Loads engines from subdirectories of the plugin directory
        /// </summary>
        public void LoadEnginesFromSubdirectories()
        {
            if (!Directory.Exists(_pluginDirectory))
            {
                Console.WriteLine($"Plugin directory not found: {_pluginDirectory}");
                return;
            }
            
            foreach (var directory in Directory.GetDirectories(_pluginDirectory))
            {
                try
                {
                    var directoryName = new DirectoryInfo(directory).Name;
                    Console.WriteLine($"Loading plugins from directory: {directoryName}");
                    
                    foreach (var file in Directory.GetFiles(directory, "*.dll"))
                    {
                        try
                        {
                            Console.WriteLine($"Loading plugin: {file}");
                            var assembly = Assembly.LoadFrom(file);
                            
                            foreach (var type in assembly.GetTypes()
                                .Where(t => typeof(ITtsEngine).IsAssignableFrom(t) && !t.IsAbstract))
                            {
                                try
                                {
                                    var engine = (ITtsEngine)Activator.CreateInstance(type);
                                    _engineManager.RegisterEngine(engine);
                                    Console.WriteLine($"Registered engine: {engine.EngineName} ({engine.EngineVersion})");
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"Error creating or registering engine instance from type {type.FullName}: {ex.Message}");
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error loading plugin {file}: {ex.Message}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing directory {directory}: {ex.Message}");
                }
            }
        }
    }
} 