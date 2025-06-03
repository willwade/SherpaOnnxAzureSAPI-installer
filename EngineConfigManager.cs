using System;
using System.IO;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Collections.Generic;

namespace SherpaOnnxSAPIInstaller
{
    /// <summary>
    /// Manages the engines_config.json file for TTS engine configuration
    /// This class handles adding/removing Azure and SherpaOnnx voice configurations
    /// </summary>
    public static class EngineConfigManager
    {
        private static readonly string ConfigPath = @"C:\Program Files\OpenAssistive\OpenSpeech\engines_config.json";
        
        /// <summary>
        /// Adds an Azure TTS voice configuration to engines_config.json
        /// </summary>
        /// <param name="voiceName">Azure voice name (e.g., "en-GB-ElliotNeural")</param>
        /// <param name="subscriptionKey">Azure subscription key</param>
        /// <param name="region">Azure region</param>
        /// <param name="language">Language code (e.g., "en-GB")</param>
        /// <param name="sapiVoiceName">Full SAPI voice name for mapping</param>
        public static void AddAzureVoice(string voiceName, string subscriptionKey, string region, 
            string language, string sapiVoiceName)
        {
            try
            {
                Console.WriteLine($"Adding Azure voice configuration: {voiceName}");
                
                // Load existing configuration
                var config = LoadConfiguration();
                
                // Create engine ID (e.g., "azure-elliot")
                string engineId = GenerateAzureEngineId(voiceName);
                
                // Create engine configuration
                var engineConfig = new JsonObject
                {
                    ["type"] = "azure",
                    ["config"] = new JsonObject
                    {
                        ["subscriptionKey"] = subscriptionKey,
                        ["region"] = region,
                        ["voiceName"] = voiceName,
                        ["language"] = language,
                        ["sampleRate"] = 24000,
                        ["channels"] = 1,
                        ["bitsPerSample"] = 16
                    }
                };
                
                // Add to engines section
                var engines = config["engines"]?.AsObject();
                if (engines != null)
                {
                    engines[engineId] = engineConfig;
                    Console.WriteLine($"Added engine: {engineId}");
                }
                
                // Add to voices section
                var voices = config["voices"]?.AsObject();
                if (voices != null)
                {
                    // Add short name mapping (e.g., "elliot" -> "azure-elliot")
                    string shortName = GenerateShortName(voiceName);
                    voices[shortName] = engineId;
                    Console.WriteLine($"Added voice mapping: {shortName} -> {engineId}");
                    
                    // Add SAPI voice name mapping
                    if (!string.IsNullOrEmpty(sapiVoiceName))
                    {
                        voices[sapiVoiceName] = engineId;
                        Console.WriteLine($"Added SAPI mapping: {sapiVoiceName} -> {engineId}");
                    }
                }
                
                // Save configuration
                SaveConfiguration(config);
                Console.WriteLine($"Azure voice '{voiceName}' configuration added successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error adding Azure voice configuration: {ex.Message}");
                throw;
            }
        }
        
        /// <summary>
        /// Removes an Azure TTS voice configuration from engines_config.json
        /// </summary>
        /// <param name="voiceName">Azure voice name to remove</param>
        public static void RemoveAzureVoice(string voiceName)
        {
            try
            {
                Console.WriteLine($"Removing Azure voice configuration: {voiceName}");
                
                var config = LoadConfiguration();
                string engineId = GenerateAzureEngineId(voiceName);
                
                // Remove from engines
                var engines = config["engines"]?.AsObject();
                if (engines != null && engines.ContainsKey(engineId))
                {
                    engines.Remove(engineId);
                    Console.WriteLine($"Removed engine: {engineId}");
                }
                
                // Remove from voices (find all mappings to this engine)
                var voices = config["voices"]?.AsObject();
                if (voices != null)
                {
                    var toRemove = new List<string>();
                    foreach (var kvp in voices)
                    {
                        if (kvp.Value?.ToString() == engineId)
                        {
                            toRemove.Add(kvp.Key);
                        }
                    }
                    
                    foreach (var key in toRemove)
                    {
                        voices.Remove(key);
                        Console.WriteLine($"Removed voice mapping: {key}");
                    }
                }
                
                SaveConfiguration(config);
                Console.WriteLine($"Azure voice '{voiceName}' configuration removed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error removing Azure voice configuration: {ex.Message}");
                throw;
            }
        }
        
        /// <summary>
        /// Updates Azure credentials for all existing Azure voices
        /// </summary>
        /// <param name="subscriptionKey">New subscription key</param>
        /// <param name="region">New region</param>
        public static void UpdateAzureCredentials(string subscriptionKey, string region)
        {
            try
            {
                Console.WriteLine("Updating Azure credentials for all voices...");
                
                var config = LoadConfiguration();
                var engines = config["engines"]?.AsObject();
                
                if (engines != null)
                {
                    int updatedCount = 0;
                    foreach (var kvp in engines)
                    {
                        var engine = kvp.Value?.AsObject();
                        if (engine != null && engine["type"]?.ToString() == "azure")
                        {
                            var engineConfig = engine["config"]?.AsObject();
                            if (engineConfig != null)
                            {
                                engineConfig["subscriptionKey"] = subscriptionKey;
                                engineConfig["region"] = region;
                                updatedCount++;
                                Console.WriteLine($"Updated credentials for: {kvp.Key}");
                            }
                        }
                    }
                    
                    if (updatedCount > 0)
                    {
                        SaveConfiguration(config);
                        Console.WriteLine($"Updated credentials for {updatedCount} Azure voices.");
                    }
                    else
                    {
                        Console.WriteLine("No Azure voices found to update.");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error updating Azure credentials: {ex.Message}");
                throw;
            }
        }
        
        /// <summary>
        /// Checks if the engines_config.json file exists and is valid
        /// </summary>
        /// <returns>True if configuration is valid</returns>
        public static bool IsConfigurationValid()
        {
            try
            {
                if (!File.Exists(ConfigPath))
                {
                    Console.WriteLine($"Configuration file not found: {ConfigPath}");
                    return false;
                }
                
                var config = LoadConfiguration();
                return config["engines"] != null && config["voices"] != null;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Configuration validation error: {ex.Message}");
                return false;
            }
        }
        
        private static JsonNode LoadConfiguration()
        {
            if (!File.Exists(ConfigPath))
            {
                throw new FileNotFoundException($"Configuration file not found: {ConfigPath}");
            }
            
            string json = File.ReadAllText(ConfigPath);
            return JsonNode.Parse(json) ?? throw new InvalidOperationException("Failed to parse configuration file");
        }
        
        private static void SaveConfiguration(JsonNode config)
        {
            // Create backup
            string backupPath = ConfigPath + ".backup";
            if (File.Exists(ConfigPath))
            {
                File.Copy(ConfigPath, backupPath, true);
            }
            
            // Save with pretty formatting
            var options = new JsonSerializerOptions
            {
                WriteIndented = true
            };
            
            string json = config.ToJsonString(options);
            File.WriteAllText(ConfigPath, json);
        }
        
        private static string GenerateAzureEngineId(string voiceName)
        {
            // Convert "en-GB-ElliotNeural" to "azure-elliot"
            string[] parts = voiceName.Split('-');
            if (parts.Length >= 3)
            {
                string name = parts[2].Replace("Neural", "").ToLower();
                return $"azure-{name}";
            }
            
            // Fallback
            return $"azure-{voiceName.ToLower().Replace("-", "").Replace("neural", "")}";
        }
        
        private static string GenerateShortName(string voiceName)
        {
            // Convert "en-GB-ElliotNeural" to "elliot"
            string[] parts = voiceName.Split('-');
            if (parts.Length >= 3)
            {
                return parts[2].Replace("Neural", "").ToLower();
            }
            
            // Fallback
            return voiceName.ToLower().Replace("-", "").Replace("neural", "");
        }
    }
}
