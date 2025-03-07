using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using Installer.Core.Models;
using Newtonsoft.Json;

namespace Installer.Core.Managers
{
    /// <summary>
    /// Manages configuration for all TTS engines
    /// </summary>
    public class ConfigurationManager
    {
        private readonly string _configDir;
        private readonly string _configFile;
        private EngineConfiguration _configuration;
        
        /// <summary>
        /// Creates a new instance of the ConfigurationManager class
        /// </summary>
        public ConfigurationManager()
        {
            _configDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OpenSpeech");
            _configFile = Path.Combine(_configDir, "config.json");
        }
        
        /// <summary>
        /// Creates a new instance of the ConfigurationManager class with a custom configuration directory
        /// </summary>
        /// <param name="configDir">The configuration directory</param>
        public ConfigurationManager(string configDir)
        {
            _configDir = configDir;
            _configFile = Path.Combine(_configDir, "config.json");
        }
        
        /// <summary>
        /// Loads the configuration from the configuration file
        /// </summary>
        /// <returns>The engine configuration</returns>
        public EngineConfiguration LoadConfiguration()
        {
            try
            {
                // Create configuration directory if it doesn't exist
                if (!Directory.Exists(_configDir))
                {
                    Directory.CreateDirectory(_configDir);
                }
                
                // Create default configuration if the file doesn't exist
                if (!File.Exists(_configFile))
                {
                    _configuration = CreateDefaultConfiguration();
                    SaveConfiguration(_configuration);
                    return _configuration;
                }
                
                // Load configuration from file
                string json = File.ReadAllText(_configFile);
                _configuration = JsonConvert.DeserializeObject<EngineConfiguration>(json);
                
                // Decrypt sensitive values
                if (_configuration.SecureStorage)
                {
                    foreach (var engine in _configuration.Engines)
                    {
                        foreach (var param in engine.Value.Parameters)
                        {
                            if (param.Value.StartsWith("ENCRYPTED:"))
                            {
                                engine.Value.Parameters[param.Key] = DecryptValue(param.Value.Substring(10));
                            }
                        }
                    }
                }
                
                return _configuration;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading configuration: {ex.Message}");
                _configuration = CreateDefaultConfiguration();
                return _configuration;
            }
        }
        
        /// <summary>
        /// Saves the configuration to the configuration file
        /// </summary>
        /// <param name="config">The engine configuration</param>
        public void SaveConfiguration(EngineConfiguration config)
        {
            try
            {
                // Create configuration directory if it doesn't exist
                if (!Directory.Exists(_configDir))
                {
                    Directory.CreateDirectory(_configDir);
                }
                
                // Create a copy of the configuration to avoid modifying the original
                var configCopy = JsonConvert.DeserializeObject<EngineConfiguration>(JsonConvert.SerializeObject(config));
                
                // Encrypt sensitive values
                if (configCopy.SecureStorage)
                {
                    foreach (var engine in configCopy.Engines)
                    {
                        var engineName = engine.Key;
                        var parameters = GetRequiredParameters(engineName);
                        
                        foreach (var param in parameters)
                        {
                            if (param.IsSecret && engine.Value.Parameters.ContainsKey(param.Name))
                            {
                                string value = engine.Value.Parameters[param.Name];
                                if (!string.IsNullOrEmpty(value) && !value.StartsWith("ENCRYPTED:"))
                                {
                                    engine.Value.Parameters[param.Name] = "ENCRYPTED:" + EncryptValue(value);
                                }
                            }
                        }
                    }
                }
                
                // Update last updated timestamp
                configCopy.LastUpdated = DateTime.UtcNow;
                
                // Save configuration to file
                string json = JsonConvert.SerializeObject(configCopy, Formatting.Indented);
                File.WriteAllText(_configFile, json);
                
                // Update the current configuration
                _configuration = config;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving configuration: {ex.Message}");
            }
        }
        
        /// <summary>
        /// Gets the configuration for a specific engine
        /// </summary>
        /// <param name="engineName">The name of the engine</param>
        /// <returns>The engine configuration parameters</returns>
        public Dictionary<string, string> GetEngineConfiguration(string engineName)
        {
            if (_configuration == null)
            {
                LoadConfiguration();
            }
            
            if (!_configuration.Engines.ContainsKey(engineName))
            {
                return new Dictionary<string, string>();
            }
            
            // Create a copy of the parameters to avoid modifying the original
            var parameters = new Dictionary<string, string>(_configuration.Engines[engineName].Parameters);
            
            // Resolve environment variables
            foreach (var param in parameters.Keys.ToArray())
            {
                parameters[param] = ResolveEnvironmentVariables(parameters[param]);
            }
            
            return parameters;
        }
        
        /// <summary>
        /// Updates the configuration for a specific engine
        /// </summary>
        /// <param name="engineName">The name of the engine</param>
        /// <param name="parameters">The engine configuration parameters</param>
        public void UpdateEngineConfiguration(string engineName, Dictionary<string, string> parameters)
        {
            if (_configuration == null)
            {
                LoadConfiguration();
            }
            
            if (!_configuration.Engines.ContainsKey(engineName))
            {
                _configuration.Engines[engineName] = new EngineConfigEntry();
            }
            
            // Update parameters
            foreach (var param in parameters)
            {
                _configuration.Engines[engineName].Parameters[param.Key] = param.Value;
            }
            
            // Save configuration
            SaveConfiguration(_configuration);
        }
        
        /// <summary>
        /// Encrypts a value using DPAPI
        /// </summary>
        /// <param name="value">The value to encrypt</param>
        /// <returns>The encrypted value</returns>
        public string EncryptValue(string value)
        {
            try
            {
                byte[] data = Encoding.UTF8.GetBytes(value);
                byte[] encrypted = ProtectedData.Protect(data, null, DataProtectionScope.CurrentUser);
                return Convert.ToBase64String(encrypted);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error encrypting value: {ex.Message}");
                return value;
            }
        }
        
        /// <summary>
        /// Decrypts a value using DPAPI
        /// </summary>
        /// <param name="encryptedValue">The encrypted value</param>
        /// <returns>The decrypted value</returns>
        public string DecryptValue(string encryptedValue)
        {
            try
            {
                byte[] data = Convert.FromBase64String(encryptedValue);
                byte[] decrypted = ProtectedData.Unprotect(data, null, DataProtectionScope.CurrentUser);
                return Encoding.UTF8.GetString(decrypted);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error decrypting value: {ex.Message}");
                return encryptedValue;
            }
        }
        
        /// <summary>
        /// Resolves environment variables in a value
        /// </summary>
        /// <param name="value">The value to resolve</param>
        /// <returns>The resolved value</returns>
        public string ResolveEnvironmentVariables(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return value;
            }
            
            // Replace environment variables in the format %VARIABLE%
            return Regex.Replace(value, @"%([^%]+)%", match =>
            {
                string envVar = match.Groups[1].Value;
                string envValue = Environment.GetEnvironmentVariable(envVar);
                return envValue ?? match.Value;
            });
        }
        
        /// <summary>
        /// Creates a default configuration
        /// </summary>
        /// <returns>The default configuration</returns>
        private EngineConfiguration CreateDefaultConfiguration()
        {
            return new EngineConfiguration
            {
                DefaultEngine = "SherpaOnnx",
                SecureStorage = true,
                Engines = new Dictionary<string, EngineConfigEntry>
                {
                    ["SherpaOnnx"] = new EngineConfigEntry
                    {
                        Enabled = true,
                        Parameters = new Dictionary<string, string>()
                    },
                    ["AzureTTS"] = new EngineConfigEntry
                    {
                        Enabled = true,
                        Parameters = new Dictionary<string, string>
                        {
                            ["subscriptionKey"] = "",
                            ["region"] = "eastus"
                        }
                    }
                }
            };
        }
        
        /// <summary>
        /// Gets the required parameters for a specific engine
        /// </summary>
        /// <param name="engineName">The name of the engine</param>
        /// <returns>The required parameters</returns>
        private IEnumerable<ConfigurationParameter> GetRequiredParameters(string engineName)
        {
            // This is a temporary implementation - in the future, this should be retrieved from the engine
            switch (engineName)
            {
                case "SherpaOnnx":
                    return new List<ConfigurationParameter>();
                
                case "AzureTTS":
                    return new List<ConfigurationParameter>
                    {
                        new ConfigurationParameter
                        {
                            Name = "subscriptionKey",
                            DisplayName = "Subscription Key",
                            Description = "Azure Cognitive Services subscription key",
                            IsSecret = true
                        },
                        new ConfigurationParameter
                        {
                            Name = "region",
                            DisplayName = "Region",
                            Description = "Azure region (e.g., eastus, westus)"
                        }
                    };
                
                case "ElevenLabs":
                    return new List<ConfigurationParameter>
                    {
                        new ConfigurationParameter
                        {
                            Name = "apiKey",
                            DisplayName = "API Key",
                            Description = "ElevenLabs API key",
                            IsSecret = true
                        },
                        new ConfigurationParameter
                        {
                            Name = "modelId",
                            DisplayName = "Model ID",
                            Description = "ElevenLabs model ID (optional)",
                            IsRequired = false,
                            DefaultValue = "eleven_monolingual_v1"
                        }
                    };
                
                case "PlayHT":
                    return new List<ConfigurationParameter>
                    {
                        new ConfigurationParameter
                        {
                            Name = "apiKey",
                            DisplayName = "API Key",
                            Description = "PlayHT API key",
                            IsSecret = true
                        },
                        new ConfigurationParameter
                        {
                            Name = "userId",
                            DisplayName = "User ID",
                            Description = "PlayHT user ID"
                        }
                    };
                
                default:
                    return new List<ConfigurationParameter>();
            }
        }
    }
} 