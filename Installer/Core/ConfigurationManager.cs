using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Microsoft.Win32;

namespace Installer.Core
{
    /// <summary>
    /// Manages TTS engine configurations
    /// </summary>
    public class ConfigurationManager
    {
        private const string RegistryKey = @"SOFTWARE\OpenSpeech\TTS";
        private const string ConfigPathValue = "ConfigPath";
        private const string EncryptionKey = "OpenSpeechTTSConfigEncryptionKey";
        
        private string _configPath;
        private JObject _config;
        
        /// <summary>
        /// Initializes a new instance of the ConfigurationManager class
        /// </summary>
        public ConfigurationManager()
        {
            LoadConfigPath();
            LoadConfiguration();
        }
        
        /// <summary>
        /// Gets the configuration for a specific engine
        /// </summary>
        /// <param name="engineName">The name of the engine</param>
        /// <returns>The engine configuration</returns>
        public Dictionary<string, string> GetEngineConfiguration(string engineName)
        {
            if (_config == null)
            {
                return new Dictionary<string, string>();
            }
            
            var engines = _config["engines"] as JObject;
            if (engines == null || !engines.ContainsKey(engineName))
            {
                return new Dictionary<string, string>();
            }
            
            var engineConfig = engines[engineName] as JObject;
            if (engineConfig == null || !engineConfig.ContainsKey("parameters"))
            {
                return new Dictionary<string, string>();
            }
            
            var parameters = engineConfig["parameters"] as JObject;
            if (parameters == null)
            {
                return new Dictionary<string, string>();
            }
            
            var result = new Dictionary<string, string>();
            
            foreach (var property in parameters.Properties())
            {
                string value = property.Value.ToString();
                
                // Decrypt sensitive values
                if (IsSensitiveParameter(property.Name))
                {
                    value = DecryptValue(value);
                }
                
                result[property.Name] = value;
            }
            
            return result;
        }
        
        /// <summary>
        /// Updates the configuration for a specific engine
        /// </summary>
        /// <param name="engineName">The name of the engine</param>
        /// <param name="configuration">The engine configuration</param>
        public void UpdateEngineConfiguration(string engineName, Dictionary<string, string> configuration)
        {
            if (_config == null)
            {
                _config = new JObject();
            }
            
            if (!_config.ContainsKey("engines"))
            {
                _config["engines"] = new JObject();
            }
            
            var engines = _config["engines"] as JObject;
            
            if (!engines.ContainsKey(engineName))
            {
                engines[engineName] = new JObject();
                engines[engineName]["enabled"] = false;
            }
            
            var engineConfig = engines[engineName] as JObject;
            
            if (!engineConfig.ContainsKey("parameters"))
            {
                engineConfig["parameters"] = new JObject();
            }
            
            var parameters = engineConfig["parameters"] as JObject;
            
            foreach (var kvp in configuration)
            {
                string value = kvp.Value;
                
                // Encrypt sensitive values
                if (IsSensitiveParameter(kvp.Key))
                {
                    value = EncryptValue(value);
                }
                
                parameters[kvp.Key] = value;
            }
            
            // Update last updated timestamp
            _config["lastUpdated"] = DateTime.Now.ToString("o");
            
            SaveConfiguration();
        }
        
        /// <summary>
        /// Enables or disables an engine
        /// </summary>
        /// <param name="engineName">The name of the engine</param>
        /// <param name="enabled">Whether the engine is enabled</param>
        public void SetEngineEnabled(string engineName, bool enabled)
        {
            if (_config == null)
            {
                _config = new JObject();
            }
            
            if (!_config.ContainsKey("engines"))
            {
                _config["engines"] = new JObject();
            }
            
            var engines = _config["engines"] as JObject;
            
            if (!engines.ContainsKey(engineName))
            {
                engines[engineName] = new JObject();
                engines[engineName]["parameters"] = new JObject();
            }
            
            var engineConfig = engines[engineName] as JObject;
            engineConfig["enabled"] = enabled;
            
            // Update last updated timestamp
            _config["lastUpdated"] = DateTime.Now.ToString("o");
            
            SaveConfiguration();
        }
        
        /// <summary>
        /// Gets whether an engine is enabled
        /// </summary>
        /// <param name="engineName">The name of the engine</param>
        /// <returns>Whether the engine is enabled</returns>
        public bool IsEngineEnabled(string engineName)
        {
            if (_config == null)
            {
                return false;
            }
            
            var engines = _config["engines"] as JObject;
            if (engines == null || !engines.ContainsKey(engineName))
            {
                return false;
            }
            
            var engineConfig = engines[engineName] as JObject;
            if (engineConfig == null || !engineConfig.ContainsKey("enabled"))
            {
                return false;
            }
            
            return engineConfig["enabled"].Value<bool>();
        }
        
        /// <summary>
        /// Gets the default engine
        /// </summary>
        /// <returns>The default engine name</returns>
        public string GetDefaultEngine()
        {
            if (_config == null || !_config.ContainsKey("defaultEngine"))
            {
                return "SherpaOnnx";
            }
            
            return _config["defaultEngine"].Value<string>();
        }
        
        /// <summary>
        /// Sets the default engine
        /// </summary>
        /// <param name="engineName">The name of the engine</param>
        public void SetDefaultEngine(string engineName)
        {
            if (_config == null)
            {
                _config = new JObject();
            }
            
            _config["defaultEngine"] = engineName;
            
            // Update last updated timestamp
            _config["lastUpdated"] = DateTime.Now.ToString("o");
            
            SaveConfiguration();
        }
        
        /// <summary>
        /// Gets the plugin directory
        /// </summary>
        /// <returns>The plugin directory</returns>
        public string GetPluginDirectory()
        {
            if (_config == null || !_config.ContainsKey("pluginDirectory"))
            {
                return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "OpenSpeech", "plugins");
            }
            
            return _config["pluginDirectory"].Value<string>();
        }
        
        /// <summary>
        /// Sets the plugin directory
        /// </summary>
        /// <param name="directory">The plugin directory</param>
        public void SetPluginDirectory(string directory)
        {
            if (_config == null)
            {
                _config = new JObject();
            }
            
            _config["pluginDirectory"] = directory;
            
            // Update last updated timestamp
            _config["lastUpdated"] = DateTime.Now.ToString("o");
            
            SaveConfiguration();
        }
        
        private void LoadConfigPath()
        {
            try
            {
                using (var key = Registry.LocalMachine.OpenSubKey(RegistryKey))
                {
                    if (key != null)
                    {
                        _configPath = key.GetValue(ConfigPathValue) as string;
                    }
                }
                
                if (string.IsNullOrEmpty(_configPath))
                {
                    _configPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "OpenSpeech", "config", "engine-config.json");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading config path: {ex.Message}");
                _configPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "OpenSpeech", "config", "engine-config.json");
            }
        }
        
        private void LoadConfiguration()
        {
            try
            {
                if (File.Exists(_configPath))
                {
                    string json = File.ReadAllText(_configPath);
                    _config = JObject.Parse(json);
                }
                else
                {
                    _config = new JObject();
                    _config["version"] = "1.1.0";
                    _config["engines"] = new JObject();
                    _config["defaultEngine"] = "SherpaOnnx";
                    _config["pluginDirectory"] = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "OpenSpeech", "plugins");
                    _config["lastUpdated"] = DateTime.Now.ToString("o");
                    
                    SaveConfiguration();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading configuration: {ex.Message}");
                _config = new JObject();
            }
        }
        
        private void SaveConfiguration()
        {
            try
            {
                string directory = Path.GetDirectoryName(_configPath);
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }
                
                string json = _config.ToString(Formatting.Indented);
                File.WriteAllText(_configPath, json);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving configuration: {ex.Message}");
            }
        }
        
        private bool IsSensitiveParameter(string parameterName)
        {
            // List of parameter names that should be encrypted
            return parameterName.Contains("key", StringComparison.OrdinalIgnoreCase) ||
                   parameterName.Contains("secret", StringComparison.OrdinalIgnoreCase) ||
                   parameterName.Contains("password", StringComparison.OrdinalIgnoreCase) ||
                   parameterName.Contains("token", StringComparison.OrdinalIgnoreCase);
        }
        
        private string EncryptValue(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return value;
            }
            
            try
            {
                byte[] encryptedData = ProtectedData.Protect(
                    Encoding.UTF8.GetBytes(value),
                    Encoding.UTF8.GetBytes(EncryptionKey),
                    DataProtectionScope.LocalMachine);
                
                return Convert.ToBase64String(encryptedData);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error encrypting value: {ex.Message}");
                return value;
            }
        }
        
        private string DecryptValue(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return value;
            }
            
            try
            {
                byte[] decryptedData = ProtectedData.Unprotect(
                    Convert.FromBase64String(value),
                    Encoding.UTF8.GetBytes(EncryptionKey),
                    DataProtectionScope.LocalMachine);
                
                return Encoding.UTF8.GetString(decryptedData);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error decrypting value: {ex.Message}");
                return value;
            }
        }
    }
} 