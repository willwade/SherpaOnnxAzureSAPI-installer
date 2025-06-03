using System;
using System.IO;
using System.Text;
using System.Security.Cryptography;
using Newtonsoft.Json;
using Installer.Shared;

namespace Installer
{
    public class AzureConfigManager
    {
        private const string CONFIG_FILENAME = "azure_config.json";
        
        public static string GetConfigFilePath()
        {
            string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string configDir = Path.Combine(appDataPath, "OpenSpeech");
            
            // Ensure directory exists
            if (!Directory.Exists(configDir))
            {
                Directory.CreateDirectory(configDir);
            }
            
            return Path.Combine(configDir, CONFIG_FILENAME);
        }
        
        public static AzureConfig LoadConfig()
        {
            string configPath = GetConfigFilePath();
            
            if (!File.Exists(configPath))
            {
                return new AzureConfig();
            }
            
            try
            {
                string json = File.ReadAllText(configPath);
                var config = JsonConvert.DeserializeObject<AzureConfig>(json);
                
                // Decrypt the key if secure storage is enabled
                if (config.SecureStorage && !string.IsNullOrEmpty(config.DefaultKey))
                {
                    try
                    {
                        config.DefaultKey = DecryptKey(config.DefaultKey);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Warning: Could not decrypt the stored key: {ex.Message}");
                        config.DefaultKey = null;
                    }
                }
                
                return config;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading Azure configuration: {ex.Message}");
                return new AzureConfig();
            }
        }
        
        public static void SaveConfig(string key, string region, bool secureStorage = true)
        {
            try
            {
                var config = new AzureConfig
                {
                    DefaultRegion = region,
                    SecureStorage = secureStorage
                };
                
                // Encrypt the key if secure storage is enabled
                config.DefaultKey = secureStorage ? EncryptKey(key) : key;
                
                string json = JsonConvert.SerializeObject(config, Formatting.Indented);
                File.WriteAllText(GetConfigFilePath(), json);
                
                Console.WriteLine("Azure configuration saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving Azure configuration: {ex.Message}");
            }
        }
        
        private static string EncryptKey(string key)
        {
            if (string.IsNullOrEmpty(key))
                return key;
                
            try
            {
                // Use DPAPI for Windows
                byte[] keyBytes = Encoding.UTF8.GetBytes(key);
                byte[] encryptedBytes = ProtectedData.Protect(keyBytes, null, DataProtectionScope.CurrentUser);
                return Convert.ToBase64String(encryptedBytes);
            }
            catch (Exception)
            {
                // Fallback to simple encoding if DPAPI is not available
                return Convert.ToBase64String(Encoding.UTF8.GetBytes(key));
            }
        }
        
        private static string DecryptKey(string encryptedKey)
        {
            if (string.IsNullOrEmpty(encryptedKey))
                return encryptedKey;
                
            try
            {
                // Use DPAPI for Windows
                byte[] encryptedBytes = Convert.FromBase64String(encryptedKey);
                byte[] keyBytes = ProtectedData.Unprotect(encryptedBytes, null, DataProtectionScope.CurrentUser);
                return Encoding.UTF8.GetString(keyBytes);
            }
            catch (Exception)
            {
                // Fallback to simple decoding if DPAPI is not available
                return Encoding.UTF8.GetString(Convert.FromBase64String(encryptedKey));
            }
        }
    }
} 