using System;
using Newtonsoft.Json;

namespace Installer.Shared
{
    public class AzureConfig
    {
        [JsonProperty("defaultKey")]
        public string DefaultKey { get; set; }
        
        [JsonProperty("defaultRegion")]
        public string DefaultRegion { get; set; }
        
        [JsonProperty("secureStorage")]
        public bool SecureStorage { get; set; } = true;
        
        public AzureConfig()
        {
            // Default constructor
        }
        
        public AzureConfig(string key, string region, bool secureStorage = true)
        {
            DefaultKey = key;
            DefaultRegion = region;
            SecureStorage = secureStorage;
        }
    }
} 