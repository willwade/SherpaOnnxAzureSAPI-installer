using System;
using System.Collections.Generic;
using Newtonsoft.Json;

namespace Installer.Core.Models
{
    /// <summary>
    /// Configuration for all TTS engines
    /// </summary>
    public class EngineConfiguration
    {
        /// <summary>
        /// Dictionary of engine configurations, keyed by engine name
        /// </summary>
        [JsonProperty("engines")]
        public Dictionary<string, EngineConfigEntry> Engines { get; set; } = new Dictionary<string, EngineConfigEntry>();
        
        /// <summary>
        /// Default engine to use
        /// </summary>
        [JsonProperty("defaultEngine")]
        public string DefaultEngine { get; set; } = "SherpaOnnx";
        
        /// <summary>
        /// Whether to use secure storage for sensitive information
        /// </summary>
        [JsonProperty("secureStorage")]
        public bool SecureStorage { get; set; } = true;
        
        /// <summary>
        /// Last time the configuration was updated
        /// </summary>
        [JsonProperty("lastUpdated")]
        public DateTime LastUpdated { get; set; } = DateTime.UtcNow;
    }
    
    /// <summary>
    /// Configuration entry for a single engine
    /// </summary>
    public class EngineConfigEntry
    {
        /// <summary>
        /// Whether the engine is enabled
        /// </summary>
        [JsonProperty("enabled")]
        public bool Enabled { get; set; } = true;
        
        /// <summary>
        /// Engine-specific parameters
        /// </summary>
        [JsonProperty("parameters")]
        public Dictionary<string, string> Parameters { get; set; } = new Dictionary<string, string>();
    }
} 