using System.Collections.Generic;

namespace Installer.Core.Models
{
    /// <summary>
    /// Configuration parameter for TTS engines
    /// </summary>
    public class ConfigurationParameter
    {
        /// <summary>
        /// Parameter name (used as key in configuration dictionary)
        /// </summary>
        public string Name { get; set; }
        
        /// <summary>
        /// Display name for UI
        /// </summary>
        public string DisplayName { get; set; }
        
        /// <summary>
        /// Description of the parameter
        /// </summary>
        public string Description { get; set; }
        
        /// <summary>
        /// Whether the parameter is required
        /// </summary>
        public bool IsRequired { get; set; } = true;
        
        /// <summary>
        /// Whether the parameter contains sensitive information that should be encrypted
        /// </summary>
        public bool IsSecret { get; set; } = false;
        
        /// <summary>
        /// Default value for the parameter
        /// </summary>
        public string DefaultValue { get; set; }
        
        /// <summary>
        /// List of allowed values (if empty, any value is allowed)
        /// </summary>
        public List<string> AllowedValues { get; set; } = new List<string>();
        
        /// <summary>
        /// Regular expression for validation
        /// </summary>
        public string ValidationRegex { get; set; }
        
        public ConfigurationParameter()
        {
        }
        
        public ConfigurationParameter(string name, string displayName, string description, bool isRequired = true, bool isSecret = false)
        {
            Name = name;
            DisplayName = displayName;
            Description = description;
            IsRequired = isRequired;
            IsSecret = isSecret;
        }
    }
} 