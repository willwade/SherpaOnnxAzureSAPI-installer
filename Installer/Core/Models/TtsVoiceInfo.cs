using System.Collections.Generic;

namespace Installer.Core.Models
{
    /// <summary>
    /// Standard voice information structure for all TTS engines
    /// </summary>
    public class TtsVoiceInfo
    {
        // Basic properties
        public string Id { get; set; }
        public string Name { get; set; }
        public string DisplayName { get; set; }
        public string Gender { get; set; }
        public string Locale { get; set; }
        public string Age { get; set; } = "Adult";
        
        // Engine-specific properties
        public string EngineName { get; set; }
        public Dictionary<string, string> AdditionalAttributes { get; set; } = new Dictionary<string, string>();
        
        // Feature support
        public bool SupportsStyles { get; set; }
        public List<string> SupportedStyles { get; set; } = new List<string>();
        public bool SupportsRoles { get; set; }
        public List<string> SupportedRoles { get; set; } = new List<string>();
        
        // Selected options
        public string SelectedStyle { get; set; }
        public string SelectedRole { get; set; }
        
        public TtsVoiceInfo()
        {
        }
        
        public TtsVoiceInfo(string id, string name, string locale, string gender, string engineName)
        {
            Id = id;
            Name = name;
            DisplayName = name;
            Locale = locale;
            Gender = gender;
            EngineName = engineName;
        }
    }
} 