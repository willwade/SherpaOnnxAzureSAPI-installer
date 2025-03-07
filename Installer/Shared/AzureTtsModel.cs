using System.Collections.Generic;

namespace Installer.Shared
{
    public class AzureTtsModel
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string DisplayName { get; set; }
        public string LocalName { get; set; }
        public string ShortName { get; set; }
        public string Gender { get; set; }
        public string Locale { get; set; }
        public string LocaleCode { get; set; }
        public string SampleRateHertz { get; set; }
        public string VoiceType { get; set; }
        public string Status { get; set; }
        public List<string> StyleList { get; set; } = new List<string>();
        public List<string> RoleList { get; set; } = new List<string>();
        
        // Azure subscription information
        public string SubscriptionKey { get; set; }
        public string Region { get; set; }
        
        // Default style and role if applicable
        public string SelectedStyle { get; set; }
        public string SelectedRole { get; set; }
        
        public AzureTtsModel()
        {
        }
        
        public AzureTtsModel(string id, string name, string locale, string gender)
        {
            Id = id;
            Name = name;
            DisplayName = name;
            Locale = locale;
            Gender = gender;
        }
    }
} 