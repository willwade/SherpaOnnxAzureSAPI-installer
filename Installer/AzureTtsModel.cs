using System;
using System.Collections.Generic;
using Newtonsoft.Json;

namespace Installer
{
    public class AzureTtsModel
    {
        [JsonProperty("id")]
        public string Id { get; set; }

        [JsonProperty("name")]
        public string Name { get; set; }

        [JsonProperty("gender")]
        public string Gender { get; set; }

        [JsonProperty("language")]
        public List<LanguageInfo> Language { get; set; }

        [JsonProperty("developer")]
        public string Developer { get; set; } = "Microsoft";

        [JsonProperty("region")]
        public string Region { get; set; }

        [JsonProperty("voice_name")]
        public string VoiceName { get; set; }

        [JsonProperty("style_list")]
        public List<string> StyleList { get; set; }

        [JsonProperty("role_list")]
        public List<string> RoleList { get; set; }

        // This is used to store the subscription key at runtime (not serialized)
        [JsonIgnore]
        public string SubscriptionKey { get; set; }

        public AzureTtsModel()
        {
            Language = new List<LanguageInfo>();
            StyleList = new List<string>();
            RoleList = new List<string>();
        }
    }
}
