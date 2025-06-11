using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;
using Microsoft.Win32;

namespace Installer
{
    /// <summary>
    /// Manages configuration-based SAPI voices that communicate with AACSpeakHelper pipe service
    /// Each voice is defined by a configuration file containing TTS engine settings
    /// </summary>
    public class ConfigBasedVoiceManager
    {
        private const string VoiceConfigsDirectory = @"C:\Program Files\OpenAssistive\OpenSpeech\voice_configs";
        private const string PipeServiceClsid = "{4A8B9C2D-1E3F-4567-8901-234567890ABC}"; // New CLSID for pipe service voices
        private const string RegistryBasePath = @"SOFTWARE\Microsoft\SPEECH\Voices\Tokens";

        /// <summary>
        /// Configuration structure for a pipe service voice
        /// </summary>
        public class PipeVoiceConfig
        {
            [JsonPropertyName("name")]
            public string Name { get; set; }

            [JsonPropertyName("displayName")]
            public string DisplayName { get; set; }

            [JsonPropertyName("description")]
            public string Description { get; set; }

            [JsonPropertyName("language")]
            public string Language { get; set; }

            [JsonPropertyName("locale")]
            public string Locale { get; set; }

            [JsonPropertyName("gender")]
            public string Gender { get; set; }

            [JsonPropertyName("age")]
            public string Age { get; set; } = "Adult";

            [JsonPropertyName("vendor")]
            public string Vendor { get; set; }

            [JsonPropertyName("ttsConfig")]
            public TtsEngineConfig TtsConfig { get; set; }
        }

        /// <summary>
        /// TTS engine configuration that will be sent to AACSpeakHelper
        /// </summary>
        public class TtsEngineConfig
        {
            [JsonPropertyName("engine")]
            public string Engine { get; set; }

            [JsonPropertyName("voice_id")]
            public string VoiceId { get; set; }

            [JsonPropertyName("azureTTS")]
            public AzureTtsConfig AzureTTS { get; set; }

            [JsonPropertyName("googleTTS")]
            public GoogleTtsConfig GoogleTTS { get; set; }

            [JsonPropertyName("TTS")]
            public GeneralTtsConfig TTS { get; set; }

            [JsonPropertyName("translate")]
            public TranslateConfig Translate { get; set; }
        }

        public class AzureTtsConfig
        {
            [JsonPropertyName("key")]
            public string Key { get; set; }

            [JsonPropertyName("location")]
            public string Location { get; set; }

            [JsonPropertyName("voice")]
            public string Voice { get; set; }

            [JsonPropertyName("style")]
            public string Style { get; set; }

            [JsonPropertyName("role")]
            public string Role { get; set; }
        }

        public class GoogleTtsConfig
        {
            [JsonPropertyName("creds")]
            public string Creds { get; set; }

            [JsonPropertyName("voice")]
            public string Voice { get; set; }

            [JsonPropertyName("lang")]
            public string Lang { get; set; }
        }

        public class GeneralTtsConfig
        {
            [JsonPropertyName("engine")]
            public string Engine { get; set; }

            [JsonPropertyName("voice_id")]
            public string VoiceId { get; set; }

            [JsonPropertyName("bypass_tts")]
            public bool BypassTts { get; set; } = false;
        }

        public class TranslateConfig
        {
            [JsonPropertyName("no_translate")]
            public bool NoTranslate { get; set; } = true;

            [JsonPropertyName("provider")]
            public string Provider { get; set; }

            [JsonPropertyName("start_lang")]
            public string StartLang { get; set; } = "auto";

            [JsonPropertyName("end_lang")]
            public string EndLang { get; set; } = "en";

            [JsonPropertyName("replace_pb")]
            public bool ReplacePb { get; set; } = false;
        }

        /// <summary>
        /// Creates a new configuration-based voice
        /// </summary>
        public void CreateVoice(PipeVoiceConfig config)
        {
            try
            {
                // Ensure voice configs directory exists
                Directory.CreateDirectory(VoiceConfigsDirectory);

                // Save configuration file
                string configPath = Path.Combine(VoiceConfigsDirectory, $"{config.Name}.json");
                string jsonContent = JsonSerializer.Serialize(config, new JsonSerializerOptions 
                { 
                    WriteIndented = true,
                    PropertyNamingPolicy = JsonNamingPolicy.CamelCase
                });
                File.WriteAllText(configPath, jsonContent);

                // Register SAPI voice
                RegisterSapiVoice(config, configPath);

                Console.WriteLine($"Created configuration-based voice: {config.DisplayName}");
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to create voice configuration: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Registers a SAPI voice that points to a configuration file
        /// </summary>
        private void RegisterSapiVoice(PipeVoiceConfig config, string configPath)
        {
            string voiceRegistryPath = $@"{RegistryBasePath}\{config.Name}";
            string lcid = GetLcidFromLocale(config.Locale);

            using (var voiceKey = Registry.LocalMachine.CreateSubKey(voiceRegistryPath))
            {
                // Basic SAPI registration
                voiceKey.SetValue("", config.DisplayName);
                voiceKey.SetValue(lcid, config.DisplayName);
                voiceKey.SetValue("CLSID", PipeServiceClsid);
                voiceKey.SetValue("ConfigPath", configPath); // Store path to config file

                // Set voice attributes
                using (var attributesKey = voiceKey.CreateSubKey("Attributes"))
                {
                    attributesKey.SetValue("Language", lcid);
                    attributesKey.SetValue("Gender", config.Gender ?? "Male");
                    attributesKey.SetValue("Age", config.Age);
                    attributesKey.SetValue("Vendor", config.Vendor ?? "OpenAssistive");
                    attributesKey.SetValue("Version", "1.0");
                    attributesKey.SetValue("Name", config.DisplayName);
                    attributesKey.SetValue("VoiceType", "PipeService");
                    attributesKey.SetValue("Description", config.Description ?? "");
                }
            }

            // Register the CLSID token
            using (var clsidKey = Registry.ClassesRoot.CreateSubKey($@"CLSID\{PipeServiceClsid}"))
            {
                using (var tokenKey = clsidKey.CreateSubKey("Token"))
                {
                    tokenKey.SetValue("", config.DisplayName);
                }
            }
        }

        /// <summary>
        /// Removes a configuration-based voice
        /// </summary>
        public void RemoveVoice(string voiceName)
        {
            try
            {
                // Remove configuration file
                string configPath = Path.Combine(VoiceConfigsDirectory, $"{voiceName}.json");
                if (File.Exists(configPath))
                {
                    File.Delete(configPath);
                }

                // Unregister SAPI voice
                string voiceRegistryPath = $@"{RegistryBasePath}\{voiceName}";
                Registry.LocalMachine.DeleteSubKeyTree(voiceRegistryPath, false);

                Console.WriteLine($"Removed configuration-based voice: {voiceName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error removing voice {voiceName}: {ex.Message}");
            }
        }

        /// <summary>
        /// Lists all configuration-based voices
        /// </summary>
        public List<PipeVoiceConfig> ListVoices()
        {
            var voices = new List<PipeVoiceConfig>();

            if (!Directory.Exists(VoiceConfigsDirectory))
                return voices;

            foreach (string configFile in Directory.GetFiles(VoiceConfigsDirectory, "*.json"))
            {
                try
                {
                    string jsonContent = File.ReadAllText(configFile);
                    var config = JsonSerializer.Deserialize<PipeVoiceConfig>(jsonContent);
                    voices.Add(config);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error loading voice config {configFile}: {ex.Message}");
                }
            }

            return voices;
        }

        /// <summary>
        /// Gets LCID from locale string (e.g., "en-US" -> "409")
        /// </summary>
        private string GetLcidFromLocale(string locale)
        {
            var localeMap = new Dictionary<string, string>
            {
                { "en-US", "409" },
                { "en-GB", "809" },
                { "en-AU", "c09" },
                { "en-CA", "1009" },
                { "fr-FR", "40c" },
                { "de-DE", "407" },
                { "es-ES", "c0a" },
                { "it-IT", "410" },
                { "pt-BR", "416" },
                { "ja-JP", "411" },
                { "ko-KR", "412" },
                { "zh-CN", "804" },
                { "zh-TW", "404" },
                { "ru-RU", "419" },
                { "ar-SA", "401" },
                { "hi-IN", "439" },
                { "th-TH", "41e" },
                { "vi-VN", "42a" },
                { "tr-TR", "41f" },
                { "pl-PL", "415" },
                { "nl-NL", "413" },
                { "sv-SE", "41d" },
                { "da-DK", "406" },
                { "no-NO", "414" },
                { "fi-FI", "40b" }
            };

            return localeMap.TryGetValue(locale ?? "en-US", out string lcid) ? lcid : "409";
        }
    }
}
