using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using Installer.Core.Base;
using Installer.Core.Models;
using Newtonsoft.Json;

namespace Installer.Engines.ElevenLabs
{
    /// <summary>
    /// ElevenLabs TTS engine implementation
    /// </summary>
    public class ElevenLabsEngine : TtsEngineBase
    {
        private readonly HttpClient _httpClient;
        private const string ApiBaseUrl = "https://api.elevenlabs.io/v1";
        
        /// <summary>
        /// Creates a new instance of the ElevenLabsEngine class
        /// </summary>
        public ElevenLabsEngine()
        {
            _httpClient = new HttpClient();
        }
        
        // Basic properties
        public override string EngineName => "ElevenLabs";
        public override string EngineVersion => "1.0";
        public override string EngineDescription => "ElevenLabs AI Voice Generation";
        public override bool RequiresSsml => false;
        public override bool RequiresAuthentication => true;
        public override bool SupportsOfflineUsage => false;
        
        // Configuration
        public override IEnumerable<ConfigurationParameter> GetRequiredParameters()
        {
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
        }
        
        // SAPI integration
        public override Guid GetEngineClsid() => new Guid("3d8f5c5f-9d6b-4b92-a12b-1a6dff80b6b4");
        public override Type GetSapiImplementationType() => typeof(ElevenLabsSapi5VoiceImpl);
        
        // Voice management
        public override async Task<IEnumerable<TtsVoiceInfo>> GetAvailableVoicesAsync(Dictionary<string, string> config)
        {
            var voices = new List<TtsVoiceInfo>();
            
            try
            {
                // Validate configuration
                if (!ValidateConfiguration(config))
                {
                    LogError("Invalid configuration");
                    return voices;
                }
                
                string apiKey = config["apiKey"];
                
                // Get voices from ElevenLabs API
                var voicesResponse = await GetVoicesAsync(apiKey);
                
                // Convert to voice info
                foreach (var voice in voicesResponse.Voices)
                {
                    var voiceInfo = new TtsVoiceInfo
                    {
                        Id = voice.VoiceId,
                        Name = voice.Name,
                        DisplayName = voice.Name,
                        Gender = DetermineGender(voice),
                        Locale = "en-US", // ElevenLabs primarily supports English
                        EngineName = EngineName,
                        AdditionalAttributes = new Dictionary<string, string>
                        {
                            ["Description"] = voice.Description,
                            ["PreviewUrl"] = voice.PreviewUrl,
                            ["Category"] = voice.Category,
                            ["ModelId"] = config.ContainsKey("modelId") ? config["modelId"] : "eleven_monolingual_v1"
                        }
                    };
                    
                    voices.Add(voiceInfo);
                }
            }
            catch (Exception ex)
            {
                LogError("Error getting available voices", ex);
            }
            
            return voices;
        }
        
        // Speech synthesis
        public override async Task<byte[]> SynthesizeSpeechAsync(string text, string voiceId, Dictionary<string, string> parameters)
        {
            try
            {
                // Get configuration
                string apiKey = parameters["apiKey"];
                string modelId = parameters.ContainsKey("modelId") ? parameters["modelId"] : "eleven_monolingual_v1";
                
                // Create request body
                var requestBody = new
                {
                    text = text,
                    model_id = modelId,
                    voice_settings = new
                    {
                        stability = 0.5,
                        similarity_boost = 0.5
                    }
                };
                
                // Synthesize speech
                return await SynthesizeSpeechAsync(apiKey, voiceId, requestBody);
            }
            catch (Exception ex)
            {
                LogError($"Error synthesizing speech for voice {voiceId}", ex);
                return new byte[0];
            }
        }
        
        // Helper methods
        private async Task<VoicesResponse> GetVoicesAsync(string apiKey)
        {
            string voicesUrl = $"{ApiBaseUrl}/voices";
            
            using (var request = new HttpRequestMessage(HttpMethod.Get, voicesUrl))
            {
                request.Headers.Add("xi-api-key", apiKey);
                
                using (var response = await _httpClient.SendAsync(request))
                {
                    response.EnsureSuccessStatusCode();
                    string json = await response.Content.ReadAsStringAsync();
                    return JsonConvert.DeserializeObject<VoicesResponse>(json);
                }
            }
        }
        
        private async Task<byte[]> SynthesizeSpeechAsync(string apiKey, string voiceId, object requestBody)
        {
            string synthesisUrl = $"{ApiBaseUrl}/text-to-speech/{voiceId}";
            
            using (var request = new HttpRequestMessage(HttpMethod.Post, synthesisUrl))
            {
                request.Headers.Add("xi-api-key", apiKey);
                request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("audio/mpeg"));
                
                string json = JsonConvert.SerializeObject(requestBody);
                request.Content = new StringContent(json, Encoding.UTF8, "application/json");
                
                using (var response = await _httpClient.SendAsync(request))
                {
                    response.EnsureSuccessStatusCode();
                    return await response.Content.ReadAsByteArrayAsync();
                }
            }
        }
        
        private string DetermineGender(Voice voice)
        {
            // ElevenLabs doesn't provide gender information directly
            // This is a simple heuristic based on voice name or category
            string nameLower = voice.Name.ToLower();
            string categoryLower = voice.Category?.ToLower() ?? "";
            
            if (nameLower.Contains("female") || 
                nameLower.Contains("woman") || 
                nameLower.Contains("girl") ||
                categoryLower.Contains("female"))
            {
                return "Female";
            }
            else if (nameLower.Contains("male") || 
                     nameLower.Contains("man") || 
                     nameLower.Contains("boy") ||
                     categoryLower.Contains("male"))
            {
                return "Male";
            }
            
            // Default to neutral if we can't determine
            return "Neutral";
        }
        
        // ElevenLabs API response classes
        private class VoicesResponse
        {
            public List<Voice> Voices { get; set; }
        }
        
        private class Voice
        {
            public string VoiceId { get; set; }
            public string Name { get; set; }
            public string Description { get; set; }
            public string PreviewUrl { get; set; }
            public string Category { get; set; }
            public List<string> AvailableForTiers { get; set; }
            public Dictionary<string, object> Settings { get; set; }
        }
    }
} 