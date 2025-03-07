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

namespace Installer.Engines.PlayHT
{
    /// <summary>
    /// PlayHT TTS engine implementation
    /// </summary>
    public class PlayHTEngine : TtsEngineBase
    {
        private readonly HttpClient _httpClient;
        private const string ApiBaseUrl = "https://api.play.ht/api/v2";
        
        /// <summary>
        /// Creates a new instance of the PlayHTEngine class
        /// </summary>
        public PlayHTEngine()
        {
            _httpClient = new HttpClient();
        }
        
        // Basic properties
        public override string EngineName => "PlayHT";
        public override string EngineVersion => "1.0";
        public override string EngineDescription => "PlayHT AI Voice Generation";
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
                    Description = "PlayHT API key",
                    IsSecret = true
                },
                new ConfigurationParameter
                {
                    Name = "userId",
                    DisplayName = "User ID",
                    Description = "PlayHT user ID"
                },
                new ConfigurationParameter
                {
                    Name = "quality",
                    DisplayName = "Voice Quality",
                    Description = "Voice quality (draft or premium)",
                    IsRequired = false,
                    DefaultValue = "premium",
                    AllowedValues = new List<string> { "draft", "premium" }
                }
            };
        }
        
        // SAPI integration
        public override Guid GetEngineClsid() => new Guid("3d8f5c60-9d6b-4b92-a12b-1a6dff80b6b5");
        public override Type GetSapiImplementationType() => typeof(PlayHTSapi5VoiceImpl);
        
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
                string userId = config["userId"];
                
                // Get voices from PlayHT API
                var voicesResponse = await GetVoicesAsync(apiKey, userId);
                
                // Convert to voice info
                foreach (var voice in voicesResponse)
                {
                    var voiceInfo = new TtsVoiceInfo
                    {
                        Id = voice.Id,
                        Name = voice.Name,
                        DisplayName = voice.Name,
                        Gender = voice.Gender,
                        Locale = voice.Language,
                        EngineName = EngineName,
                        AdditionalAttributes = new Dictionary<string, string>
                        {
                            ["VoiceEngine"] = voice.VoiceEngine,
                            ["SampleUrl"] = voice.Sample,
                            ["CloneVoiceId"] = voice.CloneVoiceId,
                            ["Quality"] = config.ContainsKey("quality") ? config["quality"] : "premium"
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
                string userId = parameters["userId"];
                string quality = parameters.ContainsKey("quality") ? parameters["quality"] : "premium";
                
                // Create request body
                var requestBody = new
                {
                    text = text,
                    voice = voiceId,
                    quality = quality,
                    output_format = "mp3",
                    speed = 1.0,
                    sample_rate = 24000
                };
                
                // Generate speech
                var conversionResponse = await GenerateSpeechAsync(apiKey, userId, requestBody);
                
                // Get audio URL
                string audioUrl = conversionResponse.AudioUrl;
                
                // Download audio
                return await DownloadAudioAsync(audioUrl);
            }
            catch (Exception ex)
            {
                LogError($"Error synthesizing speech for voice {voiceId}", ex);
                return new byte[0];
            }
        }
        
        // Helper methods
        private async Task<List<PlayHTVoice>> GetVoicesAsync(string apiKey, string userId)
        {
            string voicesUrl = $"{ApiBaseUrl}/voices";
            
            using (var request = new HttpRequestMessage(HttpMethod.Get, voicesUrl))
            {
                request.Headers.Add("X-API-KEY", apiKey);
                request.Headers.Add("AUTHORIZATION", userId);
                
                using (var response = await _httpClient.SendAsync(request))
                {
                    response.EnsureSuccessStatusCode();
                    string json = await response.Content.ReadAsStringAsync();
                    return JsonConvert.DeserializeObject<List<PlayHTVoice>>(json);
                }
            }
        }
        
        private async Task<ConversionResponse> GenerateSpeechAsync(string apiKey, string userId, object requestBody)
        {
            string conversionUrl = $"{ApiBaseUrl}/tts";
            
            using (var request = new HttpRequestMessage(HttpMethod.Post, conversionUrl))
            {
                request.Headers.Add("X-API-KEY", apiKey);
                request.Headers.Add("AUTHORIZATION", userId);
                request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                
                string json = JsonConvert.SerializeObject(requestBody);
                request.Content = new StringContent(json, Encoding.UTF8, "application/json");
                
                using (var response = await _httpClient.SendAsync(request))
                {
                    response.EnsureSuccessStatusCode();
                    string responseJson = await response.Content.ReadAsStringAsync();
                    return JsonConvert.DeserializeObject<ConversionResponse>(responseJson);
                }
            }
        }
        
        private async Task<byte[]> DownloadAudioAsync(string audioUrl)
        {
            using (var response = await _httpClient.GetAsync(audioUrl))
            {
                response.EnsureSuccessStatusCode();
                return await response.Content.ReadAsByteArrayAsync();
            }
        }
        
        // PlayHT API response classes
        private class PlayHTVoice
        {
            public string Id { get; set; }
            public string Name { get; set; }
            public string Gender { get; set; }
            public string Language { get; set; }
            public string VoiceEngine { get; set; }
            public string Sample { get; set; }
            public string CloneVoiceId { get; set; }
        }
        
        private class ConversionResponse
        {
            [JsonProperty("audioUrl")]
            public string AudioUrl { get; set; }
            
            [JsonProperty("transcriptionId")]
            public string TranscriptionId { get; set; }
        }
    }
} 