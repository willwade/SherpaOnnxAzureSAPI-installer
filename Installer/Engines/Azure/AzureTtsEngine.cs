using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Installer.Core.Base;
using Installer.Core.Models;
using Newtonsoft.Json;

namespace Installer.Engines.Azure
{
    /// <summary>
    /// Azure TTS engine implementation
    /// </summary>
    public class AzureTtsEngine : TtsEngineBase
    {
        private readonly HttpClient _httpClient;
        
        /// <summary>
        /// Creates a new instance of the AzureTtsEngine class
        /// </summary>
        public AzureTtsEngine()
        {
            _httpClient = new HttpClient();
        }
        
        // Basic properties
        public override string EngineName => "AzureTTS";
        public override string EngineVersion => "1.0";
        public override string EngineDescription => "Azure Cognitive Services Text-to-Speech";
        public override bool RequiresSsml => true;
        public override bool RequiresAuthentication => true;
        public override bool SupportsOfflineUsage => false;
        
        // Configuration
        public override IEnumerable<ConfigurationParameter> GetRequiredParameters()
        {
            return new List<ConfigurationParameter>
            {
                new ConfigurationParameter
                {
                    Name = "subscriptionKey",
                    DisplayName = "Subscription Key",
                    Description = "Azure Cognitive Services subscription key",
                    IsSecret = true
                },
                new ConfigurationParameter
                {
                    Name = "region",
                    DisplayName = "Region",
                    Description = "Azure region (e.g., eastus, westus)"
                }
            };
        }
        
        // SAPI integration
        public override Guid GetEngineClsid() => new Guid("3d8f5c5e-9d6b-4b92-a12b-1a6dff80b6b3");
        public override Type GetSapiImplementationType() => typeof(AzureSapi5VoiceImpl);
        
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
                
                string subscriptionKey = config["subscriptionKey"];
                string region = config["region"];
                
                // Get access token
                string accessToken = await GetAccessTokenAsync(subscriptionKey, region);
                
                // Get voices
                var voicesResponse = await GetVoicesAsync(accessToken, region);
                
                // Convert to voice info
                foreach (var voice in voicesResponse)
                {
                    var voiceInfo = new TtsVoiceInfo
                    {
                        Id = voice.ShortName,
                        Name = voice.DisplayName,
                        DisplayName = voice.DisplayName,
                        Gender = voice.Gender,
                        Locale = voice.Locale,
                        EngineName = EngineName,
                        SupportsStyles = voice.StyleList != null && voice.StyleList.Count > 0,
                        SupportedStyles = voice.StyleList ?? new List<string>(),
                        SupportsRoles = voice.RolePlayList != null && voice.RolePlayList.Count > 0,
                        SupportedRoles = voice.RolePlayList ?? new List<string>(),
                        AdditionalAttributes = new Dictionary<string, string>
                        {
                            ["VoiceName"] = voice.ShortName,
                            ["LocalName"] = voice.LocalName,
                            ["SampleRateHertz"] = voice.SampleRateHertz,
                            ["VoiceType"] = voice.VoiceType,
                            ["Status"] = voice.Status,
                            ["SubscriptionKey"] = subscriptionKey,
                            ["Region"] = region
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
                string subscriptionKey = parameters["subscriptionKey"];
                string region = parameters["region"];
                string selectedStyle = parameters.ContainsKey("selectedStyle") ? parameters["selectedStyle"] : null;
                string selectedRole = parameters.ContainsKey("selectedRole") ? parameters["selectedRole"] : null;
                
                // Get access token
                string accessToken = await GetAccessTokenAsync(subscriptionKey, region);
                
                // Create SSML
                string ssml = CreateSsml(text, voiceId, selectedStyle, selectedRole);
                
                // Synthesize speech
                return await SynthesizeSpeechAsync(accessToken, region, ssml);
            }
            catch (Exception ex)
            {
                LogError($"Error synthesizing speech for voice {voiceId}", ex);
                return new byte[0];
            }
        }
        
        // Helper methods
        private async Task<string> GetAccessTokenAsync(string subscriptionKey, string region)
        {
            string tokenUrl = $"https://{region}.api.cognitive.microsoft.com/sts/v1.0/issueToken";
            
            using (var request = new HttpRequestMessage(HttpMethod.Post, tokenUrl))
            {
                request.Headers.Add("Ocp-Apim-Subscription-Key", subscriptionKey);
                
                using (var response = await _httpClient.SendAsync(request))
                {
                    response.EnsureSuccessStatusCode();
                    return await response.Content.ReadAsStringAsync();
                }
            }
        }
        
        private async Task<List<AzureVoice>> GetVoicesAsync(string accessToken, string region)
        {
            string voicesUrl = $"https://{region}.tts.speech.microsoft.com/cognitiveservices/voices/list";
            
            using (var request = new HttpRequestMessage(HttpMethod.Get, voicesUrl))
            {
                request.Headers.Add("Authorization", $"Bearer {accessToken}");
                
                using (var response = await _httpClient.SendAsync(request))
                {
                    response.EnsureSuccessStatusCode();
                    string json = await response.Content.ReadAsStringAsync();
                    return JsonConvert.DeserializeObject<List<AzureVoice>>(json);
                }
            }
        }
        
        private string CreateSsml(string text, string voiceId, string selectedStyle, string selectedRole)
        {
            var ssmlBuilder = new StringBuilder();
            ssmlBuilder.AppendLine("<speak version=\"1.0\" xmlns=\"http://www.w3.org/2001/10/synthesis\" xmlns:mstts=\"http://www.w3.org/2001/mstts\" xml:lang=\"en-US\">");
            
            if (!string.IsNullOrEmpty(selectedStyle) || !string.IsNullOrEmpty(selectedRole))
            {
                ssmlBuilder.Append("<voice name=\"");
                ssmlBuilder.Append(voiceId);
                ssmlBuilder.AppendLine("\">");
                
                if (!string.IsNullOrEmpty(selectedStyle))
                {
                    ssmlBuilder.Append("<mstts:express-as style=\"");
                    ssmlBuilder.Append(selectedStyle);
                    ssmlBuilder.Append("\"");
                    
                    if (!string.IsNullOrEmpty(selectedRole))
                    {
                        ssmlBuilder.Append(" role=\"");
                        ssmlBuilder.Append(selectedRole);
                        ssmlBuilder.Append("\"");
                    }
                    
                    ssmlBuilder.Append(">");
                    ssmlBuilder.Append(text);
                    ssmlBuilder.AppendLine("</mstts:express-as>");
                }
                else if (!string.IsNullOrEmpty(selectedRole))
                {
                    ssmlBuilder.Append("<mstts:express-as role=\"");
                    ssmlBuilder.Append(selectedRole);
                    ssmlBuilder.Append("\">");
                    ssmlBuilder.Append(text);
                    ssmlBuilder.AppendLine("</mstts:express-as>");
                }
                
                ssmlBuilder.AppendLine("</voice>");
            }
            else
            {
                ssmlBuilder.Append("<voice name=\"");
                ssmlBuilder.Append(voiceId);
                ssmlBuilder.Append("\">");
                ssmlBuilder.Append(text);
                ssmlBuilder.AppendLine("</voice>");
            }
            
            ssmlBuilder.AppendLine("</speak>");
            
            return ssmlBuilder.ToString();
        }
        
        private async Task<byte[]> SynthesizeSpeechAsync(string accessToken, string region, string ssml)
        {
            string synthesisUrl = $"https://{region}.tts.speech.microsoft.com/cognitiveservices/v1";
            
            using (var request = new HttpRequestMessage(HttpMethod.Post, synthesisUrl))
            {
                request.Headers.Add("Authorization", $"Bearer {accessToken}");
                request.Headers.Add("X-Microsoft-OutputFormat", "riff-24khz-16bit-mono-pcm");
                request.Headers.Add("User-Agent", "OpenSpeechTTS");
                
                request.Content = new StringContent(ssml, Encoding.UTF8, "application/ssml+xml");
                
                using (var response = await _httpClient.SendAsync(request))
                {
                    response.EnsureSuccessStatusCode();
                    return await response.Content.ReadAsByteArrayAsync();
                }
            }
        }
        
        // Azure voice class
        private class AzureVoice
        {
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
            public List<string> StyleList { get; set; }
            public List<string> RolePlayList { get; set; }
        }
    }
} 