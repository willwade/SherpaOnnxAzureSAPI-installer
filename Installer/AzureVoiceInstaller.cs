using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Installer.Shared;

namespace Installer
{
    public class AzureVoiceInstaller
    {
        private readonly string configDirectory;
        private readonly HttpClient _httpClient;

        public AzureVoiceInstaller()
        {
            // Use Program Files for all users
            configDirectory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                "OpenSpeech",
                "azure"
            );
            
            // Ensure directory exists with proper permissions
            Directory.CreateDirectory(configDirectory);
            
            _httpClient = new HttpClient();
        }

        /// <summary>
        /// Fetches available Azure voices for a given subscription key and region
        /// </summary>
        public async Task<List<AzureTtsModel>> GetAvailableVoicesAsync(string subscriptionKey, string region)
        {
            try
            {
                // Create request to Azure Cognitive Services
                var request = new HttpRequestMessage(HttpMethod.Get, 
                    $"https://{region}.tts.speech.microsoft.com/cognitiveservices/voices/list");
                
                request.Headers.Add("Ocp-Apim-Subscription-Key", subscriptionKey);
                
                var response = await _httpClient.SendAsync(request);
                response.EnsureSuccessStatusCode();
                
                var voicesJson = await response.Content.ReadAsStringAsync();
                var azureVoices = JsonConvert.DeserializeObject<List<AzureVoiceInfo>>(voicesJson);
                
                // Convert Azure voice info to our model format
                var result = new List<AzureTtsModel>();
                foreach (var voice in azureVoices)
                {
                    var model = new AzureTtsModel
                    {
                        Id = voice.ShortName,
                        Name = $"Azure {voice.DisplayName}",
                        Gender = voice.Gender,
                        Region = region,
                        VoiceName = voice.ShortName,
                        StyleList = voice.StyleList ?? new List<string>(),
                        RoleList = voice.RolePlayList ?? new List<string>(),
                        SubscriptionKey = subscriptionKey
                    };
                    
                    // Parse language code
                    var langParts = voice.Locale.Split('-');
                    if (langParts.Length >= 2)
                    {
                        model.Language.Add(new LanguageInfo
                        {
                            LangCode = langParts[0],
                            Country = langParts[1],
                            LanguageName = voice.LocalName
                        });
                    }
                    else
                    {
                        model.Language.Add(new LanguageInfo
                        {
                            LangCode = voice.Locale,
                            LanguageName = voice.LocalName
                        });
                    }
                    
                    result.Add(model);
                }
                
                return result;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error fetching Azure voices: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Saves Azure voice configuration to a local file
        /// </summary>
        public void SaveVoiceConfig(AzureTtsModel model)
        {
            try
            {
                string configPath = Path.Combine(configDirectory, $"{model.Id}.json");
                
                // Create a copy without the subscription key for local storage
                var configModel = new AzureTtsModel
                {
                    Id = model.Id,
                    Name = model.Name,
                    Gender = model.Gender,
                    Region = model.Region,
                    VoiceName = model.VoiceName,
                    StyleList = model.StyleList,
                    RoleList = model.RoleList,
                    Language = model.Language,
                    Developer = model.Developer
                };
                
                string json = JsonConvert.SerializeObject(configModel, Formatting.Indented);
                File.WriteAllText(configPath, json);
                
                Console.WriteLine($"Saved Azure voice configuration to {configPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving Azure voice configuration: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Tests the Azure TTS voice by making a simple request
        /// </summary>
        public async Task<bool> TestVoiceAsync(AzureTtsModel model)
        {
            try
            {
                // Create SSML for testing
                string ssml = $@"
                <speak version='1.0' xmlns='http://www.w3.org/2001/10/synthesis' xml:lang='en-US'>
                    <voice name='{model.VoiceName}'>
                        This is a test of the Azure Text to Speech voice.
                    </voice>
                </speak>";

                // Set up the request
                var request = new HttpRequestMessage(HttpMethod.Post, 
                    $"https://{model.Region}.tts.speech.microsoft.com/cognitiveservices/v1");
                
                request.Headers.Add("Authorization", $"Bearer {await GetTokenAsync(model.SubscriptionKey, model.Region)}");
                request.Headers.Add("X-Microsoft-OutputFormat", "riff-24khz-16bit-mono-pcm");
                request.Headers.Add("User-Agent", "OpenSpeechTTS");
                
                request.Content = new StringContent(ssml, System.Text.Encoding.UTF8, "application/ssml+xml");

                // Send the request
                var response = await _httpClient.SendAsync(request);
                response.EnsureSuccessStatusCode();
                
                // If we get here, the test was successful
                Console.WriteLine($"Successfully tested Azure voice: {model.Name}");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error testing Azure voice: {ex.Message}");
                return false;
            }
        }

        private async Task<string> GetTokenAsync(string subscriptionKey, string region)
        {
            var request = new HttpRequestMessage(HttpMethod.Post, 
                $"https://{region}.api.cognitive.microsoft.com/sts/v1.0/issueToken");
            
            request.Headers.Add("Ocp-Apim-Subscription-Key", subscriptionKey);
            
            var response = await _httpClient.SendAsync(request);
            response.EnsureSuccessStatusCode();
            
            return await response.Content.ReadAsStringAsync();
        }
    }

    // Class to parse the Azure Cognitive Services voice list response
    public class AzureVoiceInfo
    {
        [JsonProperty("Name")]
        public string Name { get; set; }
        
        [JsonProperty("DisplayName")]
        public string DisplayName { get; set; }
        
        [JsonProperty("LocalName")]
        public string LocalName { get; set; }
        
        [JsonProperty("ShortName")]
        public string ShortName { get; set; }
        
        [JsonProperty("Gender")]
        public string Gender { get; set; }
        
        [JsonProperty("Locale")]
        public string Locale { get; set; }
        
        [JsonProperty("LocaleName")]
        public string LocaleName { get; set; }
        
        [JsonProperty("SampleRateHertz")]
        public string SampleRateHertz { get; set; }
        
        [JsonProperty("VoiceType")]
        public string VoiceType { get; set; }
        
        [JsonProperty("Status")]
        public string Status { get; set; }
        
        [JsonProperty("StyleList")]
        public List<string> StyleList { get; set; }
        
        [JsonProperty("RolePlayList")]
        public List<string> RolePlayList { get; set; }
    }
}
