using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Installer.Shared;

namespace Installer
{
    public class AzureTtsService
    {
        private readonly string _subscriptionKey;
        private readonly string _region;
        private readonly HttpClient _httpClient;

        public AzureTtsService(string subscriptionKey, string region)
        {
            _subscriptionKey = subscriptionKey ?? throw new ArgumentNullException(nameof(subscriptionKey));
            _region = region ?? throw new ArgumentNullException(nameof(region));
            _httpClient = new HttpClient();
            _httpClient.DefaultRequestHeaders.Add("Ocp-Apim-Subscription-Key", _subscriptionKey);
        }

        public async Task<List<AzureTtsModel>> GetAvailableVoicesAsync()
        {
            try
            {
                string endpoint = $"https://{_region}.tts.speech.microsoft.com/cognitiveservices/voices/list";
                
                var response = await _httpClient.GetAsync(endpoint);
                response.EnsureSuccessStatusCode();
                
                var json = await response.Content.ReadAsStringAsync();
                var voices = JsonConvert.DeserializeObject<List<AzureTtsModel>>(json);
                
                // Process the voices to set IDs and other properties
                foreach (var voice in voices)
                {
                    // Set a unique ID for the voice
                    voice.Id = $"azure-{voice.ShortName}";
                    
                    // Set subscription info
                    voice.SubscriptionKey = _subscriptionKey;
                    voice.Region = _region;
                }
                
                return voices;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error fetching Azure voices: {ex.Message}");
                throw;
            }
        }

        public async Task<bool> ValidateSubscriptionAsync()
        {
            try
            {
                string endpoint = $"https://{_region}.tts.speech.microsoft.com/cognitiveservices/voices/list";
                
                var response = await _httpClient.GetAsync(endpoint);
                return response.IsSuccessStatusCode;
            }
            catch
            {
                return false;
            }
        }

        public async Task<byte[]> SynthesizeSpeechAsync(AzureTtsModel voice, string text, string style = null, string role = null)
        {
            try
            {
                string endpoint = $"https://{_region}.tts.speech.microsoft.com/cognitiveservices/v1";
                
                var request = new HttpRequestMessage(HttpMethod.Post, endpoint);
                request.Headers.Add("X-Microsoft-OutputFormat", "riff-24khz-16bit-mono-pcm");
                
                // Build SSML
                var ssml = new StringBuilder();
                ssml.Append("<speak version='1.0' xmlns='http://www.w3.org/2001/10/synthesis' xmlns:mstts='https://www.w3.org/2001/mstts' xml:lang='");
                ssml.Append(voice.Locale);
                ssml.Append("'><voice name='");
                ssml.Append(voice.ShortName);
                ssml.Append("'>");
                
                // Add style if specified and available
                if (!string.IsNullOrEmpty(style) && voice.StyleList.Contains(style))
                {
                    ssml.Append("<mstts:express-as style='");
                    ssml.Append(style);
                    ssml.Append("'>");
                }
                
                // Add role if specified and available
                if (!string.IsNullOrEmpty(role) && voice.RoleList.Contains(role))
                {
                    ssml.Append("<mstts:express-as role='");
                    ssml.Append(role);
                    ssml.Append("'>");
                }
                
                ssml.Append(text);
                
                // Close tags in reverse order
                if (!string.IsNullOrEmpty(role) && voice.RoleList.Contains(role))
                {
                    ssml.Append("</mstts:express-as>");
                }
                
                if (!string.IsNullOrEmpty(style) && voice.StyleList.Contains(style))
                {
                    ssml.Append("</mstts:express-as>");
                }
                
                ssml.Append("</voice></speak>");
                
                request.Content = new StringContent(ssml.ToString(), Encoding.UTF8, "application/ssml+xml");
                
                var response = await _httpClient.SendAsync(request);
                response.EnsureSuccessStatusCode();
                
                return await response.Content.ReadAsByteArrayAsync();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error synthesizing speech: {ex.Message}");
                throw;
            }
        }
    }
} 