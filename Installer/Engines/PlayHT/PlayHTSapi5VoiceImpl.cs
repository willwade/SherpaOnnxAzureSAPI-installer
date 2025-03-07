using System;
using System.IO;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Installer.Core.Base;
using Installer.Core.Interfaces;
using Microsoft.Win32;
using NAudio.Wave;
using Newtonsoft.Json;

namespace Installer.Engines.PlayHT
{
    /// <summary>
    /// SAPI voice implementation for PlayHT
    /// </summary>
    [Guid("3d8f5c60-9d6b-4b92-a12b-1a6dff80b6b5")]
    [ComVisible(true)]
    public class PlayHTSapi5VoiceImpl : SapiVoiceImplBase
    {
        private string _apiKey;
        private string _userId;
        private string _voiceId;
        private string _quality;
        private HttpClient _httpClient;
        private const string ApiBaseUrl = "https://api.play.ht/api/v2";
        
        /// <summary>
        /// Creates a new instance of the PlayHTSapi5VoiceImpl class
        /// </summary>
        public PlayHTSapi5VoiceImpl()
        {
            try
            {
                // Get the voice token from the registry
                string voiceToken = GetVoiceToken(new Guid("3d8f5c60-9d6b-4b92-a12b-1a6dff80b6b5"));
                
                // Get voice attributes
                var attributesKey = GetVoiceAttributes(voiceToken);
                
                // Get PlayHT parameters
                _apiKey = (string)attributesKey.GetValue("ApiKey");
                _userId = (string)attributesKey.GetValue("UserId");
                _voiceId = (string)attributesKey.GetValue("VoiceId");
                _quality = (string)attributesKey.GetValue("Quality") ?? "premium";
                
                if (string.IsNullOrEmpty(_apiKey))
                    throw new Exception("ApiKey not found in registry");
                if (string.IsNullOrEmpty(_userId))
                    throw new Exception("UserId not found in registry");
                if (string.IsNullOrEmpty(_voiceId))
                    throw new Exception("VoiceId not found in registry");
                
                // Create HTTP client
                _httpClient = new HttpClient();
                _initialized = true;
                
                LogMessage($"Initialized PlayHTSapi5VoiceImpl with voice ID: {_voiceId}");
            }
            catch (Exception ex)
            {
                LogError("Error initializing PlayHTSapi5VoiceImpl", ex);
                throw;
            }
        }
        
        /// <summary>
        /// Speaks the provided text
        /// </summary>
        /// <param name="text">Text to speak</param>
        /// <param name="flags">SAPI flags</param>
        /// <param name="reserved">Reserved parameter</param>
        public override void Speak(string text, uint flags, IntPtr reserved)
        {
            try
            {
                if (!_initialized || _httpClient == null)
                {
                    LogError("PlayHTSapi5VoiceImpl not initialized");
                    return;
                }
                
                LogMessage($"Speaking text: {text}");
                
                // Generate audio
                byte[] audioData = SynthesizeSpeechAsync(text).GetAwaiter().GetResult();
                
                if (audioData == null || audioData.Length == 0)
                {
                    LogError("Failed to generate audio");
                    return;
                }
                
                // Play audio
                PlayAudio(audioData);
            }
            catch (Exception ex)
            {
                LogError("Error speaking text", ex);
            }
        }
        
        /// <summary>
        /// Gets the output format for the voice
        /// </summary>
        /// <param name="targetFormatId">Target format ID</param>
        /// <param name="targetFormat">Target format</param>
        /// <param name="actualFormatId">Actual format ID</param>
        /// <param name="actualFormat">Actual format</param>
        public override void GetOutputFormat(ref Guid targetFormatId, ref WaveFormatEx targetFormat, out Guid actualFormatId, out WaveFormatEx actualFormat)
        {
            // PlayHT uses 24kHz 16-bit mono MP3
            actualFormatId = targetFormatId;
            actualFormat = new WaveFormatEx
            {
                wFormatTag = 1, // PCM
                nChannels = 1, // Mono
                nSamplesPerSec = 24000, // 24kHz
                wBitsPerSample = 16, // 16-bit
                nBlockAlign = 2, // 2 bytes per sample (16-bit mono)
                nAvgBytesPerSec = 48000, // 24000 * 2
                cbSize = 0
            };
        }
        
        /// <summary>
        /// Synthesizes speech using the PlayHT API
        /// </summary>
        /// <param name="text">The text to synthesize</param>
        /// <returns>The audio data</returns>
        private async Task<byte[]> SynthesizeSpeechAsync(string text)
        {
            try
            {
                // Create request body
                var requestBody = new
                {
                    text = text,
                    voice = _voiceId,
                    quality = _quality,
                    output_format = "mp3",
                    speed = 1.0,
                    sample_rate = 24000
                };
                
                // Generate speech
                var conversionResponse = await GenerateSpeechAsync(requestBody);
                
                // Get audio URL
                string audioUrl = conversionResponse.AudioUrl;
                
                // Download audio
                return await DownloadAudioAsync(audioUrl);
            }
            catch (Exception ex)
            {
                LogError($"Error synthesizing speech: {ex.Message}", ex);
                return new byte[0];
            }
        }
        
        /// <summary>
        /// Generates speech using the PlayHT API
        /// </summary>
        /// <param name="requestBody">The request body</param>
        /// <returns>The conversion response</returns>
        private async Task<ConversionResponse> GenerateSpeechAsync(object requestBody)
        {
            string conversionUrl = $"{ApiBaseUrl}/tts";
            
            using (var request = new HttpRequestMessage(HttpMethod.Post, conversionUrl))
            {
                request.Headers.Add("X-API-KEY", _apiKey);
                request.Headers.Add("AUTHORIZATION", _userId);
                request.Headers.Add("Accept", "application/json");
                
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
        
        /// <summary>
        /// Downloads audio from a URL
        /// </summary>
        /// <param name="audioUrl">The audio URL</param>
        /// <returns>The audio data</returns>
        private async Task<byte[]> DownloadAudioAsync(string audioUrl)
        {
            using (var response = await _httpClient.GetAsync(audioUrl))
            {
                response.EnsureSuccessStatusCode();
                return await response.Content.ReadAsByteArrayAsync();
            }
        }
        
        /// <summary>
        /// Plays audio data
        /// </summary>
        /// <param name="audioData">The audio data to play</param>
        private void PlayAudio(byte[] audioData)
        {
            try
            {
                using (var stream = new MemoryStream(audioData))
                using (var reader = new Mp3FileReader(stream))
                using (var waveOut = new WaveOutEvent())
                {
                    waveOut.Init(reader);
                    waveOut.Play();
                    
                    // Wait for playback to complete
                    while (waveOut.PlaybackState == PlaybackState.Playing)
                    {
                        System.Threading.Thread.Sleep(100);
                    }
                }
            }
            catch (Exception ex)
            {
                LogError($"Error playing audio: {ex.Message}", ex);
            }
        }
        
        /// <summary>
        /// Finalizes the instance
        /// </summary>
        ~PlayHTSapi5VoiceImpl()
        {
            try
            {
                _httpClient?.Dispose();
                _httpClient = null;
            }
            catch
            {
                // Ignore errors during finalization
            }
        }
        
        /// <summary>
        /// PlayHT API conversion response
        /// </summary>
        private class ConversionResponse
        {
            [JsonProperty("audioUrl")]
            public string AudioUrl { get; set; }
            
            [JsonProperty("transcriptionId")]
            public string TranscriptionId { get; set; }
        }
    }
} 