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

namespace Installer.Engines.ElevenLabs
{
    /// <summary>
    /// SAPI voice implementation for ElevenLabs
    /// </summary>
    [Guid("3d8f5c5f-9d6b-4b92-a12b-1a6dff80b6b4")]
    [ComVisible(true)]
    public class ElevenLabsSapi5VoiceImpl : SapiVoiceImplBase
    {
        private string _apiKey;
        private string _voiceId;
        private string _modelId;
        private HttpClient _httpClient;
        private const string ApiBaseUrl = "https://api.elevenlabs.io/v1";
        
        /// <summary>
        /// Creates a new instance of the ElevenLabsSapi5VoiceImpl class
        /// </summary>
        public ElevenLabsSapi5VoiceImpl()
        {
            try
            {
                // Get the voice token from the registry
                string voiceToken = GetVoiceToken(new Guid("3d8f5c5f-9d6b-4b92-a12b-1a6dff80b6b4"));
                
                // Get voice attributes
                var attributesKey = GetVoiceAttributes(voiceToken);
                
                // Get ElevenLabs parameters
                _apiKey = (string)attributesKey.GetValue("ApiKey");
                _voiceId = (string)attributesKey.GetValue("VoiceId");
                _modelId = (string)attributesKey.GetValue("ModelId") ?? "eleven_monolingual_v1";
                
                if (string.IsNullOrEmpty(_apiKey))
                    throw new Exception("ApiKey not found in registry");
                if (string.IsNullOrEmpty(_voiceId))
                    throw new Exception("VoiceId not found in registry");
                
                // Create HTTP client
                _httpClient = new HttpClient();
                _initialized = true;
                
                LogMessage($"Initialized ElevenLabsSapi5VoiceImpl with voice ID: {_voiceId}");
            }
            catch (Exception ex)
            {
                LogError("Error initializing ElevenLabsSapi5VoiceImpl", ex);
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
                    LogError("ElevenLabsSapi5VoiceImpl not initialized");
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
            // ElevenLabs uses 44.1kHz 16-bit mono MP3
            actualFormatId = targetFormatId;
            actualFormat = new WaveFormatEx
            {
                wFormatTag = 1, // PCM
                nChannels = 1, // Mono
                nSamplesPerSec = 44100, // 44.1kHz
                wBitsPerSample = 16, // 16-bit
                nBlockAlign = 2, // 2 bytes per sample (16-bit mono)
                nAvgBytesPerSec = 88200, // 44100 * 2
                cbSize = 0
            };
        }
        
        /// <summary>
        /// Synthesizes speech using the ElevenLabs API
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
                    model_id = _modelId,
                    voice_settings = new
                    {
                        stability = 0.5,
                        similarity_boost = 0.5
                    }
                };
                
                string synthesisUrl = $"{ApiBaseUrl}/text-to-speech/{_voiceId}";
                
                using (var request = new HttpRequestMessage(HttpMethod.Post, synthesisUrl))
                {
                    request.Headers.Add("xi-api-key", _apiKey);
                    request.Headers.Add("Accept", "audio/mpeg");
                    
                    string json = JsonConvert.SerializeObject(requestBody);
                    request.Content = new StringContent(json, Encoding.UTF8, "application/json");
                    
                    using (var response = await _httpClient.SendAsync(request))
                    {
                        response.EnsureSuccessStatusCode();
                        return await response.Content.ReadAsByteArrayAsync();
                    }
                }
            }
            catch (Exception ex)
            {
                LogError($"Error synthesizing speech: {ex.Message}", ex);
                return new byte[0];
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
        ~ElevenLabsSapi5VoiceImpl()
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
    }
} 