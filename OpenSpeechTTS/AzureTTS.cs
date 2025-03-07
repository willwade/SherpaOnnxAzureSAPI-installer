using System;
using System.IO;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using NAudio.Wave;
using Newtonsoft.Json;

namespace OpenSpeechTTS
{
    /// <summary>
    /// Implements text-to-speech functionality using Azure Cognitive Services
    /// </summary>
    public class AzureTTS
    {
        private readonly string _subscriptionKey;
        private readonly string _region;
        private readonly string _voiceName;
        private readonly HttpClient _httpClient;
        private string _accessToken;
        private DateTime _tokenExpiration;

        /// <summary>
        /// Creates a new instance of the AzureTTS class
        /// </summary>
        /// <param name="subscriptionKey">Azure Cognitive Services subscription key</param>
        /// <param name="region">Azure region (e.g., "eastus")</param>
        /// <param name="voiceName">Azure voice name (e.g., "en-US-AriaNeural")</param>
        public AzureTTS(string subscriptionKey, string region, string voiceName)
        {
            _subscriptionKey = subscriptionKey;
            _region = region;
            _voiceName = voiceName;
            _httpClient = new HttpClient();
            _tokenExpiration = DateTime.MinValue;
        }

        /// <summary>
        /// Speaks the specified text and returns the audio as a WAV stream
        /// </summary>
        /// <param name="text">The text to speak</param>
        /// <param name="outputStream">The stream to write the WAV data to</param>
        public void SpeakToWaveStream(string text, Stream outputStream)
        {
            // Ensure we have a valid access token
            EnsureTokenAsync().Wait();

            // Create SSML document
            string ssml = GenerateSsml(text);

            // Set up the request
            var request = new HttpRequestMessage(HttpMethod.Post, 
                $"https://{_region}.tts.speech.microsoft.com/cognitiveservices/v1");
            
            request.Headers.Add("Authorization", $"Bearer {_accessToken}");
            request.Headers.Add("X-Microsoft-OutputFormat", "riff-24khz-16bit-mono-pcm");
            request.Headers.Add("User-Agent", "OpenSpeechTTS");
            
            request.Content = new StringContent(ssml, Encoding.UTF8, "application/ssml+xml");

            // Send the request and get the audio
            var response = _httpClient.SendAsync(request).Result;
            response.EnsureSuccessStatusCode();

            // Copy the audio data to the output stream
            using (var audioStream = response.Content.ReadAsStreamAsync().Result)
            {
                audioStream.CopyTo(outputStream);
                outputStream.Position = 0;
            }
        }

        /// <summary>
        /// Speaks the specified text and returns the audio as a byte array
        /// </summary>
        /// <param name="text">The text to speak</param>
        /// <returns>The audio data as a WAV byte array</returns>
        public byte[] SpeakToWaveBytes(string text)
        {
            using (var ms = new MemoryStream())
            {
                SpeakToWaveStream(text, ms);
                return ms.ToArray();
            }
        }

        /// <summary>
        /// Speaks the specified text and saves the audio to a WAV file
        /// </summary>
        /// <param name="text">The text to speak</param>
        /// <param name="outputPath">The path to save the WAV file</param>
        public void SpeakToWaveFile(string text, string outputPath)
        {
            using (var fs = new FileStream(outputPath, FileMode.Create))
            {
                SpeakToWaveStream(text, fs);
            }
        }

        /// <summary>
        /// Speaks the specified text to the default audio device
        /// </summary>
        /// <param name="text">The text to speak</param>
        public void SpeakToDefaultDevice(string text)
        {
            byte[] audioData = SpeakToWaveBytes(text);
            
            using (var ms = new MemoryStream(audioData))
            using (var reader = new WaveFileReader(ms))
            using (var waveOut = new NAudio.Wave.WaveOutEvent())
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

        private string GenerateSsml(string text)
        {
            // Create a simple SSML document
            var ssml = new XDocument(
                new XElement("speak",
                    new XAttribute("version", "1.0"),
                    new XAttribute(XNamespace.Xmlns + "mstts", "https://www.w3.org/2001/mstts"),
                    new XAttribute(XNamespace.Xml + "lang", "en-US"),
                    new XElement("voice",
                        new XAttribute("name", _voiceName),
                        text
                    )
                )
            );

            return ssml.ToString();
        }

        private async Task EnsureTokenAsync()
        {
            // If the token is still valid, don't request a new one
            if (DateTime.UtcNow < _tokenExpiration)
            {
                return;
            }

            // Request a new token
            var request = new HttpRequestMessage(HttpMethod.Post, 
                $"https://{_region}.api.cognitive.microsoft.com/sts/v1.0/issueToken");
            
            request.Headers.Add("Ocp-Apim-Subscription-Key", _subscriptionKey);
            
            var response = await _httpClient.SendAsync(request);
            response.EnsureSuccessStatusCode();
            
            _accessToken = await response.Content.ReadAsStringAsync();
            
            // Tokens are valid for 10 minutes, but we'll refresh after 9 minutes to be safe
            _tokenExpiration = DateTime.UtcNow.AddMinutes(9);
        }
    }
}
