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
    public class AzureTTS : IDisposable
    {
        private readonly string _subscriptionKey;
        private readonly string _region;
        private readonly string _voiceName;
        private readonly string _selectedStyle;
        private readonly string _selectedRole;
        private readonly HttpClient _httpClient;
        private bool _disposed;
        private string _accessToken;
        private DateTime _tokenExpiration;

        /// <summary>
        /// Creates a new instance of the AzureTTS class
        /// </summary>
        /// <param name="subscriptionKey">Azure Cognitive Services subscription key</param>
        /// <param name="region">Azure region (e.g., "eastus")</param>
        /// <param name="voiceName">Azure voice name (e.g., "en-US-AriaNeural")</param>
        /// <param name="selectedStyle">Optional style for the voice</param>
        /// <param name="selectedRole">Optional role for the voice</param>
        public AzureTTS(string subscriptionKey, string region, string voiceName, string selectedStyle = null, string selectedRole = null)
        {
            if (string.IsNullOrEmpty(subscriptionKey))
                throw new ArgumentNullException(nameof(subscriptionKey));
            
            if (string.IsNullOrEmpty(region))
                throw new ArgumentNullException(nameof(region));
            
            if (string.IsNullOrEmpty(voiceName))
                throw new ArgumentNullException(nameof(voiceName));

            _subscriptionKey = subscriptionKey;
            _region = region;
            _voiceName = voiceName;
            _selectedStyle = selectedStyle;
            _selectedRole = selectedRole;
            
            _httpClient = new HttpClient();
            _httpClient.DefaultRequestHeaders.Add("Ocp-Apim-Subscription-Key", _subscriptionKey);
        }

        /// <summary>
        /// Speaks the specified text and returns the audio as a WAV stream
        /// </summary>
        /// <param name="text">The text to speak</param>
        /// <param name="outputStream">The stream to write the WAV data to</param>
        public void SpeakToWaveStream(string text, Stream outputStream)
        {
            if (_disposed)
                throw new ObjectDisposedException(nameof(AzureTTS));

            try
            {
                // Get the audio data asynchronously but wait for it to complete
                byte[] audioData = SynthesizeSpeechAsync(text).GetAwaiter().GetResult();
                
                // Write the audio data directly to the stream
                // Azure TTS already returns WAV format data
                outputStream.Write(audioData, 0, audioData.Length);
            }
            catch (Exception ex)
            {
                // Log the error for debugging
                try
                {
                    File.AppendAllText("C:\\OpenSpeech\\azure_error.log", 
                        $"{DateTime.Now}: Error in SpeakToWaveStream: {ex.Message}\nText: {text}\n{ex.StackTrace}\n\n");
                }
                catch { }
                
                throw;
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

        public async Task<byte[]> SynthesizeSpeechAsync(string text, string locale = "en-US")
        {
            if (_disposed)
                throw new ObjectDisposedException(nameof(AzureTTS));

            try
            {
                string endpoint = $"https://{_region}.tts.speech.microsoft.com/cognitiveservices/v1";
                
                var request = new HttpRequestMessage(HttpMethod.Post, endpoint);
                request.Headers.Add("X-Microsoft-OutputFormat", "riff-24khz-16bit-mono-pcm");
                
                // Build SSML
                var ssml = new StringBuilder();
                ssml.Append("<speak version='1.0' xmlns='http://www.w3.org/2001/10/synthesis' xmlns:mstts='https://www.w3.org/2001/mstts' xml:lang='");
                ssml.Append(locale);
                ssml.Append("'><voice name='");
                ssml.Append(_voiceName);
                ssml.Append("'>");
                
                // Add style if specified
                if (!string.IsNullOrEmpty(_selectedStyle))
                {
                    ssml.Append("<mstts:express-as style='");
                    ssml.Append(_selectedStyle);
                    ssml.Append("'>");
                }
                
                // Add role if specified
                if (!string.IsNullOrEmpty(_selectedRole))
                {
                    ssml.Append("<mstts:express-as role='");
                    ssml.Append(_selectedRole);
                    ssml.Append("'>");
                }
                
                ssml.Append(text);
                
                // Close tags in reverse order
                if (!string.IsNullOrEmpty(_selectedRole))
                {
                    ssml.Append("</mstts:express-as>");
                }
                
                if (!string.IsNullOrEmpty(_selectedStyle))
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
                // Log the error for debugging
                try
                {
                    File.AppendAllText("C:\\OpenSpeech\\azure_error.log", 
                        $"{DateTime.Now}: Error in SynthesizeSpeechAsync: {ex.Message}\nText: {text}\n{ex.StackTrace}\n\n");
                }
                catch { }
                
                throw;
            }
        }

        public byte[] GenerateAudio(string text, string locale = "en-US")
        {
            if (_disposed)
                throw new ObjectDisposedException(nameof(AzureTTS));

            try
            {
                // Get the audio data asynchronously but wait for it to complete
                return SynthesizeSpeechAsync(text, locale).GetAwaiter().GetResult();
            }
            catch (Exception ex)
            {
                // Log the error for debugging
                try
                {
                    File.AppendAllText("C:\\OpenSpeech\\azure_error.log", 
                        $"{DateTime.Now}: Error in GenerateAudio: {ex.Message}\nText: {text}\n{ex.StackTrace}\n\n");
                }
                catch { }
                
                throw;
            }
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                _httpClient?.Dispose();
                _disposed = true;
            }
        }
    }
}
