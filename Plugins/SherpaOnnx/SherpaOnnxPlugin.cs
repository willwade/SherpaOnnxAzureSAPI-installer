using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using OpenSpeech.TTS.Plugins;

namespace OpenSpeech.TTS.Plugins.SherpaOnnx
{
    /// <summary>
    /// Sherpa ONNX TTS engine plugin
    /// </summary>
    public class SherpaOnnxPlugin : TtsEngineBase
    {
        /// <summary>
        /// Gets the name of the TTS engine
        /// </summary>
        public override string EngineName => "SherpaOnnx";
        
        /// <summary>
        /// Gets the version of the TTS engine
        /// </summary>
        public override string EngineVersion => "1.0.0";
        
        /// <summary>
        /// Gets a description of the TTS engine
        /// </summary>
        public override string EngineDescription => "Sherpa ONNX Text-to-Speech Engine";
        
        /// <summary>
        /// Validates the engine configuration
        /// </summary>
        /// <param name="configuration">The engine configuration</param>
        /// <returns>True if the configuration is valid, otherwise false</returns>
        public override bool ValidateConfiguration(Dictionary<string, string> configuration)
        {
            string modelsDirectory = GetConfigValue<string>(configuration, "modelsDirectory", string.Empty);
            
            if (string.IsNullOrEmpty(modelsDirectory) || !Directory.Exists(modelsDirectory))
            {
                return false;
            }
            
            return true;
        }
        
        /// <summary>
        /// Gets the available voices for the TTS engine
        /// </summary>
        /// <param name="configuration">The engine configuration</param>
        /// <returns>A collection of voice information</returns>
        public override Task<IEnumerable<TtsVoiceInfo>> GetAvailableVoicesAsync(Dictionary<string, string> configuration)
        {
            string modelsDirectory = GetConfigValue<string>(configuration, "modelsDirectory", string.Empty);
            
            if (string.IsNullOrEmpty(modelsDirectory) || !Directory.Exists(modelsDirectory))
            {
                return Task.FromResult<IEnumerable<TtsVoiceInfo>>(new List<TtsVoiceInfo>());
            }
            
            var voices = new List<TtsVoiceInfo>();
            
            // Look for model directories
            foreach (var directory in Directory.GetDirectories(modelsDirectory))
            {
                string dirName = Path.GetFileName(directory);
                
                // Check if this is a valid model directory
                if (File.Exists(Path.Combine(directory, "model.onnx")) && 
                    File.Exists(Path.Combine(directory, "tokens.txt")))
                {
                    voices.Add(new TtsVoiceInfo
                    {
                        Id = dirName,
                        Name = dirName,
                        Gender = "Unknown",
                        Language = "en-US",
                        Locale = "en-US"
                    });
                }
            }
            
            return Task.FromResult<IEnumerable<TtsVoiceInfo>>(voices);
        }
        
        /// <summary>
        /// Synthesizes speech from text
        /// </summary>
        /// <param name="text">The text to synthesize</param>
        /// <param name="voiceId">The ID of the voice to use</param>
        /// <param name="configuration">The engine configuration</param>
        /// <returns>The synthesized audio data</returns>
        public override Task<byte[]> SynthesizeSpeechAsync(string text, string voiceId, Dictionary<string, string> configuration)
        {
            string modelsDirectory = GetConfigValue<string>(configuration, "modelsDirectory", string.Empty);
            
            if (string.IsNullOrEmpty(modelsDirectory) || !Directory.Exists(modelsDirectory))
            {
                throw new InvalidOperationException("Models directory not found");
            }
            
            string voiceDirectory = Path.Combine(modelsDirectory, voiceId);
            
            if (!Directory.Exists(voiceDirectory))
            {
                throw new InvalidOperationException($"Voice directory not found: {voiceId}");
            }
            
            string modelPath = Path.Combine(voiceDirectory, "model.onnx");
            string tokensPath = Path.Combine(voiceDirectory, "tokens.txt");
            string lexiconPath = Path.Combine(voiceDirectory, "lexicon.txt");
            
            if (!File.Exists(modelPath))
            {
                throw new InvalidOperationException($"Model file not found: {modelPath}");
            }
            
            if (!File.Exists(tokensPath))
            {
                throw new InvalidOperationException($"Tokens file not found: {tokensPath}");
            }
            
            // Use the OpenSpeechTTS library to generate audio
            using (var tts = new global::OpenSpeechTTS.SherpaTTS(modelPath, tokensPath, lexiconPath, voiceDirectory))
            {
                byte[] audioData = tts.GenerateAudio(text);
                return Task.FromResult(audioData);
            }
        }
    }
} 