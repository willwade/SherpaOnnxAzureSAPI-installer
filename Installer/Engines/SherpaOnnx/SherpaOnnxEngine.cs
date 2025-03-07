using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Installer.Core.Base;
using Installer.Core.Models;
using Newtonsoft.Json;
using SherpaOnnx;

namespace Installer.Engines.SherpaOnnx
{
    /// <summary>
    /// Sherpa ONNX TTS engine implementation
    /// </summary>
    public class SherpaOnnxEngine : TtsEngineBase
    {
        private readonly string _modelsDirectory;
        private readonly string _modelsJsonFile;
        
        /// <summary>
        /// Creates a new instance of the SherpaOnnxEngine class
        /// </summary>
        public SherpaOnnxEngine()
        {
            _modelsDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OpenSpeech", "models");
            _modelsJsonFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "merged_models.json");
        }
        
        /// <summary>
        /// Creates a new instance of the SherpaOnnxEngine class with a custom models directory
        /// </summary>
        /// <param name="modelsDirectory">The directory containing the models</param>
        /// <param name="modelsJsonFile">The JSON file containing the model definitions</param>
        public SherpaOnnxEngine(string modelsDirectory, string modelsJsonFile)
        {
            _modelsDirectory = modelsDirectory;
            _modelsJsonFile = modelsJsonFile;
        }
        
        // Basic properties
        public override string EngineName => "SherpaOnnx";
        public override string EngineVersion => "1.0";
        public override string EngineDescription => "Sherpa ONNX Text-to-Speech Engine";
        public override bool RequiresSsml => false;
        public override bool RequiresAuthentication => false;
        public override bool SupportsOfflineUsage => true;
        
        // SAPI integration
        public override Guid GetEngineClsid() => new Guid("3d8f5c5d-9d6b-4b92-a12b-1a6dff80b6b2");
        public override Type GetSapiImplementationType() => typeof(SherpaOnnxSapi5VoiceImpl);
        
        // Voice management
        public override async Task<IEnumerable<TtsVoiceInfo>> GetAvailableVoicesAsync(Dictionary<string, string> config)
        {
            return await Task.Run(() =>
            {
                var voices = new List<TtsVoiceInfo>();
                
                try
                {
                    // Load models from JSON file
                    if (!File.Exists(_modelsJsonFile))
                    {
                        LogError($"Models JSON file not found: {_modelsJsonFile}");
                        return voices;
                    }
                    
                    string json = File.ReadAllText(_modelsJsonFile);
                    var models = JsonConvert.DeserializeObject<List<TtsModelDefinition>>(json);
                    
                    if (models == null || models.Count == 0)
                    {
                        LogError("No models found in JSON file");
                        return voices;
                    }
                    
                    // Convert models to voice info
                    foreach (var model in models)
                    {
                        var voice = new TtsVoiceInfo
                        {
                            Id = model.Id,
                            Name = model.Name,
                            DisplayName = model.Name,
                            Gender = model.Gender,
                            Locale = model.Language.FirstOrDefault()?.Code ?? "en-US",
                            EngineName = EngineName,
                            AdditionalAttributes = new Dictionary<string, string>
                            {
                                ["ModelType"] = model.ModelType,
                                ["Developer"] = model.Developer,
                                ["Quality"] = model.Quality,
                                ["SampleRate"] = model.SampleRate.ToString(),
                                ["NumSpeakers"] = model.NumSpeakers.ToString(),
                                ["Url"] = model.Url,
                                ["FilesizeMb"] = model.FilesizeMb.ToString()
                            }
                        };
                        
                        voices.Add(voice);
                    }
                }
                catch (Exception ex)
                {
                    LogError("Error getting available voices", ex);
                }
                
                return voices;
            });
        }
        
        // Speech synthesis
        public override async Task<byte[]> SynthesizeSpeechAsync(string text, string voiceId, Dictionary<string, string> parameters)
        {
            return await Task.Run(() =>
            {
                try
                {
                    // Get model paths
                    string modelPath = GetModelPath(voiceId);
                    string tokensPath = GetTokensPath(voiceId);
                    string lexiconPath = GetLexiconPath(voiceId);
                    string dataDir = GetDataDir();
                    
                    // Create Sherpa TTS instance
                    using (var tts = new OpenSpeechTTS.SherpaTTS(modelPath, tokensPath, lexiconPath, dataDir))
                    {
                        // Generate audio
                        return tts.GenerateAudio(text);
                    }
                }
                catch (Exception ex)
                {
                    LogError($"Error synthesizing speech for voice {voiceId}", ex);
                    return new byte[0];
                }
            });
        }
        
        // Helper methods
        private string GetModelPath(string voiceId)
        {
            return Path.Combine(_modelsDirectory, voiceId, "model.onnx");
        }
        
        private string GetTokensPath(string voiceId)
        {
            return Path.Combine(_modelsDirectory, voiceId, "tokens.txt");
        }
        
        private string GetLexiconPath(string voiceId)
        {
            string lexiconPath = Path.Combine(_modelsDirectory, voiceId, "lexicon.txt");
            return File.Exists(lexiconPath) ? lexiconPath : null;
        }
        
        private string GetDataDir()
        {
            return _modelsDirectory;
        }
        
        // Model definition class
        private class TtsModelDefinition
        {
            public string Id { get; set; }
            public string ModelType { get; set; }
            public string Developer { get; set; }
            public string Name { get; set; }
            public List<LanguageInfo> Language { get; set; }
            public string Quality { get; set; }
            public int SampleRate { get; set; }
            public int NumSpeakers { get; set; }
            public string Url { get; set; }
            public bool Compression { get; set; }
            public double FilesizeMb { get; set; }
            public string Gender { get; set; }
        }
        
        private class LanguageInfo
        {
            public string Name { get; set; }
            public string Code { get; set; }
        }
    }
} 