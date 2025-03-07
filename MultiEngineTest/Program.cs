using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Speech.Synthesis;
using System.Threading.Tasks;
using NAudio.Wave;
using Newtonsoft.Json;
using OpenSpeechTTS;

namespace MultiEngineTest
{
    class Program
    {
        private static string _outputDirectory = "C:\\OpenSpeech\\TestOutput";
        private static string _sapiDllPath = "C:\\Program Files\\OpenAssistive\\OpenSpeech\\OpenSpeechTTS.dll";
        private static string _pluginDirectory = "C:\\Program Files\\OpenAssistive\\OpenSpeech\\plugins";
        private static ConfigurationManager _configManager;
        private static TtsEngineManager _engineManager;
        private static PluginLoader _pluginLoader;
        private static MultiEngineTestConfig _testConfig;
        
        static async Task Main(string[] args)
        {
            Console.WriteLine("OpenSpeech Multi-Engine TTS Test");
            Console.WriteLine("================================");
            Console.WriteLine();
            
            // Load test configuration
            LoadTestConfig();
            
            // Create output directory if it doesn't exist
            if (!Directory.Exists(_outputDirectory))
            {
                Directory.CreateDirectory(_outputDirectory);
            }
            
            // Run tests based on configuration
            if (_testConfig.TestOptions.RunDirectTests)
            {
                Console.WriteLine("Running Direct Tests...");
                await RunDirectTests();
                Console.WriteLine();
            }
            
            if (_testConfig.TestOptions.RunSapiTests)
            {
                Console.WriteLine("Running SAPI Tests...");
                await RunSapiTests();
                Console.WriteLine();
            }
            
            if (_testConfig.TestOptions.RunPluginTests)
            {
                Console.WriteLine("Running Plugin System Tests...");
                await RunPluginSystemTests();
                Console.WriteLine();
            }
            
            Console.WriteLine("All tests completed.");
        }
        
        private static void LoadTestConfig()
        {
            try
            {
                string configPath = "test-config.json";
                
                if (!File.Exists(configPath))
                {
                    Console.WriteLine("Test configuration file not found. Using default settings.");
                    _testConfig = new MultiEngineTestConfig();
                    return;
                }
                
                string json = File.ReadAllText(configPath);
                _testConfig = JsonConvert.DeserializeObject<MultiEngineTestConfig>(json);
                
                if (_testConfig == null)
                {
                    Console.WriteLine("Failed to parse test configuration. Using default settings.");
                    _testConfig = new MultiEngineTestConfig();
                    return;
                }
                
                // Update settings from config
                _outputDirectory = _testConfig.TestOptions.OutputDirectory;
                _sapiDllPath = _testConfig.TestOptions.SapiDllPath;
                _pluginDirectory = _testConfig.TestOptions.PluginDirectory;
                
                Console.WriteLine("Test configuration loaded successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading test configuration: {ex.Message}");
                _testConfig = new MultiEngineTestConfig();
            }
        }
        
        private static async Task RunDirectTests()
        {
            // Test Sherpa ONNX
            if (_testConfig.Engines.ContainsKey("SherpaOnnx") && _testConfig.Engines["SherpaOnnx"].Enabled)
            {
                await TestSherpaOnnxDirect();
            }
            
            // Test Azure TTS
            if (_testConfig.Engines.ContainsKey("AzureTTS") && _testConfig.Engines["AzureTTS"].Enabled)
            {
                await TestAzureTtsDirect();
            }
            
            // Test ElevenLabs
            if (_testConfig.Engines.ContainsKey("ElevenLabs") && _testConfig.Engines["ElevenLabs"].Enabled)
            {
                await TestElevenLabsDirect();
            }
            
            // Test PlayHT
            if (_testConfig.Engines.ContainsKey("PlayHT") && _testConfig.Engines["PlayHT"].Enabled)
            {
                await TestPlayHtDirect();
            }
        }
        
        private static async Task TestSherpaOnnxDirect()
        {
            try
            {
                Console.WriteLine("Testing Sherpa ONNX directly...");
                
                var engineConfig = _testConfig.Engines["SherpaOnnx"];
                var parameters = engineConfig.Parameters;
                string modelsDirectory = parameters.ContainsKey("modelsDirectory") ? parameters["modelsDirectory"] : "C:\\OpenSpeech\\models";
                string voiceId = engineConfig.TestVoiceId;
                string text = engineConfig.TestText;
                
                // Get model paths
                string modelPath = Path.Combine(modelsDirectory, voiceId, "model.onnx");
                string tokensPath = Path.Combine(modelsDirectory, voiceId, "tokens.txt");
                string lexiconPath = Path.Combine(modelsDirectory, voiceId, "lexicon.txt");
                
                if (!File.Exists(modelPath))
                {
                    Console.WriteLine($"Model file not found: {modelPath}");
                    return;
                }
                
                if (!File.Exists(tokensPath))
                {
                    Console.WriteLine($"Tokens file not found: {tokensPath}");
                    return;
                }
                
                // Create Sherpa TTS instance
                using (var tts = new SherpaTTS(modelPath, tokensPath, lexiconPath, modelsDirectory))
                {
                    Console.WriteLine($"Generating audio for text: \"{text}\"");
                    
                    // Generate audio
                    byte[] audioData = tts.GenerateAudio(text);
                    
                    // Save audio to file
                    string outputFile = Path.Combine(_outputDirectory, $"sherpa-onnx-direct-{voiceId}.wav");
                    File.WriteAllBytes(outputFile, audioData);
                    
                    Console.WriteLine($"Audio saved to: {outputFile}");
                    
                    // Play audio
                    PlayAudio(audioData);
                }
                
                Console.WriteLine("Sherpa ONNX direct test completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error testing Sherpa ONNX directly: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
        
        private static async Task TestAzureTtsDirect()
        {
            try
            {
                Console.WriteLine("Testing Azure TTS directly...");
                
                var engineConfig = _testConfig.Engines["AzureTTS"];
                var parameters = engineConfig.Parameters;
                string subscriptionKey = parameters.ContainsKey("subscriptionKey") ? parameters["subscriptionKey"] : "";
                string region = parameters.ContainsKey("region") ? parameters["region"] : "eastus";
                string voiceId = engineConfig.TestVoiceId;
                string text = engineConfig.TestText;
                
                if (string.IsNullOrEmpty(subscriptionKey))
                {
                    Console.WriteLine("Azure subscription key not provided in test configuration.");
                    return;
                }
                
                // Create Azure TTS instance
                using (var tts = new AzureTTS(subscriptionKey, region, voiceId))
                {
                    Console.WriteLine($"Generating audio for text: \"{text}\"");
                    
                    // Generate audio
                    byte[] audioData = await tts.SynthesizeSpeechAsync(text);
                    
                    // Save audio to file
                    string outputFile = Path.Combine(_outputDirectory, $"azure-direct-{voiceId}.wav");
                    File.WriteAllBytes(outputFile, audioData);
                    
                    Console.WriteLine($"Audio saved to: {outputFile}");
                    
                    // Play audio
                    PlayAudio(audioData);
                }
                
                Console.WriteLine("Azure TTS direct test completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error testing Azure TTS directly: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
        
        private static async Task TestElevenLabsDirect()
        {
            try
            {
                Console.WriteLine("Testing ElevenLabs directly...");
                
                var engineConfig = _testConfig.Engines["ElevenLabs"];
                var parameters = engineConfig.Parameters;
                string apiKey = parameters.ContainsKey("apiKey") ? parameters["apiKey"] : "";
                string modelId = parameters.ContainsKey("modelId") ? parameters["modelId"] : "eleven_monolingual_v1";
                string voiceId = engineConfig.TestVoiceId;
                string text = engineConfig.TestText;
                
                if (string.IsNullOrEmpty(apiKey))
                {
                    Console.WriteLine("ElevenLabs API key not provided in test configuration.");
                    return;
                }
                
                // Create HTTP client
                using (var httpClient = new System.Net.Http.HttpClient())
                {
                    Console.WriteLine($"Generating audio for text: \"{text}\"");
                    
                    // Create request body
                    var requestBody = new
                    {
                        text = text,
                        model_id = modelId,
                        voice_settings = new
                        {
                            stability = 0.5,
                            similarity_boost = 0.5
                        }
                    };
                    
                    string synthesisUrl = $"https://api.elevenlabs.io/v1/text-to-speech/{voiceId}";
                    
                    using (var request = new System.Net.Http.HttpRequestMessage(System.Net.Http.HttpMethod.Post, synthesisUrl))
                    {
                        request.Headers.Add("xi-api-key", apiKey);
                        request.Headers.Add("Accept", "audio/mpeg");
                        
                        string json = JsonConvert.SerializeObject(requestBody);
                        request.Content = new System.Net.Http.StringContent(json, System.Text.Encoding.UTF8, "application/json");
                        
                        using (var response = await httpClient.SendAsync(request))
                        {
                            response.EnsureSuccessStatusCode();
                            byte[] audioData = await response.Content.ReadAsByteArrayAsync();
                            
                            // Save audio to file
                            string outputFile = Path.Combine(_outputDirectory, $"elevenlabs-direct-{voiceId}.mp3");
                            File.WriteAllBytes(outputFile, audioData);
                            
                            Console.WriteLine($"Audio saved to: {outputFile}");
                            
                            // Play audio
                            PlayAudio(audioData, true);
                        }
                    }
                }
                
                Console.WriteLine("ElevenLabs direct test completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error testing ElevenLabs directly: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
        
        private static async Task TestPlayHtDirect()
        {
            try
            {
                Console.WriteLine("Testing PlayHT directly...");
                
                var engineConfig = _testConfig.Engines["PlayHT"];
                var parameters = engineConfig.Parameters;
                string apiKey = parameters.ContainsKey("apiKey") ? parameters["apiKey"] : "";
                string userId = parameters.ContainsKey("userId") ? parameters["userId"] : "";
                string quality = parameters.ContainsKey("quality") ? parameters["quality"] : "premium";
                string voiceId = engineConfig.TestVoiceId;
                string text = engineConfig.TestText;
                
                if (string.IsNullOrEmpty(apiKey) || string.IsNullOrEmpty(userId))
                {
                    Console.WriteLine("PlayHT API key or user ID not provided in test configuration.");
                    return;
                }
                
                // Create HTTP client
                using (var httpClient = new System.Net.Http.HttpClient())
                {
                    Console.WriteLine($"Generating audio for text: \"{text}\"");
                    
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
                    
                    string conversionUrl = "https://api.play.ht/api/v2/tts";
                    
                    using (var request = new System.Net.Http.HttpRequestMessage(System.Net.Http.HttpMethod.Post, conversionUrl))
                    {
                        request.Headers.Add("X-API-KEY", apiKey);
                        request.Headers.Add("AUTHORIZATION", userId);
                        request.Headers.Add("Accept", "application/json");
                        
                        string json = JsonConvert.SerializeObject(requestBody);
                        request.Content = new System.Net.Http.StringContent(json, System.Text.Encoding.UTF8, "application/json");
                        
                        using (var response = await httpClient.SendAsync(request))
                        {
                            response.EnsureSuccessStatusCode();
                            string responseJson = await response.Content.ReadAsStringAsync();
                            var conversionResponse = JsonConvert.DeserializeObject<PlayHtConversionResponse>(responseJson);
                            
                            if (conversionResponse == null || string.IsNullOrEmpty(conversionResponse.AudioUrl))
                            {
                                Console.WriteLine("Failed to get audio URL from PlayHT API.");
                                return;
                            }
                            
                            // Download audio
                            using (var audioResponse = await httpClient.GetAsync(conversionResponse.AudioUrl))
                            {
                                audioResponse.EnsureSuccessStatusCode();
                                byte[] audioData = await audioResponse.Content.ReadAsByteArrayAsync();
                                
                                // Save audio to file
                                string outputFile = Path.Combine(_outputDirectory, $"playht-direct-{Path.GetFileName(voiceId)}.mp3");
                                File.WriteAllBytes(outputFile, audioData);
                                
                                Console.WriteLine($"Audio saved to: {outputFile}");
                                
                                // Play audio
                                PlayAudio(audioData, true);
                            }
                        }
                    }
                }
                
                Console.WriteLine("PlayHT direct test completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error testing PlayHT directly: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
        
        private static async Task RunSapiTests()
        {
            try
            {
                Console.WriteLine("Testing SAPI integration...");
                
                // Initialize speech synthesizer
                using (var synthesizer = new SpeechSynthesizer())
                {
                    // Get installed voices
                    var installedVoices = synthesizer.GetInstalledVoices();
                    
                    Console.WriteLine($"Found {installedVoices.Count} installed voices:");
                    
                    foreach (var voice in installedVoices)
                    {
                        var info = voice.VoiceInfo;
                        Console.WriteLine($"- {info.Name} ({info.Gender}, {info.Culture}, {info.Age})");
                        
                        // Test each voice
                        string text = "This is a test of the SAPI voice.";
                        string outputFile = Path.Combine(_outputDirectory, $"sapi-{info.Name}.wav");
                        
                        try
                        {
                            Console.WriteLine($"Testing voice: {info.Name}");
                            synthesizer.SelectVoice(info.Name);
                            synthesizer.SetOutputToWaveFile(outputFile);
                            synthesizer.Speak(text);
                            synthesizer.SetOutputToNull();
                            
                            Console.WriteLine($"Audio saved to: {outputFile}");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error testing voice {info.Name}: {ex.Message}");
                        }
                    }
                }
                
                Console.WriteLine("SAPI tests completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error running SAPI tests: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
        
        private static async Task RunPluginSystemTests()
        {
            try
            {
                Console.WriteLine("Initializing plugin system...");
                
                // Create plugin directory if it doesn't exist
                if (!Directory.Exists(_pluginDirectory))
                {
                    Directory.CreateDirectory(_pluginDirectory);
                }
                
                // Initialize configuration manager
                _configManager = new ConfigurationManager();
                
                // Initialize engine manager
                _engineManager = new TtsEngineManager(_configManager);
                
                // Initialize plugin loader
                _pluginLoader = new PluginLoader(_pluginDirectory, _engineManager);
                
                // Register built-in engines
                _engineManager.RegisterEngine(new SherpaOnnxEngine());
                _engineManager.RegisterEngine(new AzureTtsEngine());
                _engineManager.RegisterEngine(new ElevenLabsEngine());
                _engineManager.RegisterEngine(new PlayHTEngine());
                
                // Load plugins
                _pluginLoader.LoadAllEngines();
                
                Console.WriteLine("Plugin system initialized successfully.");
                Console.WriteLine($"Registered engines: {string.Join(", ", _engineManager.GetEngineNames())}");
                
                // Test each engine
                foreach (var engine in _engineManager.GetAllEngines())
                {
                    await TestPluginEngine(engine);
                }
                
                Console.WriteLine("Plugin system tests completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error running plugin system tests: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
        
        private static async Task TestPluginEngine(ITtsEngine engine)
        {
            try
            {
                Console.WriteLine($"Testing engine: {engine.EngineName}");
                
                // Skip if engine is not enabled in config
                if (!_testConfig.Engines.ContainsKey(engine.EngineName) || !_testConfig.Engines[engine.EngineName].Enabled)
                {
                    Console.WriteLine($"Engine {engine.EngineName} is disabled in test configuration. Skipping.");
                    return;
                }
                
                var engineConfig = _testConfig.Engines[engine.EngineName];
                var parameters = engineConfig.Parameters;
                string voiceId = engineConfig.TestVoiceId;
                string text = engineConfig.TestText;
                
                // Configure engine
                _configManager.UpdateEngineConfiguration(engine.EngineName, parameters);
                var config = _configManager.GetEngineConfiguration(engine.EngineName);
                
                // Validate configuration
                if (!engine.ValidateConfiguration(config))
                {
                    Console.WriteLine($"Invalid configuration for engine {engine.EngineName}. Skipping.");
                    return;
                }
                
                // Get available voices
                Console.WriteLine($"Getting available voices for {engine.EngineName}...");
                var voices = await engine.GetAvailableVoicesAsync(config);
                Console.WriteLine($"Found {voices.Count()} voices.");
                
                // Find test voice
                var testVoice = voices.FirstOrDefault(v => v.Id == voiceId);
                
                if (testVoice == null)
                {
                    Console.WriteLine($"Test voice {voiceId} not found for engine {engine.EngineName}. Skipping.");
                    return;
                }
                
                // Synthesize speech
                Console.WriteLine($"Synthesizing speech for text: \"{text}\"");
                byte[] audioData = await engine.SynthesizeSpeechAsync(text, voiceId, config);
                
                // Save audio to file
                string outputFile = Path.Combine(_outputDirectory, $"plugin-{engine.EngineName}-{voiceId}.wav");
                File.WriteAllBytes(outputFile, audioData);
                
                Console.WriteLine($"Audio saved to: {outputFile}");
                
                // Play audio
                PlayAudio(audioData, engine.EngineName != "SherpaOnnx" && engine.EngineName != "AzureTTS");
                
                Console.WriteLine($"Engine {engine.EngineName} test completed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error testing engine {engine.EngineName}: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
        
        private static void PlayAudio(byte[] audioData, bool isMp3 = false)
        {
            try
            {
                using (var stream = new MemoryStream(audioData))
                {
                    IWaveProvider waveProvider;
                    
                    if (isMp3)
                    {
                        waveProvider = new Mp3FileReader(stream);
                    }
                    else
                    {
                        waveProvider = new WaveFileReader(stream);
                    }
                    
                    using (var waveOut = new WaveOutEvent())
                    {
                        waveOut.Init(waveProvider);
                        waveOut.Play();
                        
                        // Wait for playback to complete
                        while (waveOut.PlaybackState == PlaybackState.Playing)
                        {
                            System.Threading.Thread.Sleep(100);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error playing audio: {ex.Message}");
            }
        }
    }
} 