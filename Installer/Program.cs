using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Principal;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Installer;

class Program
{
    private const string OnlineJsonUrl = "https://github.com/willwade/tts-wrapper/raw/main/tts_wrapper/engines/sherpaonnx/merged_models.json";
    private const string LocalJsonPath = "./merged_models.json";

    static async Task Main(string[] args)
    {
        if (!IsRunningAsAdministrator())
        {
            Console.WriteLine("This application requires administrative privileges to install voices.");
            Console.WriteLine("Please run as administrator.");
            return;
        }

        var installer = new ModelInstaller();
        var registrar = new Sapi5Registrar();
        string dllPath = @"C:\Program Files\OpenAssistive\OpenSpeech\OpenSpeechTTS.dll";

        // Check for command line arguments
        if (args.Length >= 2)
        {
            string command = args[0].ToLower();
            string modelId = args[1];

            switch (command)
            {
                case "install":
                    await InstallSpecificVoice(modelId, installer, registrar, dllPath);
                    return;

                case "verify":
                    await VerifyVoiceInstallation(modelId);
                    return;

                case "uninstall":
                    if (modelId == "all")
                    {
                        await UninstallVoicesAndDll(registrar, dllPath);
                    }
                    else
                    {
                        await UninstallSpecificVoice(modelId, registrar);
                    }
                    return;

                default:
                    Console.WriteLine("Invalid command. Use 'install <model-id>', 'verify <model-id>', or 'uninstall <model-id|all>'");
                    return;
            }
        }

        // If no arguments provided, continue with interactive mode
        Console.WriteLine("Interactive Mode:");

        // Load models JSON
        var models = await LoadModelsAsync();
        Console.WriteLine("You can search for voices by:");
        Console.WriteLine(" - Language code (e.g., 'en', 'fs')");
        Console.WriteLine(" - Model name (e.g., 'xiaomaiiwn')");
        Console.WriteLine(" - Model ID (e.g., 'cantonese-fs-xiaomaiiwn')");
        Console.WriteLine(" - Model type (e.g., 'vits', 'mms', 'piper', 'coqui')");
        Console.WriteLine();

        while (true)
        {
            Console.Write("Enter a search term (or 'exit' to quit): ");
            string searchTerm = Console.ReadLine();

            if (string.IsNullOrWhiteSpace(searchTerm))
                continue;

            if (searchTerm.Equals("exit", StringComparison.OrdinalIgnoreCase))
                return;

            // Filter models
            var filteredModels = models.Values.Where(model =>
                (model.Id?.Contains(searchTerm, StringComparison.OrdinalIgnoreCase) ?? false) ||
                (model.Name?.Contains(searchTerm, StringComparison.OrdinalIgnoreCase) ?? false) ||
                (model.Developer?.Contains(searchTerm, StringComparison.OrdinalIgnoreCase) ?? false) ||
                (model.Language != null && model.Language.Any(lang =>
                    (lang.LangCode?.Contains(searchTerm, StringComparison.OrdinalIgnoreCase) ?? false) ||
                    (lang.LanguageName?.Contains(searchTerm, StringComparison.OrdinalIgnoreCase) ?? false)))
            );

            if (!filteredModels.Any())
            {
                Console.WriteLine("No voices matched your search. Try again.");
                continue;
            }

            // Display filtered results
            Console.WriteLine("Matching voices:");
            foreach (var model in filteredModels)
            {
                Console.WriteLine($" - {model.Id} ({model.Name}, Type: {model.ModelType}, Language: {string.Join(", ", model.Language.Select(l => l.LanguageName))})");
            }

            // Prompt user to install
            Console.Write("Enter the model ID to install (or 'search' to search again): ");
            string chosenModelId = Console.ReadLine();

            if (chosenModelId.Equals("search", StringComparison.OrdinalIgnoreCase))
                continue;

            if (models.TryGetValue(chosenModelId, out var chosenModel))
            {
                try
                {
                    Console.WriteLine($"Downloading and installing {chosenModel.Name}...");
                    await InstallSpecificVoice(chosenModelId, installer, registrar, dllPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to install {chosenModel.Name}: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine("Invalid model ID entered. Try again.");
            }
        }
    }

    static async Task<Dictionary<string, TtsModel>> LoadModelsAsync()
    {
        try
        {
            using var client = new System.Net.Http.HttpClient();
            Console.WriteLine("Downloading merged_models.json...");
            var json = await client.GetStringAsync(OnlineJsonUrl);
            File.WriteAllText(LocalJsonPath, json); // Cache the downloaded file
            Console.WriteLine("Successfully loaded merged_models.json from the web.");

            // Deserialize JSON into a dictionary of TtsModel
            return JsonConvert.DeserializeObject<Dictionary<string, TtsModel>>(json);
        }
        catch
        {
            Console.WriteLine("Failed to download merged_models.json. Using local copy.");
            if (!File.Exists(LocalJsonPath))
                throw new FileNotFoundException("Local merged_models.json not found.");

            var json = File.ReadAllText(LocalJsonPath);
            return JsonConvert.DeserializeObject<Dictionary<string, TtsModel>>(json);
        }
    }

    private static async Task UninstallVoicesAndDll(Sapi5Registrar registrar, string dllPath)
    {
        Console.WriteLine("Uninstalling voices...");

        var models = await LoadModelsAsync();
        foreach (var model in models.Values)
        {
            try
            {
                registrar.UnregisterVoice(model.Id);
                Console.WriteLine($"Unregistered voice: {model.Name}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error uninstalling {model.Name}: {ex.Message}");
            }
        }

        // Check if any models remain in Program Files
        string modelsDir = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
            "OpenSpeech",
            "models"
        );

        var modelDirs = Directory.Exists(modelsDir) ? Directory.GetDirectories(modelsDir) : Array.Empty<string>();
        if (modelDirs.Length == 0)
        {
            UnregisterDll(dllPath);

            // Clean up OpenSpeech directory if empty
            try
            {
                if (Directory.Exists(modelsDir))
                {
                    Directory.Delete(modelsDir, true);
                    Console.WriteLine($"Cleaned up models directory: {modelsDir}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error cleaning up models directory: {ex.Message}");
            }
        }
    }

    private static void UnregisterDll(string dllPath)
    {
        try
        {
            var process = new Process();
            process.StartInfo.FileName = "regasm";
            process.StartInfo.Arguments = $"/unregister \"{dllPath}\"";
            process.StartInfo.RedirectStandardOutput = true;
            process.StartInfo.UseShellExecute = false;
            process.Start();
            process.WaitForExit();

            Console.WriteLine("DLL successfully unregistered.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error unregistering DLL: {ex.Message}");
        }
    }

    private static async Task InstallSpecificVoice(string modelId, ModelInstaller installer, Sapi5Registrar registrar, string dllPath)
    {
        Console.WriteLine($"Installing voice: {modelId}");
        var models = await LoadModelsAsync();

        if (models.TryGetValue(modelId, out var model))
        {
            // Set gender based on model name/id if not already set
            if (string.IsNullOrEmpty(model.Gender))
            {
                if (model.Name?.Contains("female", StringComparison.OrdinalIgnoreCase) == true ||
                    model.Id?.Contains("female", StringComparison.OrdinalIgnoreCase) == true)
                {
                    model.Gender = "Female";
                }
                else if (model.Name?.Contains("male", StringComparison.OrdinalIgnoreCase) == true ||
                         model.Id?.Contains("male", StringComparison.OrdinalIgnoreCase) == true)
                {
                    model.Gender = "Male";
                }
            }

            // Debug language info
            var language = model.Language?.FirstOrDefault();
            if (language != null)
            {
                Console.WriteLine($"Language: {language.LanguageName} ({language.LangCode})");
            }

            try
            {
                Console.WriteLine($"Downloading and installing {model.Name}...");
                await installer.DownloadAndExtractModelAsync(model);
                
                // Register the COM DLL first
                RegisterComDll(dllPath);
                
                // Then register the voice
                registrar.RegisterVoice(model, dllPath);
                
                Console.WriteLine($"Successfully installed voice: {model.Name}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to install {model.Name}: {ex.Message}");
                throw;
            }
        }
        else
        {
            Console.WriteLine($"Model ID {modelId} not found in the models database.");
            throw new KeyNotFoundException($"Model ID {modelId} not found");
        }
    }

    // Add a method to register the COM DLL
    private static void RegisterComDll(string dllPath)
    {
        try
        {
            Console.WriteLine($"Registering COM DLL: {dllPath}");
            
            // Ensure the directory exists
            string dllDir = Path.GetDirectoryName(dllPath);
            if (!Directory.Exists(dllDir))
            {
                Directory.CreateDirectory(dllDir);
            }
            
            // Check if the DLL exists
            if (!File.Exists(dllPath))
            {
                throw new FileNotFoundException($"DLL not found: {dllPath}");
            }
            
            // Register the DLL with regasm
            var process = new Process();
            process.StartInfo.FileName = "regasm";
            process.StartInfo.Arguments = $"/codebase \"{dllPath}\"";
            process.StartInfo.RedirectStandardOutput = true;
            process.StartInfo.RedirectStandardError = true;
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.CreateNoWindow = true;
            
            process.Start();
            string output = process.StandardOutput.ReadToEnd();
            string error = process.StandardError.ReadToEnd();
            process.WaitForExit();
            
            if (process.ExitCode != 0)
            {
                Console.WriteLine($"regasm output: {output}");
                Console.WriteLine($"regasm error: {error}");
                throw new Exception($"Failed to register COM DLL. Exit code: {process.ExitCode}");
            }
            
            Console.WriteLine("COM DLL registered successfully");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error registering COM DLL: {ex.Message}");
            throw;
        }
    }

    private static async Task UninstallSpecificVoice(string modelId, Sapi5Registrar registrar)
    {
        try
        {
            var models = await LoadModelsAsync();
            if (models.TryGetValue(modelId, out var model))
            {
                registrar.UnregisterVoice(modelId);
                Console.WriteLine($"Successfully uninstalled voice: {modelId}");
            }
            else
            {
                Console.WriteLine($"Model ID '{modelId}' not found in available voices.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to uninstall voice: {ex.Message}");
        }
    }

    private static async Task VerifyVoiceInstallation(string modelId)
    {
        try
        {
            var models = await LoadModelsAsync();
            if (!models.TryGetValue(modelId, out var model))
            {
                Console.WriteLine($"Model ID '{modelId}' not found in available voices.");
                return;
            }

            Console.WriteLine($"Verifying installation for voice: {model.Name}");
            Console.WriteLine("----------------------------------------");

            // 1. Check model files
            string modelDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                "OpenSpeech",
                "models",
                model.Id
            );
            string modelPath = Path.Combine(modelDir, "model.onnx");
            string tokensPath = Path.Combine(modelDir, "tokens.txt");

            Console.WriteLine("\nChecking model files:");
            Console.WriteLine($"Model directory: {modelDir}");
            if (Directory.Exists(modelDir))
            {
                Console.WriteLine("✓ Model directory exists");
                if (File.Exists(modelPath))
                {
                    var modelSize = new FileInfo(modelPath).Length;
                    Console.WriteLine($"✓ model.onnx exists ({modelSize / 1024.0 / 1024.0:F2} MB)");
                }
                else
                {
                    Console.WriteLine("✗ model.onnx is missing");
                }

                if (File.Exists(tokensPath))
                {
                    var tokensSize = new FileInfo(tokensPath).Length;
                    Console.WriteLine($"✓ tokens.txt exists ({tokensSize / 1024.0:F2} KB)");
                }
                else
                {
                    Console.WriteLine("✗ tokens.txt is missing");
                }
            }
            else
            {
                Console.WriteLine("✗ Model directory is missing");
            }

            // 2. Check registry entries
            Console.WriteLine("\nChecking registry entries:");
            string registryPath = $@"SOFTWARE\Microsoft\Speech\Voices\Tokens\{model.Name}";
            using (var key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(registryPath))
            {
                if (key != null)
                {
                    Console.WriteLine("✓ Voice registry key exists");
                    
                    var lang = key.GetValue("Lang");
                    Console.WriteLine($"Language ID: {lang}");
                    
                    var gender = key.GetValue("Gender");
                    Console.WriteLine($"Gender: {gender}");
                    
                    var vendor = key.GetValue("Vendor");
                    Console.WriteLine($"Vendor: {vendor}");

                    using (var attribKey = key.OpenSubKey("Attributes"))
                    {
                        if (attribKey != null)
                        {
                            Console.WriteLine("\nAttribute values:");
                            Console.WriteLine($"CLSID: {attribKey.GetValue("CLSID")}");
                            Console.WriteLine($"Voice Path: {attribKey.GetValue("VoicePath")}");
                            Console.WriteLine($"Model Path: {attribKey.GetValue("ModelPath")}");
                            Console.WriteLine($"Tokens Path: {attribKey.GetValue("TokensPath")}");
                        }
                        else
                        {
                            Console.WriteLine("✗ Attributes subkey is missing");
                        }
                    }
                }
                else
                {
                    Console.WriteLine("✗ Voice registry key is missing");
                }
            }

            // 3. Test voice with SAPI
            Console.WriteLine("\nTesting voice with SAPI:");
            try
            {
                var synth = new System.Speech.Synthesis.SpeechSynthesizer();
                var voices = synth.GetInstalledVoices();
                var voice = voices.FirstOrDefault(v => v.VoiceInfo.Name == model.Name);
                
                if (voice != null)
                {
                    Console.WriteLine($"✓ Voice found in SAPI");
                    Console.WriteLine($"Voice Info:");
                    Console.WriteLine($"  Name: {voice.VoiceInfo.Name}");
                    Console.WriteLine($"  Culture: {voice.VoiceInfo.Culture}");
                    Console.WriteLine($"  Gender: {voice.VoiceInfo.Gender}");
                    Console.WriteLine($"  Age: {voice.VoiceInfo.Age}");
                    
                    Console.WriteLine("\nTesting speech synthesis...");
                    synth.SelectVoice(model.Name);
                    synth.SetOutputToDefaultAudioDevice();
                    synth.Speak("This is a test of the voice installation.");
                    Console.WriteLine("✓ Speech test completed");
                }
                else
                {
                    Console.WriteLine("✗ Voice not found in SAPI");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ Speech test failed: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Verification failed: {ex.Message}");
        }
    }

    private static bool IsRunningAsAdministrator()
    {
        using (var identity = System.Security.Principal.WindowsIdentity.GetCurrent())
        {
            var principal = new System.Security.Principal.WindowsPrincipal(identity);
            return principal.IsInRole(System.Security.Principal.WindowsBuiltInRole.Administrator);
        }
    }
}
