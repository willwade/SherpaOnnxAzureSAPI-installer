using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Principal;
using System.Threading.Tasks;
using Microsoft.Win32;
using Newtonsoft.Json;
using Installer.Shared;

namespace Installer
{
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
            var registrar = new Sapi5RegistrarExtended();
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

                    case "install-azure":
                        string subscriptionKey = null;
                        string region = null;
                        string voiceName = modelId;
                        string style = null;
                        string role = null;

                        // Parse additional arguments for Azure
                        for (int i = 2; i < args.Length; i++)
                        {
                            if (args[i] == "--key" && i + 1 < args.Length)
                            {
                                subscriptionKey = args[i + 1];
                                i++;
                            }
                            else if (args[i] == "--region" && i + 1 < args.Length)
                            {
                                region = args[i + 1];
                                i++;
                            }
                            else if (args[i] == "--style" && i + 1 < args.Length)
                            {
                                style = args[i + 1];
                                i++;
                            }
                            else if (args[i] == "--role" && i + 1 < args.Length)
                            {
                                role = args[i + 1];
                                i++;
                            }
                        }

                        // If key or region not provided, try to load from config
                        if (string.IsNullOrEmpty(subscriptionKey) || string.IsNullOrEmpty(region))
                        {
                            var config = AzureConfigManager.LoadConfig();
                            
                            if (string.IsNullOrEmpty(subscriptionKey) && !string.IsNullOrEmpty(config.DefaultKey))
                            {
                                subscriptionKey = config.DefaultKey;
                                Console.WriteLine("Using subscription key from configuration file.");
                            }
                            
                            if (string.IsNullOrEmpty(region) && !string.IsNullOrEmpty(config.DefaultRegion))
                            {
                                region = config.DefaultRegion;
                                Console.WriteLine("Using region from configuration file.");
                            }
                        }

                        if (string.IsNullOrEmpty(subscriptionKey) || string.IsNullOrEmpty(region))
                        {
                            Console.WriteLine("Error: Azure subscription key and region are required.");
                            Console.WriteLine("Usage: Installer.exe install-azure <voice-name> --key <subscription-key> --region <region> [--style <style>] [--role <role>]");
                            Console.WriteLine("Or set up a configuration file using: Installer.exe save-azure-config --key <subscription-key> --region <region>");
                            return;
                        }

                        await InstallAzureVoice(voiceName, subscriptionKey, region, style, role, registrar, dllPath);
                        return;

                    case "list-azure-voices":
                        string listSubscriptionKey = null;
                        string listRegion = null;

                        // Parse additional arguments for Azure
                        for (int i = 2; i < args.Length; i++)
                        {
                            if (args[i] == "--key" && i + 1 < args.Length)
                            {
                                listSubscriptionKey = args[i + 1];
                                i++;
                            }
                            else if (args[i] == "--region" && i + 1 < args.Length)
                            {
                                listRegion = args[i + 1];
                                i++;
                            }
                        }

                        // If key or region not provided, try to load from config
                        if (string.IsNullOrEmpty(listSubscriptionKey) || string.IsNullOrEmpty(listRegion))
                        {
                            var config = AzureConfigManager.LoadConfig();
                            
                            if (string.IsNullOrEmpty(listSubscriptionKey) && !string.IsNullOrEmpty(config.DefaultKey))
                            {
                                listSubscriptionKey = config.DefaultKey;
                                Console.WriteLine("Using subscription key from configuration file.");
                            }
                            
                            if (string.IsNullOrEmpty(listRegion) && !string.IsNullOrEmpty(config.DefaultRegion))
                            {
                                listRegion = config.DefaultRegion;
                                Console.WriteLine("Using region from configuration file.");
                            }
                        }

                        if (string.IsNullOrEmpty(listSubscriptionKey) || string.IsNullOrEmpty(listRegion))
                        {
                            Console.WriteLine("Error: Azure subscription key and region are required.");
                            Console.WriteLine("Usage: Installer.exe list-azure-voices --key <subscription-key> --region <region>");
                            Console.WriteLine("Or set up a configuration file using: Installer.exe save-azure-config --key <subscription-key> --region <region>");
                            return;
                        }

                        await ListAzureVoices(listSubscriptionKey, listRegion);
                        return;

                    case "save-azure-config":
                        string configKey = null;
                        string configRegion = null;
                        bool secureStorage = true;
                        
                        // Parse additional arguments
                        for (int i = 1; i < args.Length; i++)
                        {
                            if (args[i] == "--key" && i + 1 < args.Length)
                            {
                                configKey = args[i + 1];
                                i++;
                            }
                            else if (args[i] == "--region" && i + 1 < args.Length)
                            {
                                configRegion = args[i + 1];
                                i++;
                            }
                            else if (args[i] == "--secure" && i + 1 < args.Length)
                            {
                                bool.TryParse(args[i + 1], out secureStorage);
                                i++;
                            }
                        }
                        
                        if (string.IsNullOrEmpty(configKey) || string.IsNullOrEmpty(configRegion))
                        {
                            Console.WriteLine("Error: Azure subscription key and region are required.");
                            Console.WriteLine("Usage: Installer.exe save-azure-config --key <subscription-key> --region <region> [--secure <true|false>]");
                            return;
                        }
                        
                        AzureConfigManager.SaveConfig(configKey, configRegion, secureStorage);
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

            // Interactive mode
            Console.WriteLine("OpenSpeech TTS SAPI Installer");
            Console.WriteLine("============================");
            Console.WriteLine();
            Console.WriteLine("Select an option:");
            Console.WriteLine("1. Install Sherpa ONNX voice");
            Console.WriteLine("2. Install Azure TTS voice");
            Console.WriteLine("3. Uninstall all voices");
            Console.WriteLine("4. Exit");
            Console.WriteLine();
            Console.Write("Enter your choice (1-4): ");
            
            string choice = Console.ReadLine();
            
            switch (choice)
            {
                case "1":
                    await InstallSherpaOnnxVoiceInteractive(installer, registrar, dllPath);
                    break;
                    
                case "2":
                    await InstallAzureVoiceInteractive(registrar, dllPath);
                    break;
                    
                case "3":
                    await UninstallVoicesAndDll(registrar, dllPath);
                    break;
                    
                case "4":
                    return;
                    
                default:
                    Console.WriteLine("Invalid choice. Please try again.");
                    break;
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

        private static async Task UninstallVoicesAndDll(Sapi5RegistrarExtended registrar, string dllPath)
        {
            Console.WriteLine("Uninstalling voices...");

            // Uninstall Sherpa ONNX voices
            var models = await LoadModelsAsync();
            foreach (var model in models.Values)
            {
                try
                {
                    registrar.UnregisterVoice(model.Id);
                    Console.WriteLine($"Unregistered Sherpa ONNX voice: {model.Name}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error uninstalling {model.Name}: {ex.Message}");
                }
            }
            
            // Uninstall Azure voices
            try
            {
                // Get all registry keys under SPEECH\Voices\Tokens
                using (var voicesKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\SPEECH\Voices\Tokens", false))
                {
                    if (voicesKey != null)
                    {
                        string[] voiceNames = voicesKey.GetSubKeyNames();
                        foreach (string voiceName in voiceNames)
                        {
                            try
                            {
                                // Check if this is an Azure voice by looking for the Azure CLSID
                                using (var voiceKey = voicesKey.OpenSubKey(voiceName))
                                {
                                    if (voiceKey != null)
                                    {
                                        string clsid = voiceKey.GetValue("CLSID") as string;
                                        if (clsid == "{3d8f5c5e-9d6b-4b92-a12b-1a6dff80b6b3}") // Azure CLSID
                                        {
                                            registrar.UnregisterVoice(voiceName);
                                            Console.WriteLine($"Unregistered Azure voice: {voiceName}");
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Error uninstalling voice {voiceName}: {ex.Message}");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error uninstalling Azure voices: {ex.Message}");
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
                    
                    // Try to clean up parent directory if it's empty
                    string openSpeechDir = Path.GetDirectoryName(modelsDir);
                    if (Directory.Exists(openSpeechDir) && Directory.GetFileSystemEntries(openSpeechDir).Length == 0)
                    {
                        Directory.Delete(openSpeechDir);
                        Console.WriteLine($"Cleaned up OpenSpeech directory: {openSpeechDir}");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error cleaning up directories: {ex.Message}");
                }
            }
            else
            {
                Console.WriteLine("Some models remain installed. DLL not unregistered.");
            }
        }

        private static void UnregisterDll(string dllPath)
        {
            try
            {
                // Find the regasm path
                string regasmPath = FindRegasmPath();
                
                var process = new Process();
                process.StartInfo.FileName = regasmPath;
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

        // Helper method to find the regasm path
        private static string FindRegasmPath()
        {
            // First try to get the path from the registry
            string regasmPath = GetRegasmPathFromRegistry();
            if (!string.IsNullOrEmpty(regasmPath) && File.Exists(regasmPath))
            {
                return regasmPath;
            }

            // If registry lookup fails, try common installation paths
            // Try 64-bit .NET Framework 4.x
            regasmPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Windows),
                "Microsoft.NET", "Framework64", "v4.0.30319", "regasm.exe");

            if (File.Exists(regasmPath))
            {
                return regasmPath;
            }

            // Try 32-bit .NET Framework 4.x
            regasmPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Windows),
                "Microsoft.NET", "Framework", "v4.0.30319", "regasm.exe");

            if (File.Exists(regasmPath))
            {
                return regasmPath;
            }

            // If we still can't find it, look for any version in Framework64 directory
            string frameworkDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Windows),
                "Microsoft.NET", "Framework64");

            if (Directory.Exists(frameworkDir))
            {
                foreach (var dir in Directory.GetDirectories(frameworkDir, "v*")
                                    .OrderByDescending(d => d)) // Get newest version first
                {
                    regasmPath = Path.Combine(dir, "regasm.exe");
                    if (File.Exists(regasmPath))
                    {
                        return regasmPath;
                    }
                }
            }

            // Try the same with 32-bit Framework directory
            frameworkDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Windows),
                "Microsoft.NET", "Framework");

            if (Directory.Exists(frameworkDir))
            {
                foreach (var dir in Directory.GetDirectories(frameworkDir, "v*")
                                    .OrderByDescending(d => d)) // Get newest version first
                {
                    regasmPath = Path.Combine(dir, "regasm.exe");
                    if (File.Exists(regasmPath))
                    {
                        return regasmPath;
                    }
                }
            }

            // If we still can't find it, throw an exception
            throw new FileNotFoundException(
                "Could not find regasm.exe. Please ensure that .NET Framework 4.x is installed on this system.");
        }

        private static string GetRegasmPathFromRegistry()
        {
            try
            {
                // Look for the .NET Framework installation directory in the registry
                using (var key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\.NETFramework"))
                {
                    if (key != null)
                    {
                        var installRoot = key.GetValue("InstallRoot") as string;
                        if (!string.IsNullOrEmpty(installRoot))
                        {
                            // Try to find the v4.0.30319 directory (or the latest version)
                            var frameworkDirs = Directory.GetDirectories(installRoot, "v*")
                                                        .OrderByDescending(d => d);
                            
                            foreach (var dir in frameworkDirs)
                            {
                                var regasmPath = Path.Combine(dir, "regasm.exe");
                                if (File.Exists(regasmPath))
                                {
                                    return regasmPath;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error looking up regasm path in registry: {ex.Message}");
            }
            
            return null;
        }

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
                
                // Find the regasm path
                string regasmPath = FindRegasmPath();
                
                // Register the DLL with regasm
                var process = new Process();
                process.StartInfo.FileName = regasmPath;
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

        private static async Task InstallAzureVoice(string voiceName, string subscriptionKey, string region, string style, string role, Sapi5RegistrarExtended registrar, string dllPath)
        {
            try
            {
                Console.WriteLine($"Installing Azure voice: {voiceName}");
                
                // Create Azure TTS service
                var azureService = new AzureTtsService(subscriptionKey, region);
                
                // Validate subscription
                Console.WriteLine("Validating Azure subscription...");
                bool isValid = await azureService.ValidateSubscriptionAsync();
                if (!isValid)
                {
                    Console.WriteLine("Error: Invalid Azure subscription key or region.");
                    return;
                }
                
                // Get available voices
                Console.WriteLine("Fetching available Azure voices...");
                var voices = await azureService.GetAvailableVoicesAsync();
                
                // Find the requested voice
                var voice = voices.FirstOrDefault(v => 
                    v.ShortName.Equals(voiceName, StringComparison.OrdinalIgnoreCase) || 
                    v.Name.Equals(voiceName, StringComparison.OrdinalIgnoreCase));
                
                if (voice == null)
                {
                    Console.WriteLine($"Error: Azure voice '{voiceName}' not found.");
                    Console.WriteLine("Available voices:");
                    foreach (var v in voices)
                    {
                        Console.WriteLine($" - {v.ShortName} ({v.DisplayName}, {v.Locale})");
                    }
                    return;
                }
                
                // Set style and role if specified
                if (!string.IsNullOrEmpty(style))
                {
                    if (voice.StyleList.Contains(style))
                    {
                        voice.SelectedStyle = style;
                        Console.WriteLine($"Using style: {style}");
                    }
                    else
                    {
                        Console.WriteLine($"Warning: Style '{style}' not available for this voice. Available styles: {string.Join(", ", voice.StyleList.ToArray())}");
                    }
                }
                
                if (!string.IsNullOrEmpty(role))
                {
                    if (voice.RoleList.Contains(role))
                    {
                        voice.SelectedRole = role;
                        Console.WriteLine($"Using role: {role}");
                    }
                    else
                    {
                        Console.WriteLine($"Warning: Role '{role}' not available for this voice. Available roles: {string.Join(", ", voice.RoleList.ToArray())}");
                    }
                }
                
                // Register the voice
                Console.WriteLine($"Registering Azure voice: {voice.ShortName}");
                registrar.RegisterAzureVoice(voice, dllPath);
                
                Console.WriteLine($"Azure voice '{voice.ShortName}' installed successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error installing Azure voice: {ex.Message}");
            }
        }

        private static async Task ListAzureVoices(string subscriptionKey, string region)
        {
            try
            {
                Console.WriteLine("Fetching available Azure voices...");
                
                // Create Azure TTS service
                var azureService = new AzureTtsService(subscriptionKey, region);
                
                // Validate subscription
                Console.WriteLine("Validating Azure subscription...");
                bool isValid = await azureService.ValidateSubscriptionAsync();
                if (!isValid)
                {
                    Console.WriteLine("Error: Invalid Azure subscription key or region.");
                    return;
                }
                
                // Get available voices
                var voices = await azureService.GetAvailableVoicesAsync();
                var voicesList = voices.ToList(); // Convert to list before accessing Count
                
                // Display voices
                Console.WriteLine($"Found {voicesList.Count} Azure voices:");
                Console.WriteLine("ID | Name | Locale | Gender | Styles | Roles");
                Console.WriteLine("---|------|--------|--------|--------|------");
                
                foreach (var voice in voicesList)
                {
                    string styles = voice.StyleList.Count > 0 ? string.Join(", ", voice.StyleList.ToArray()) : "None";
                    string roles = voice.RoleList.Count > 0 ? string.Join(", ", voice.RoleList.ToArray()) : "None";
                    
                    Console.WriteLine($"{voice.ShortName} | {voice.DisplayName} | {voice.Locale} | {voice.Gender} | {styles} | {roles}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error listing Azure voices: {ex.Message}");
            }
        }

        private static async Task InstallSherpaOnnxVoiceInteractive(ModelInstaller installer, Sapi5RegistrarExtended registrar, string dllPath)
        {
            Console.WriteLine("Loading available Sherpa ONNX voices...");
            var models = await LoadModelsAsync();

            Console.WriteLine($"Found {models.Count} voices.");
            Console.WriteLine("You can search for voices by:");
            Console.WriteLine(" - Language (e.g., 'english', 'spanish')");
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
                ).ToList(); // Convert to list before checking Count

                if (filteredModels.Count == 0)
                {
                    Console.WriteLine("No voices matched your search. Try again.");
                    continue;
                }

                // Display filtered results
                Console.WriteLine("Matching voices:");
                foreach (var model in filteredModels)
                {
                    var languages = model.Language.Select(l => l.LanguageName).ToList();
                    string languageStr = string.Join(", ", languages);
                    Console.WriteLine($" - {model.Id} ({model.Name}, Type: {model.ModelType}, Language: {languageStr})");
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
                        await installer.DownloadAndExtractModelAsync(chosenModel);
                        registrar.RegisterSherpaVoice(chosenModel, dllPath);
                        Console.WriteLine($"Voice {chosenModel.Name} installed successfully!");
                        return;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error installing voice: {ex.Message}");
                    }
                }
                else
                {
                    Console.WriteLine("Invalid model ID. Please try again.");
                }
            }
        }

        private static async Task InstallAzureVoiceInteractive(Sapi5RegistrarExtended registrar, string dllPath)
        {
            Console.WriteLine("Azure TTS Voice Installation");
            Console.WriteLine("===========================");
            
            // Try to load config first
            var config = AzureConfigManager.LoadConfig();
            string defaultKey = config.DefaultKey;
            string defaultRegion = config.DefaultRegion;
            
            // Get Azure subscription key and region
            if (!string.IsNullOrEmpty(defaultKey))
            {
                Console.Write($"Enter your Azure subscription key (or press Enter to use saved key): ");
            }
            else
            {
                Console.Write("Enter your Azure subscription key: ");
            }
            
            string subscriptionKey = Console.ReadLine();
            
            if (string.IsNullOrEmpty(subscriptionKey) && !string.IsNullOrEmpty(defaultKey))
            {
                subscriptionKey = defaultKey;
                Console.WriteLine("Using saved subscription key.");
            }
            
            if (string.IsNullOrEmpty(subscriptionKey))
            {
                Console.WriteLine("Subscription key is required. Aborting installation.");
                return;
            }
            
            if (!string.IsNullOrEmpty(defaultRegion))
            {
                Console.Write($"Enter your Azure region (e.g., eastus, westus) (or press Enter to use saved region '{defaultRegion}'): ");
            }
            else
            {
                Console.Write("Enter your Azure region (e.g., eastus, westus): ");
            }
            
            string region = Console.ReadLine();
            
            if (string.IsNullOrEmpty(region) && !string.IsNullOrEmpty(defaultRegion))
            {
                region = defaultRegion;
                Console.WriteLine($"Using saved region: {region}");
            }
            
            if (string.IsNullOrEmpty(region))
            {
                Console.WriteLine("Region is required. Aborting installation.");
                return;
            }
            
            // Ask if user wants to save this configuration
            if (subscriptionKey != defaultKey || region != defaultRegion)
            {
                Console.Write("Do you want to save this configuration for future use? (y/n): ");
                string saveResponse = Console.ReadLine();
                
                if (saveResponse?.ToLower() == "y" || saveResponse?.ToLower() == "yes")
                {
                    Console.Write("Do you want to encrypt the subscription key? (y/n, default: y): ");
                    string encryptResponse = Console.ReadLine();
                    bool encrypt = encryptResponse?.ToLower() != "n" && encryptResponse?.ToLower() != "no";
                    
                    AzureConfigManager.SaveConfig(subscriptionKey, region, encrypt);
                }
            }
            
            try
            {
                // Create Azure TTS service
                var azureService = new AzureTtsService(subscriptionKey, region);
                
                // Validate subscription
                Console.WriteLine("Validating Azure subscription...");
                bool isValid = await azureService.ValidateSubscriptionAsync();
                if (!isValid)
                {
                    Console.WriteLine("Error: Invalid Azure subscription key or region.");
                    return;
                }
                
                // Get available voices
                Console.WriteLine("Fetching available Azure voices...");
                var voices = await azureService.GetAvailableVoicesAsync();
                var voicesList = voices.ToList(); // Convert to list before accessing Count
                
                Console.WriteLine($"Found {voicesList.Count} Azure voices.");
                Console.WriteLine("You can search for voices by:");
                Console.WriteLine(" - Language (e.g., 'English', 'Spanish')");
                Console.WriteLine(" - Voice name (e.g., 'Guy', 'Aria')");
                Console.WriteLine(" - Gender (e.g., 'Male', 'Female')");
                Console.WriteLine();
                
                while (true)
                {
                    Console.Write("Enter a search term (or 'exit' to quit): ");
                    string searchTerm = Console.ReadLine();
                    
                    if (string.IsNullOrWhiteSpace(searchTerm))
                        continue;
                    
                    if (searchTerm.Equals("exit", StringComparison.OrdinalIgnoreCase))
                        return;
                    
                    // Filter voices
                    var filteredVoices = voices.Where(voice =>
                        voice.ShortName.Contains(searchTerm, StringComparison.OrdinalIgnoreCase) ||
                        voice.DisplayName.Contains(searchTerm, StringComparison.OrdinalIgnoreCase) ||
                        voice.Locale.Contains(searchTerm, StringComparison.OrdinalIgnoreCase) ||
                        voice.Gender.Contains(searchTerm, StringComparison.OrdinalIgnoreCase) ||
                        // Use the language code converter for intelligent language matching
                        LanguageCodeConverter.LocaleMatchesLanguage(voice.Locale, searchTerm)
                    ).ToList(); // Convert to list before checking Count
                    
                    if (filteredVoices.Count == 0)
                    {
                        Console.WriteLine("No voices matched your search. Try again.");
                        continue;
                    }
                    
                    // Display filtered results
                    Console.WriteLine("Matching voices:");
                    for (int i = 0; i < filteredVoices.Count; i++)
                    {
                        var voice = filteredVoices[i];
                        string styles = voice.StyleList.Count > 0 ? $", Styles: {string.Join(", ", voice.StyleList.ToArray())}" : "";
                        string roles = voice.RoleList.Count > 0 ? $", Roles: {string.Join(", ", voice.RoleList.ToArray())}" : "";
                        
                        Console.WriteLine($"{i + 1}. {voice.ShortName} ({voice.DisplayName}, {voice.Locale}, {voice.Gender}{styles}{roles})");
                    }
                    
                    // Prompt user to select a voice
                    Console.Write("Enter the number of the voice to install (or 'search' to search again): ");
                    string selection = Console.ReadLine();
                    
                    if (selection.Equals("search", StringComparison.OrdinalIgnoreCase))
                        continue;
                    
                    if (int.TryParse(selection, out int selectedIndex) && selectedIndex >= 1 && selectedIndex <= filteredVoices.Count)
                    {
                        var selectedVoice = filteredVoices[selectedIndex - 1];
                        
                        // Handle style selection if available
                        if (selectedVoice.StyleList.Count > 0)
                        {
                            Console.WriteLine($"Available styles for {selectedVoice.ShortName}: {string.Join(", ", selectedVoice.StyleList.ToArray())}");
                            Console.Write("Enter a style to use (or press Enter for none): ");
                            string style = Console.ReadLine();
                            
                            if (!string.IsNullOrWhiteSpace(style) && selectedVoice.StyleList.Contains(style))
                            {
                                selectedVoice.SelectedStyle = style;
                                Console.WriteLine($"Using style: {style}");
                            }
                            else if (!string.IsNullOrWhiteSpace(style))
                            {
                                Console.WriteLine($"Style '{style}' not available. No style will be used.");
                            }
                        }
                        
                        // Handle role selection if available
                        if (selectedVoice.RoleList.Count > 0)
                        {
                            Console.WriteLine($"Available roles for {selectedVoice.ShortName}: {string.Join(", ", selectedVoice.RoleList.ToArray())}");
                            Console.Write("Enter a role to use (or press Enter for none): ");
                            string role = Console.ReadLine();
                            
                            if (!string.IsNullOrWhiteSpace(role) && selectedVoice.RoleList.Contains(role))
                            {
                                selectedVoice.SelectedRole = role;
                                Console.WriteLine($"Using role: {role}");
                            }
                            else if (!string.IsNullOrWhiteSpace(role))
                            {
                                Console.WriteLine($"Role '{role}' not available. No role will be used.");
                            }
                        }
                        
                        // Register the voice
                        Console.WriteLine($"Registering Azure voice: {selectedVoice.ShortName}");
                        registrar.RegisterAzureVoice(selectedVoice, dllPath);
                        
                        Console.WriteLine($"Azure voice '{selectedVoice.ShortName}' installed successfully!");
                        return;
                    }
                    else
                    {
                        Console.WriteLine("Invalid selection. Please try again.");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error installing Azure voice: {ex.Message}");
            }
        }
    }
}
