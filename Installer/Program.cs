using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Principal;
using System.Threading.Tasks;
using Microsoft.Win32;
using Newtonsoft.Json;
using System.Text.Json;
using System.Text.Json.Nodes;
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
            string managedDllPath = @"C:\Program Files\OpenAssistive\OpenSpeech\OpenSpeechTTS.dll";
            string nativeDllPath = @"C:\Program Files\OpenAssistive\OpenSpeech\NativeTTSWrapper.dll";

            // Check for command line arguments
            if (args.Length >= 2)
            {
                string command = args[0].ToLower();
                string modelId = args[1];

                switch (command)
                {
                    case "install":
                        await InstallSpecificVoice(modelId, installer, registrar, nativeDllPath);
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

                        await InstallAzureVoice(voiceName, subscriptionKey, region, style, role, registrar, managedDllPath);
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
                            await UninstallVoicesAndDll(registrar, managedDllPath, nativeDllPath);
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
                    await InstallSherpaOnnxVoiceInteractive(installer, registrar, nativeDllPath);
                    break;

                case "2":
                    await InstallAzureVoiceInteractive(registrar, managedDllPath);
                    break;

                case "3":
                    await UninstallVoicesAndDll(registrar, managedDllPath, nativeDllPath);
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

        private static async Task UninstallVoicesAndDll(Sapi5RegistrarExtended registrar, string managedDllPath, string nativeDllPath)
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
                                // Check if this is an Azure voice by looking at the VoiceType attribute
                                using (var voiceKey = voicesKey.OpenSubKey(voiceName))
                                {
                                    if (voiceKey != null)
                                    {
                                        // Check CLSID first to ensure it's our voice
                                        string clsid = voiceKey.GetValue("CLSID") as string;
                                        if (clsid == "{A1B2C3D4-E5F6-4A5B-8C9D-1E2F3A4B5C6D}") // Our TTS Engine CLSID
                                        {
                                            // Check VoiceType attribute to distinguish Azure from Sherpa
                                            using (var attributesKey = voiceKey.OpenSubKey("Attributes"))
                                            {
                                                if (attributesKey != null)
                                                {
                                                    string voiceType = attributesKey.GetValue("VoiceType") as string;
                                                    if (voiceType == "AzureTTS")
                                                    {
                                                        registrar.UnregisterVoice(voiceName);
                                                        Console.WriteLine($"Unregistered Azure voice: {voiceName}");

                                                        // Also remove from engines_config.json
                                                        // Extract voice name from SAPI voice name if possible
                                                        if (voiceName.Contains("(") && voiceName.Contains(")"))
                                                        {
                                                            var parts = voiceName.Split('(', ')');
                                                            if (parts.Length >= 2)
                                                            {
                                                                var voiceParts = parts[1].Split(',');
                                                                if (voiceParts.Length >= 2)
                                                                {
                                                                    string azureVoiceName = voiceParts[1].Trim();
                                                                    EngineConfigManager.RemoveAzureVoice(azureVoiceName);
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
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
                UnregisterDll(managedDllPath);
                UnregisterNativeDll(nativeDllPath);

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

        private static void UnregisterNativeDll(string dllPath)
        {
            try
            {
                if (!File.Exists(dllPath))
                {
                    Console.WriteLine($"Native DLL not found: {dllPath}");
                    return;
                }

                // Use regsvr32 to unregister the native COM DLL
                var process = new Process();
                process.StartInfo.FileName = "regsvr32.exe";
                process.StartInfo.Arguments = $"/s /u \"{dllPath}\"";
                process.StartInfo.RedirectStandardOutput = true;
                process.StartInfo.UseShellExecute = false;
                process.Start();
                process.WaitForExit();

                if (process.ExitCode == 0)
                {
                    Console.WriteLine("Native DLL successfully unregistered.");
                }
                else
                {
                    Console.WriteLine($"Warning: Native DLL unregistration returned exit code: {process.ExitCode}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error unregistering native DLL: {ex.Message}");
            }
        }

        private static async Task InstallSpecificVoice(string modelId, ModelInstaller installer, Sapi5RegistrarExtended registrar, string dllPath)
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

                    // Copy and register the native COM DLL first (for SherpaOnnx voices)
                    CopyAndRegisterNativeDll(dllPath);

                    // Then register the voice
                    registrar.RegisterVoice(model, dllPath);

                    // THE BUG FIX: Update engines_config.json for SherpaOnnx voices
                    EngineConfigManager.AddSherpaVoice(model);

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

        private static void CopyAndRegisterComDll(string targetDllPath)
        {
            try
            {
                Console.WriteLine($"Copying and registering COM DLL to: {targetDllPath}");

                // Find the source DLL in the build output
                string sourceDllPath = FindSourceDllPath();
                if (string.IsNullOrEmpty(sourceDllPath) || !File.Exists(sourceDllPath))
                {
                    throw new FileNotFoundException($"Source DLL not found. Expected at: {sourceDllPath}");
                }

                Console.WriteLine($"Source DLL found at: {sourceDllPath}");

                // Ensure the target directory exists
                string targetDir = Path.GetDirectoryName(targetDllPath);
                if (!Directory.Exists(targetDir))
                {
                    Directory.CreateDirectory(targetDir);
                    Console.WriteLine($"Created directory: {targetDir}");
                }

                // Copy the DLL and its dependencies
                CopyDllAndDependencies(sourceDllPath, targetDllPath);

                // Register the DLL
                RegisterComDll(targetDllPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error copying and registering COM DLL: {ex.Message}");
                throw;
            }
        }

        private static void CopyAndRegisterNativeDll(string targetDllPath)
        {
            try
            {
                Console.WriteLine($"Copying and registering native COM DLL to: {targetDllPath}");

                // Find the source native DLL in the build output
                string sourceDllPath = FindSourceNativeDllPath();
                if (string.IsNullOrEmpty(sourceDllPath) || !File.Exists(sourceDllPath))
                {
                    throw new FileNotFoundException($"Source native DLL not found. Expected at: {sourceDllPath}");
                }

                Console.WriteLine($"Source native DLL found at: {sourceDllPath}");

                // Ensure the target directory exists
                string targetDir = Path.GetDirectoryName(targetDllPath);
                if (!Directory.Exists(targetDir))
                {
                    Directory.CreateDirectory(targetDir);
                    Console.WriteLine($"Created directory: {targetDir}");
                }

                // Copy the native DLL
                File.Copy(sourceDllPath, targetDllPath, true);
                Console.WriteLine($"Copied native DLL: {sourceDllPath} -> {targetDllPath}");

                // Also copy the SherpaWorker.exe and dependencies
                CopyProcessBridgeComponents(targetDir);

                // Register the native DLL
                RegisterNativeDll(targetDllPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error copying and registering native COM DLL: {ex.Message}");
                throw;
            }
        }

        private static string FindSourceNativeDllPath()
        {
            // Get the directory where the installer is running from
            string installerDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);

            // Look for NativeTTSWrapper.dll in the same directory
            string dllPath = Path.Combine(installerDir, "NativeTTSWrapper.dll");
            if (File.Exists(dllPath))
            {
                return dllPath;
            }

            // If not found, try relative paths from the installer location
            string[] possiblePaths = {
                Path.Combine(installerDir, "..", "..", "..", "NativeTTSWrapper", "x64", "Release", "NativeTTSWrapper.dll"),
                Path.Combine(installerDir, "..", "..", "..", "NativeTTSWrapper", "x64", "Debug", "NativeTTSWrapper.dll"),
                Path.Combine(installerDir, "NativeTTSWrapper.dll")
            };

            foreach (string path in possiblePaths)
            {
                string fullPath = Path.GetFullPath(path);
                if (File.Exists(fullPath))
                {
                    return fullPath;
                }
            }

            return null;
        }

        private static void CopyProcessBridgeComponents(string targetDir)
        {
            try
            {
                // Get the directory where the installer is running from
                string installerDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);

                // Copy SherpaWorker.exe
                string[] workerPaths = {
                    Path.Combine(installerDir, "SherpaWorker.exe"),
                    Path.Combine(installerDir, "..", "..", "..", "SherpaWorker", "bin", "Release", "net6.0", "SherpaWorker.exe"),
                    Path.Combine(installerDir, "..", "..", "..", "SherpaWorker", "bin", "Debug", "net6.0", "SherpaWorker.exe")
                };

                string sourceSherpaWorker = null;
                foreach (string path in workerPaths)
                {
                    string fullPath = Path.GetFullPath(path);
                    if (File.Exists(fullPath))
                    {
                        sourceSherpaWorker = fullPath;
                        break;
                    }
                }

                if (sourceSherpaWorker != null)
                {
                    string targetSherpaWorker = Path.Combine(targetDir, "SherpaWorker.exe");
                    File.Copy(sourceSherpaWorker, targetSherpaWorker, true);
                    Console.WriteLine($"Copied SherpaWorker.exe");

                    // Copy SherpaWorker dependencies
                    string sourceWorkerDir = Path.GetDirectoryName(sourceSherpaWorker);
                    string[] workerDependencies = {
                        "sherpa-onnx.dll",
                        "SherpaNative.dll",
                        "onnxruntime.dll",
                        "onnxruntime_providers_shared.dll",
                        "SherpaWorker.deps.json",
                        "SherpaWorker.runtimeconfig.json"
                    };

                    foreach (string dep in workerDependencies)
                    {
                        string sourceDep = Path.Combine(sourceWorkerDir, dep);
                        string targetDep = Path.Combine(targetDir, dep);

                        if (File.Exists(sourceDep))
                        {
                            File.Copy(sourceDep, targetDep, true);
                            Console.WriteLine($"Copied SherpaWorker dependency: {dep}");
                        }
                    }
                }
                else
                {
                    Console.WriteLine("Warning: SherpaWorker.exe not found");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error copying ProcessBridge components: {ex.Message}");
            }
        }

        private static void RegisterNativeDll(string dllPath)
        {
            try
            {
                Console.WriteLine($"Registering native COM DLL: {dllPath}");

                // Check if the DLL exists
                if (!File.Exists(dllPath))
                {
                    throw new FileNotFoundException($"Native DLL not found: {dllPath}");
                }

                // Register the DLL with regsvr32
                var process = new Process();
                process.StartInfo.FileName = "regsvr32.exe";
                process.StartInfo.Arguments = $"/s \"{dllPath}\"";
                process.StartInfo.RedirectStandardOutput = true;
                process.StartInfo.UseShellExecute = false;
                process.Start();
                process.WaitForExit();

                if (process.ExitCode != 0)
                {
                    throw new Exception($"Failed to register native COM DLL. Exit code: {process.ExitCode}");
                }

                Console.WriteLine("Native COM DLL registered successfully");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error registering native COM DLL: {ex.Message}");
                throw;
            }
        }

        private static string FindSourceDllPath()
        {
            // Get the directory where the installer is running from
            string installerDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);

            // Look for OpenSpeechTTS.dll in the same directory
            string dllPath = Path.Combine(installerDir, "OpenSpeechTTS.dll");
            if (File.Exists(dllPath))
            {
                return dllPath;
            }

            // If not found, try relative paths from the installer location
            string[] possiblePaths = {
                Path.Combine(installerDir, "..", "..", "..", "OpenSpeechTTS", "bin", "Debug", "net472", "OpenSpeechTTS.dll"),
                Path.Combine(installerDir, "..", "..", "..", "OpenSpeechTTS", "bin", "Release", "net472", "OpenSpeechTTS.dll"),
                Path.Combine(installerDir, "OpenSpeechTTS.dll")
            };

            foreach (string path in possiblePaths)
            {
                string fullPath = Path.GetFullPath(path);
                if (File.Exists(fullPath))
                {
                    return fullPath;
                }
            }

            return null;
        }

        private static void CopyDllAndDependencies(string sourceDllPath, string targetDllPath)
        {
            try
            {
                // Copy the main DLL
                File.Copy(sourceDllPath, targetDllPath, true);
                Console.WriteLine($"Copied DLL: {sourceDllPath} -> {targetDllPath}");

                // Copy dependencies
                string sourceDir = Path.GetDirectoryName(sourceDllPath);
                string targetDir = Path.GetDirectoryName(targetDllPath);

                // List of dependencies to copy
                string[] dependencies = {
                    "sherpa-onnx.dll",
                    "SherpaNative.dll",
                    "onnxruntime.dll",
                    "onnxruntime_providers_shared.dll"
                };

                foreach (string dep in dependencies)
                {
                    string sourceDep = Path.Combine(sourceDir, dep);
                    string targetDep = Path.Combine(targetDir, dep);

                    if (File.Exists(sourceDep))
                    {
                        File.Copy(sourceDep, targetDep, true);
                        Console.WriteLine($"Copied dependency: {dep}");
                    }
                    else
                    {
                        Console.WriteLine($"Warning: Dependency not found: {sourceDep}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error copying DLL and dependencies: {ex.Message}");
                throw;
            }
        }

        private static void RegisterComDll(string dllPath)
        {
            try
            {
                Console.WriteLine($"Registering COM DLL: {dllPath}");

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

        private static async Task UninstallSpecificVoice(string modelId, Sapi5RegistrarExtended registrar)
        {
            try
            {
                var models = await LoadModelsAsync();
                if (models.TryGetValue(modelId, out var model))
                {
                    registrar.UnregisterVoice(modelId);

                    // Also remove from engines_config.json
                    EngineConfigManager.RemoveSherpaVoice(modelId);

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

                // Update engines_config.json with the new voice configuration
                string sapiVoiceName = $"Microsoft Server Speech Text to Speech Voice ({voice.Locale}, {voice.ShortName})";
                EngineConfigManager.AddAzureVoice(voice.ShortName, subscriptionKey, region, voice.Locale, sapiVoiceName);

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
            Console.WriteLine("🎵 EASY VOICE SELECTION:");
            Console.WriteLine(" - Enter a NUMBER to select from the list");
            Console.WriteLine(" - Enter a NAME (e.g., 'amy', 'danny', 'english')");
            Console.WriteLine(" - Enter a full MODEL ID if you know it");
            Console.WriteLine(" - Popular searches: 'amy', 'english', 'piper', 'low quality'");
            Console.WriteLine();

            while (true)
            {
                Console.Write("🔍 Enter search term or 'popular' for common voices (or 'exit' to quit): ");
                string searchTerm = Console.ReadLine();

                if (string.IsNullOrWhiteSpace(searchTerm))
                    continue;

                if (searchTerm.Equals("exit", StringComparison.OrdinalIgnoreCase))
                    return;

                List<TtsModel> filteredModels;

                // Show popular English voices if requested
                if (searchTerm.Equals("popular", StringComparison.OrdinalIgnoreCase))
                {
                    filteredModels = models.Values.Where(model =>
                        model.Id?.StartsWith("piper-en-", StringComparison.OrdinalIgnoreCase) == true &&
                        (model.Quality == "low" || model.Quality == "medium") &&
                        model.FilesizeMb < 200 // Reasonable size
                    ).OrderBy(m => m.Quality).ThenBy(m => m.Name).ToList();

                    Console.WriteLine("🌟 Popular English Voices (fast download, good quality):");
                }
                else
                {
                    // Enhanced search - much more flexible
                    filteredModels = models.Values.Where(model =>
                        // Direct ID match
                        (model.Id?.Contains(searchTerm, StringComparison.OrdinalIgnoreCase) ?? false) ||
                        // Name match (most common)
                        (model.Name?.Contains(searchTerm, StringComparison.OrdinalIgnoreCase) ?? false) ||
                        // Quality match (low, medium, high)
                        (model.Quality?.Contains(searchTerm, StringComparison.OrdinalIgnoreCase) ?? false) ||
                        // Developer/type match
                        (model.Developer?.Contains(searchTerm, StringComparison.OrdinalIgnoreCase) ?? false) ||
                        (model.ModelType?.Contains(searchTerm, StringComparison.OrdinalIgnoreCase) ?? false) ||
                        // Language matching
                        (model.Language != null && model.Language.Any(lang =>
                            (lang.LangCode?.Contains(searchTerm, StringComparison.OrdinalIgnoreCase) ?? false) ||
                            (lang.LanguageName?.Contains(searchTerm, StringComparison.OrdinalIgnoreCase) ?? false))) ||
                        // Smart partial matching for common terms
                        (searchTerm.Equals("english", StringComparison.OrdinalIgnoreCase) &&
                         model.Id?.Contains("en-", StringComparison.OrdinalIgnoreCase) == true) ||
                        (searchTerm.Equals("amy", StringComparison.OrdinalIgnoreCase) &&
                         model.Id?.Contains("amy", StringComparison.OrdinalIgnoreCase) == true)
                    ).ToList();
                }

                if (filteredModels.Count == 0)
                {
                    Console.WriteLine("❌ No voices matched your search. Try:");
                    Console.WriteLine("   • 'amy' - for Amy voices");
                    Console.WriteLine("   • 'english' - for English voices");
                    Console.WriteLine("   • 'popular' - for recommended voices");
                    Console.WriteLine("   • 'piper' - for Piper voices");
                    continue;
                }

                // Display filtered results with NUMBERS for easy selection
                Console.WriteLine($"\n✅ Found {filteredModels.Count} matching voices:");
                Console.WriteLine("".PadRight(80, '='));

                for (int i = 0; i < Math.Min(filteredModels.Count, 20); i++) // Limit to 20 for readability
                {
                    var model = filteredModels[i];
                    var languages = model.Language?.Select(l => l.LanguageName).ToList() ?? new List<string>();
                    string languageStr = languages.Count > 0 ? string.Join(", ", languages) : "Unknown";
                    string sizeStr = model.FilesizeMb > 0 ? $"{model.FilesizeMb:F1}MB" : "Unknown size";

                    Console.WriteLine($"{i + 1,2}. {model.Name} ({model.Quality}) - {languageStr} [{sizeStr}]");
                    Console.WriteLine($"    ID: {model.Id}");
                    if (i < filteredModels.Count - 1) Console.WriteLine();
                }

                if (filteredModels.Count > 20)
                {
                    Console.WriteLine($"... and {filteredModels.Count - 20} more. Refine your search to see all results.");
                }

                Console.WriteLine("".PadRight(80, '='));
                Console.WriteLine();

                // Enhanced selection prompt
                Console.Write("👉 Select voice by NUMBER (1-" + Math.Min(filteredModels.Count, 20) + "), NAME, or full ID (or 'search' for new search): ");
                string selection = Console.ReadLine();

                if (selection.Equals("search", StringComparison.OrdinalIgnoreCase))
                    continue;

                TtsModel chosenModel = null;

                // Try to parse as number first (most user-friendly)
                if (int.TryParse(selection, out int selectedIndex) &&
                    selectedIndex >= 1 && selectedIndex <= Math.Min(filteredModels.Count, 20))
                {
                    chosenModel = filteredModels[selectedIndex - 1];
                    Console.WriteLine($"✅ Selected: {chosenModel.Name} ({chosenModel.Quality}) - {chosenModel.Id}");
                }
                // Try exact ID match
                else if (models.TryGetValue(selection, out chosenModel))
                {
                    Console.WriteLine($"✅ Found by ID: {chosenModel.Name} - {chosenModel.Id}");
                }
                // Try partial name matching from current filtered results
                else
                {
                    var nameMatches = filteredModels.Where(m =>
                        m.Name?.Contains(selection, StringComparison.OrdinalIgnoreCase) == true ||
                        m.Id?.Contains(selection, StringComparison.OrdinalIgnoreCase) == true
                    ).ToList();

                    if (nameMatches.Count == 1)
                    {
                        chosenModel = nameMatches[0];
                        Console.WriteLine($"✅ Found by name: {chosenModel.Name} - {chosenModel.Id}");
                    }
                    else if (nameMatches.Count > 1)
                    {
                        Console.WriteLine($"❌ Multiple matches for '{selection}'. Please be more specific or use a number.");
                        continue;
                    }
                }

                if (chosenModel != null)
                {
                    try
                    {
                        Console.WriteLine($"\n🚀 Installing {chosenModel.Name} ({chosenModel.Quality})...");
                        Console.WriteLine($"   Size: {chosenModel.FilesizeMb:F1}MB");
                        Console.WriteLine($"   ID: {chosenModel.Id}");
                        Console.WriteLine();

                        await installer.DownloadAndExtractModelAsync(chosenModel);
                        registrar.RegisterSherpaVoice(chosenModel, dllPath);

                        // THE BUG FIX: Update engines_config.json for SherpaOnnx voices
                        EngineConfigManager.AddSherpaVoice(chosenModel);

                        Console.WriteLine($"🎉 Voice '{chosenModel.Name}' installed successfully!");
                        Console.WriteLine($"💡 You can now use this voice in any Windows application that supports SAPI!");
                        return;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"❌ Error installing voice: {ex.Message}");
                        Console.WriteLine("Please try a different voice or check your internet connection.");
                    }
                }
                else
                {
                    Console.WriteLine($"❌ Could not find voice '{selection}'. Please try:");
                    Console.WriteLine("   • A number from the list above");
                    Console.WriteLine("   • A voice name like 'amy' or 'danny'");
                    Console.WriteLine("   • The full voice ID");
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

                        // Update engines_config.json with the new voice configuration
                        string sapiVoiceName = $"Microsoft Server Speech Text to Speech Voice ({selectedVoice.Locale}, {selectedVoice.ShortName})";
                        EngineConfigManager.AddAzureVoice(selectedVoice.ShortName, subscriptionKey, region, selectedVoice.Locale, sapiVoiceName);

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

    /// <summary>
    /// Manages the engines_config.json file for TTS engine configuration
    /// This class handles adding/removing Azure and SherpaOnnx voice configurations
    /// </summary>
    public static class EngineConfigManager
    {
        private static readonly string ConfigPath = @"C:\Program Files\OpenAssistive\OpenSpeech\engines_config.json";

        /// <summary>
        /// Adds an Azure TTS voice configuration to engines_config.json
        /// </summary>
        /// <param name="voiceName">Azure voice name (e.g., "en-GB-ElliotNeural")</param>
        /// <param name="subscriptionKey">Azure subscription key</param>
        /// <param name="region">Azure region</param>
        /// <param name="language">Language code (e.g., "en-GB")</param>
        /// <param name="sapiVoiceName">Full SAPI voice name for mapping</param>
        public static void AddAzureVoice(string voiceName, string subscriptionKey, string region,
            string language, string sapiVoiceName)
        {
            try
            {
                Console.WriteLine($"Updating engines configuration for: {voiceName}");

                // Load existing configuration
                var config = LoadConfiguration();

                // Create engine ID (e.g., "azure-elliot")
                string engineId = GenerateAzureEngineId(voiceName);

                // Create engine configuration
                var engineConfig = new JsonObject
                {
                    ["type"] = "azure",
                    ["config"] = new JsonObject
                    {
                        ["subscriptionKey"] = subscriptionKey,
                        ["region"] = region,
                        ["voiceName"] = voiceName,
                        ["language"] = language,
                        ["sampleRate"] = 24000,
                        ["channels"] = 1,
                        ["bitsPerSample"] = 16
                    }
                };

                // Add to engines section
                var engines = config["engines"]?.AsObject();
                if (engines != null)
                {
                    engines[engineId] = engineConfig;
                    Console.WriteLine($"Added engine configuration: {engineId}");
                }

                // Add to voices section
                var voices = config["voices"]?.AsObject();
                if (voices != null)
                {
                    // Add short name mapping (e.g., "elliot" -> "azure-elliot")
                    string shortName = GenerateShortName(voiceName);
                    voices[shortName] = engineId;
                    Console.WriteLine($"Added voice mapping: {shortName} -> {engineId}");

                    // Add SAPI voice name mapping
                    if (!string.IsNullOrEmpty(sapiVoiceName))
                    {
                        voices[sapiVoiceName] = engineId;
                        Console.WriteLine($"Added SAPI mapping for: {sapiVoiceName}");
                    }
                }

                // Save configuration
                SaveConfiguration(config);
                Console.WriteLine($"Engine configuration updated successfully for '{voiceName}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Could not update engine configuration: {ex.Message}");
                // Don't fail the installation if config update fails
            }
        }

        /// <summary>
        /// Removes an Azure TTS voice configuration from engines_config.json
        /// </summary>
        /// <param name="voiceName">Azure voice name to remove</param>
        public static void RemoveAzureVoice(string voiceName)
        {
            try
            {
                Console.WriteLine($"Removing engine configuration for: {voiceName}");

                var config = LoadConfiguration();
                string engineId = GenerateAzureEngineId(voiceName);

                // Remove from engines
                var engines = config["engines"]?.AsObject();
                if (engines != null && engines.ContainsKey(engineId))
                {
                    engines.Remove(engineId);
                    Console.WriteLine($"Removed engine: {engineId}");
                }

                // Remove from voices (find all mappings to this engine)
                var voices = config["voices"]?.AsObject();
                if (voices != null)
                {
                    var toRemove = new List<string>();
                    foreach (var kvp in voices)
                    {
                        if (kvp.Value?.ToString() == engineId)
                        {
                            toRemove.Add(kvp.Key);
                        }
                    }

                    foreach (var key in toRemove)
                    {
                        voices.Remove(key);
                        Console.WriteLine($"Removed voice mapping: {key}");
                    }
                }

                SaveConfiguration(config);
                Console.WriteLine($"Engine configuration removed for '{voiceName}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Could not remove engine configuration: {ex.Message}");
            }
        }

        /// <summary>
        /// Adds a SherpaOnnx voice configuration to engines_config.json
        /// THIS IS THE BUG FIX - SherpaOnnx voices were not being added to engine config!
        /// </summary>
        /// <param name="model">The TTS model to add</param>
        public static void AddSherpaVoice(TtsModel model)
        {
            try
            {
                Console.WriteLine($"🔧 Updating engines configuration for SherpaOnnx voice: {model.Name}");

                // Load existing configuration or create new one
                JsonNode config;
                try
                {
                    config = LoadConfiguration();
                }
                catch (FileNotFoundException)
                {
                    // Create new configuration if file doesn't exist
                    Console.WriteLine("Creating new engines_config.json file...");
                    config = new JsonObject
                    {
                        ["engines"] = new JsonObject(),
                        ["voices"] = new JsonObject()
                    };

                    // Ensure directory exists
                    string configDir = Path.GetDirectoryName(ConfigPath);
                    if (!Directory.Exists(configDir))
                    {
                        Directory.CreateDirectory(configDir);
                    }
                }

                // Create engine ID (e.g., "sherpa-amy-low")
                string engineId = GenerateSherpaEngineId(model.Id);

                // Create engine configuration
                var engineConfig = new JsonObject
                {
                    ["type"] = "sherpa",
                    ["config"] = new JsonObject
                    {
                        ["modelPath"] = $"models/{model.Id}",
                        ["sampleRate"] = model.SampleRate > 0 ? model.SampleRate : 22050,
                        ["channels"] = 1,
                        ["bitsPerSample"] = 16,
                        ["voiceId"] = model.Id,
                        ["modelType"] = model.ModelType ?? "vits"
                    }
                };

                // Add to engines section
                var engines = config["engines"]?.AsObject();
                if (engines != null)
                {
                    engines[engineId] = engineConfig;
                    Console.WriteLine($"✅ Added engine configuration: {engineId}");
                }

                // Add to voices section with multiple mappings for easy access
                var voices = config["voices"]?.AsObject();
                if (voices != null)
                {
                    // Add short name mapping (e.g., "amy" -> "sherpa-amy-low")
                    string shortName = GenerateSherpaShortName(model.Id, model.Name);
                    voices[shortName] = engineId;
                    Console.WriteLine($"✅ Added voice mapping: {shortName} -> {engineId}");

                    // Add full ID mapping
                    voices[model.Id] = engineId;
                    Console.WriteLine($"✅ Added full ID mapping: {model.Id} -> {engineId}");

                    // Add model name mapping if different from short name
                    if (!string.IsNullOrEmpty(model.Name) && model.Name.ToLower() != shortName)
                    {
                        voices[model.Name.ToLower()] = engineId;
                        Console.WriteLine($"✅ Added name mapping: {model.Name.ToLower()} -> {engineId}");
                    }
                }

                // Save configuration
                SaveConfiguration(config);
                Console.WriteLine($"🎉 Engine configuration updated successfully for SherpaOnnx voice '{model.Name}'!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"⚠️ Warning: Could not update engine configuration: {ex.Message}");
                Console.WriteLine("   Voice will still be registered in SAPI, but may not work for synthesis");
                // Don't fail the installation if config update fails
            }
        }

        /// <summary>
        /// Removes a SherpaOnnx voice configuration from engines_config.json
        /// </summary>
        /// <param name="voiceId">SherpaOnnx voice ID to remove</param>
        public static void RemoveSherpaVoice(string voiceId)
        {
            try
            {
                Console.WriteLine($"Removing SherpaOnnx engine configuration for: {voiceId}");

                JsonNode config;
                try
                {
                    config = LoadConfiguration();
                }
                catch (FileNotFoundException)
                {
                    Console.WriteLine("No engine configuration file found - nothing to remove.");
                    return;
                }

                string engineId = GenerateSherpaEngineId(voiceId);

                // Remove from engines
                var engines = config["engines"]?.AsObject();
                if (engines != null && engines.ContainsKey(engineId))
                {
                    engines.Remove(engineId);
                    Console.WriteLine($"Removed engine: {engineId}");
                }

                // Remove from voices (find all mappings to this engine)
                var voices = config["voices"]?.AsObject();
                if (voices != null)
                {
                    var toRemove = new List<string>();
                    foreach (var kvp in voices)
                    {
                        if (kvp.Value?.ToString() == engineId)
                        {
                            toRemove.Add(kvp.Key);
                        }
                    }

                    foreach (var key in toRemove)
                    {
                        voices.Remove(key);
                        Console.WriteLine($"Removed voice mapping: {key}");
                    }
                }

                SaveConfiguration(config);
                Console.WriteLine($"Engine configuration removed for '{voiceId}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Could not remove SherpaOnnx engine configuration: {ex.Message}");
            }
        }

        private static JsonNode LoadConfiguration()
        {
            if (!File.Exists(ConfigPath))
            {
                throw new FileNotFoundException($"Configuration file not found: {ConfigPath}");
            }

            string json = File.ReadAllText(ConfigPath);
            return JsonNode.Parse(json) ?? throw new InvalidOperationException("Failed to parse configuration file");
        }

        private static void SaveConfiguration(JsonNode config)
        {
            // Create backup
            string backupPath = ConfigPath + ".backup";
            if (File.Exists(ConfigPath))
            {
                File.Copy(ConfigPath, backupPath, true);
            }

            // Save with pretty formatting
            var options = new JsonSerializerOptions
            {
                WriteIndented = true
            };

            string json = config.ToJsonString(options);
            File.WriteAllText(ConfigPath, json);
        }

        private static string GenerateAzureEngineId(string voiceName)
        {
            // Convert "en-GB-ElliotNeural" to "azure-elliot"
            string[] parts = voiceName.Split('-');
            if (parts.Length >= 3)
            {
                string name = parts[2].Replace("Neural", "").ToLower();
                return $"azure-{name}";
            }

            // Fallback
            return $"azure-{voiceName.ToLower().Replace("-", "").Replace("neural", "")}";
        }

        private static string GenerateShortName(string voiceName)
        {
            // Convert "en-GB-ElliotNeural" to "elliot"
            string[] parts = voiceName.Split('-');
            if (parts.Length >= 3)
            {
                return parts[2].Replace("Neural", "").ToLower();
            }

            // Fallback
            return voiceName.ToLower().Replace("-", "").Replace("neural", "");
        }

        private static string GenerateSherpaEngineId(string voiceId)
        {
            // Convert "piper-en-amy-low" to "sherpa-amy-low"
            if (voiceId.StartsWith("piper-en-"))
            {
                return voiceId.Replace("piper-en-", "sherpa-");
            }
            else if (voiceId.StartsWith("piper-"))
            {
                return voiceId.Replace("piper-", "sherpa-");
            }
            else if (!voiceId.StartsWith("sherpa-"))
            {
                return $"sherpa-{voiceId}";
            }

            return voiceId; // Already has sherpa prefix
        }

        private static string GenerateSherpaShortName(string voiceId, string voiceName)
        {
            // Try to extract a user-friendly short name
            // "piper-en-amy-low" -> "amy"
            // "piper-en-danny-low" -> "danny"

            if (voiceId.StartsWith("piper-en-"))
            {
                string remainder = voiceId.Substring("piper-en-".Length);
                string[] parts = remainder.Split('-');
                if (parts.Length > 0)
                {
                    return parts[0].ToLower(); // Return the name part (amy, danny, etc.)
                }
            }

            // Fallback to voice name if available
            if (!string.IsNullOrEmpty(voiceName))
            {
                return voiceName.ToLower().Replace(" ", "");
            }

            // Final fallback
            return voiceId.ToLower().Replace("-", "").Replace("piper", "").Replace("en", "");
        }
    }
}
