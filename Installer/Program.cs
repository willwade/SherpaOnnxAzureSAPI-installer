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
                    Console.WriteLine("Invalid command. Use 'install <model-id>' or 'uninstall <model-id|all>'");
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
                    await installer.DownloadAndExtractModelAsync(chosenModel);
                    registrar.RegisterVoice(chosenModel, dllPath);
                    Console.WriteLine($"Successfully installed voice: {chosenModel.Name}");
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
                registrar.UnregisterVoice(model.Id, model.Name);
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
        try
        {
            var models = await LoadModelsAsync();
            if (models.TryGetValue(modelId, out var model))
            {
                Console.WriteLine($"Installing voice: {model.Name}...");
                await installer.DownloadAndExtractModelAsync(model);
                registrar.RegisterVoice(model, dllPath);
                Console.WriteLine($"Successfully installed voice: {model.Name}");
            }
            else
            {
                Console.WriteLine($"Model ID '{modelId}' not found in available voices.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to install voice: {ex.Message}");
        }
    }

    private static async Task UninstallSpecificVoice(string modelId, Sapi5Registrar registrar)
    {
        try
        {
            var models = await LoadModelsAsync();
            if (models.TryGetValue(modelId, out var model))
            {
                registrar.UnregisterVoice(model.Id, model.Name);
                Console.WriteLine($"Successfully uninstalled voice: {model.Name}");
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

    private static bool IsRunningAsAdministrator()
    {
        using (var identity = System.Security.Principal.WindowsIdentity.GetCurrent())
        {
            var principal = new System.Security.Principal.WindowsPrincipal(identity);
            return principal.IsInRole(System.Security.Principal.WindowsBuiltInRole.Administrator);
        }
    }
}
