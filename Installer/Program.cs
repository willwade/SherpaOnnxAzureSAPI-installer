using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Newtonsoft.Json;

class Program
{
    private const string OnlineJsonUrl = "https://github.com/willwade/tts-wrapper/raw/main/tts_wrapper/engines/sherpaonnx/merged_models.json";
    private const string LocalJsonPath = "./merged_models.json";

    static async Task Main(string[] args)
    {
        var installer = new ModelInstaller();
        var registrar = new Sapi5Registrar();
        string dllPath = @"C:\Program Files\OpenAssistive\OpenSpeech\OpenSpeechTTS.dll";

        // Check for uninstall argument
        if (args.Length > 0 && args[0] == "uninstall")
        {
            await UninstallVoicesAndDll(registrar, dllPath);
            return;
        }

        // Load models JSON
        var models = await LoadModelsAsync();
        Console.WriteLine("Available voices to install:");
        foreach (var model in models.Values)
        {
            Console.WriteLine($" - {model.Id} ({model.Name})");
        }

        // Simulate a user choice (this would come from GUI)
        Console.WriteLine("Enter the model ID to install (e.g., 'cantonese-fs-xiaomaiiwn'):");
        string chosenModelId = Console.ReadLine();

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
            Console.WriteLine("Invalid model ID entered.");
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
  
