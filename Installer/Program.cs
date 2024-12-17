using System;
using System.Collections.Generic;
using System.IO;
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

        // Default install behavior
        var models = await LoadModelsAsync();
        foreach (var model in models.Values)
        {
            try
            {
                await installer.DownloadAndExtractModelAsync(model);
                registrar.RegisterVoice(model, dllPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to install {model.Name}: {ex.Message}");
            }
        }
    }

    static async Task<Dictionary<string, TtsModel>> LoadModelsAsync()
    {
        try
        {
            using var client = new System.Net.Http.HttpClient();
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
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error uninstalling {model.Name}: {ex.Message}");
            }
        }

        // Check if any models remain; if not, unregister DLL
        var modelDirs = Directory.Exists("./models") ? Directory.GetDirectories("./models") : Array.Empty<string>();
        if (modelDirs.Length == 0)
        {
            UnregisterDll(dllPath);
        }

        Console.WriteLine("Uninstallation complete.");
    }

    private static void UnregisterDll(string dllPath)
    {
        try
        {
            var process = new System.Diagnostics.Process();
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
}
