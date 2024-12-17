using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Newtonsoft.Json;

class Program
{
    private const string OnlineJsonUrl = "https://github.com/willwade/tts-wrapper/raw/main/tts_wrapper/engines/sherpaonnx/merged_models.json";
    private const string LocalJsonPath = "./merged_models.json";

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
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to download merged_models.json: {ex.Message}");

            // Attempt to load local copy
            if (File.Exists(LocalJsonPath))
            {
                Console.WriteLine("Using cached local copy of merged_models.json.");
                var json = File.ReadAllText(LocalJsonPath);
                return JsonConvert.DeserializeObject<Dictionary<string, TtsModel>>(json);
            }
            else
            {
                Console.WriteLine("Local merged_models.json not found. Cannot proceed.");
                throw new FileNotFoundException("merged_models.json is missing locally and online.");
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
