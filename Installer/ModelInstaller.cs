using Downloader;
using System;
using System.IO;
using System.Threading.Tasks;

public class ModelInstaller
{
    private readonly string modelsDirectory = "./models";

    public async Task DownloadAndExtractModelAsync(TtsModel model)
    {
        string modelDir = Path.Combine(modelsDirectory, model.Id);
        Directory.CreateDirectory(modelDir);

        string modelPath = Path.Combine(modelDir, "model.onnx");
        string tokensPath = Path.Combine(modelDir, "tokens.txt");

        var downloader = new DownloadService(new DownloadConfiguration
        {
            ChunkCount = 4, // Download in parallel chunks
            ParallelDownload = true, // Enable parallel downloads
            ReserveStorageSpaceBeforeStartingDownload = true
        });

        try
        {
            // Attempt to download model files
            Console.WriteLine("Downloading model.onnx...");
            await downloader.DownloadFileTaskAsync($"{model.Url}/model.onnx", modelPath);

            Console.WriteLine("Downloading tokens.txt...");
            await downloader.DownloadFileTaskAsync($"{model.Url}/tokens.txt", tokensPath);

            // Set the paths in the model object
            model.ModelPath = Path.GetFullPath(modelPath);
            model.TokensPath = Path.GetFullPath(tokensPath);
            model.LexiconPath = ""; // Leave empty for MMS models

            Console.WriteLine($"Downloaded model and tokens for {model.Id}.");
            Console.WriteLine($"Model path: {model.ModelPath}");
            Console.WriteLine($"Tokens path: {model.TokensPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to download model: {ex.Message}");
            Console.WriteLine("Falling back to local copy...");

            // Fallback to local copy
            string localModelPath = $"./local/{model.Id}/model.onnx";
            string localTokensPath = $"./local/{model.Id}/tokens.txt";

            if (File.Exists(localModelPath) && File.Exists(localTokensPath))
            {
                File.Copy(localModelPath, modelPath, overwrite: true);
                File.Copy(localTokensPath, tokensPath, overwrite: true);

                // Set the paths in the model object
                model.ModelPath = Path.GetFullPath(modelPath);
                model.TokensPath = Path.GetFullPath(tokensPath);
                model.LexiconPath = ""; // Leave empty for MMS models

                Console.WriteLine("Successfully used the local copy.");
                Console.WriteLine($"Model path: {model.ModelPath}");
                Console.WriteLine($"Tokens path: {model.TokensPath}");
            }
            else
            {
                Console.WriteLine("Local copy not found. Cannot continue installation.");
                throw;
            }
        }
    }
}
