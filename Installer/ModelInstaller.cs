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
            await downloader.DownloadFileTaskAsync($"{model.Url}/model.onnx", Path.Combine(modelDir, "model.onnx"));

            Console.WriteLine("Downloading tokens.txt...");
            await downloader.DownloadFileTaskAsync($"{model.Url}/tokens.txt", Path.Combine(modelDir, "tokens.txt"));

            Console.WriteLine($"Downloaded model and tokens for {model.Id}.");
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
                File.Copy(localModelPath, Path.Combine(modelDir, "model.onnx"), overwrite: true);
                File.Copy(localTokensPath, Path.Combine(modelDir, "tokens.txt"), overwrite: true);
                Console.WriteLine("Successfully used the local copy.");
            }
            else
            {
                Console.WriteLine("Local copy not found. Cannot continue installation.");
                throw;
            }
        }
    }
}
