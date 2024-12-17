using System;
using Downloader;
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
            ReserveStorageSpaceBeforeStartingDownload = true, // Reserve file space
        });

        string modelUrl = $"{model.Url}/model.onnx";
        string tokensUrl = $"{model.Url}/tokens.txt";

        // Download model.onnx
        string modelFilePath = Path.Combine(modelDir, "model.onnx");
        await downloader.DownloadFileTaskAsync(modelUrl, modelFilePath);

        // Download tokens.txt
        string tokensFilePath = Path.Combine(modelDir, "tokens.txt");
        await downloader.DownloadFileTaskAsync(tokensUrl, tokensFilePath);

        Console.WriteLine($"Downloaded model and tokens to: {modelDir}");
    }
}
