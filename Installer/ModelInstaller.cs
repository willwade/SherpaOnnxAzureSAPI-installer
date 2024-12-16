
// This file contains the logic for downloading and extracting TTS models.
using System;
using System.IO;
using System.Threading.Tasks;
using Downloader;
using SharpCompress.Archives;
using SharpCompress.Common;

public class ModelInstaller
{
    private readonly string modelsDirectory = "./models";

    public async Task DownloadAndExtractModelAsync(TtsModel model)
    {
        string modelDir = Path.Combine(modelsDirectory, model.Id);
        Directory.CreateDirectory(modelDir);

        if (model.ModelType.Equals("mms", StringComparison.OrdinalIgnoreCase))
        {
            await DownloadMmsModelAsync(model.Url, modelDir);
        }
        else
        {
            string archivePath = Path.Combine(modelDir, "model.tar.bz2");
            await DownloadFileAsync(model.Url, archivePath);
            ExtractArchive(archivePath, modelDir);
            File.Delete(archivePath);
        }
    }

    private async Task DownloadMmsModelAsync(string baseUrl, string destinationDirectory)
    {
        string modelUrl = $"{baseUrl}/model.onnx";
        string tokensUrl = $"{baseUrl}/tokens.txt";

        await DownloadFileAsync(modelUrl, Path.Combine(destinationDirectory, "model.onnx"));
        await DownloadFileAsync(tokensUrl, Path.Combine(destinationDirectory, "tokens.txt"));
    }

    private async Task DownloadFileAsync(string url, string destinationPath)
    {
        var downloadService = new DownloadService();
        var downloadPackage = new DownloadPackage
        {
            Address = url,
            FileName = Path.GetFileName(destinationPath),
            FilePath = Path.GetDirectoryName(destinationPath)
        };

        downloadService.DownloadProgressChanged += (s, e) =>
        {
            Console.WriteLine($"Downloading {e.ProgressPercentage}%");
        };

        await downloadService.DownloadFileTaskAsync(downloadPackage);
        Console.WriteLine("Download complete.");
    }

    private void ExtractArchive(string archivePath, string destinationDirectory)
    {
        using (var archive = ArchiveFactory.Open(archivePath))
        {
            foreach (var entry in archive.Entries)
            {
                if (!entry.IsDirectory)
                {
                    entry.WriteToDirectory(destinationDirectory, new ExtractionOptions
                    {
                        ExtractFullPath = true,
                        Overwrite = true
                    });
                }
            }
        }
        Console.WriteLine("Extraction complete.");
    }
}
