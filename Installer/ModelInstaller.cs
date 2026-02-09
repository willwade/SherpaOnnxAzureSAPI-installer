using System;
using System.IO;
using System.Threading.Tasks;
using System.IO.Compression;
using System.Net.Http;
using ICSharpCode.SharpZipLib.BZip2;
using ICSharpCode.SharpZipLib.Tar;

public class ModelInstaller
{
    private readonly string modelsDirectory;
    private readonly string tempDirectory;

    public ModelInstaller()
    {
        // Use Program Files for all users
        modelsDirectory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
            "OpenSpeech",
            "models"
        );
        
        tempDirectory = Path.Combine(Path.GetTempPath(), "OpenSpeech");
        
        // Ensure directories exist with proper permissions
        Directory.CreateDirectory(modelsDirectory);
        Directory.CreateDirectory(tempDirectory);
    }

    public async Task DownloadAndExtractModelAsync(TtsModel model)
    {
        string modelDir = Path.Combine(modelsDirectory, model.Id);
        Console.WriteLine($"Creating model directory: {modelDir}");
        Directory.CreateDirectory(modelDir);

        string modelPath = Path.Combine(modelDir, "model.onnx");
        string tokensPath = Path.Combine(modelDir, "tokens.txt");
        string archivePath = Path.Combine(tempDirectory, $"{model.Id}.tar.bz2");

        Console.WriteLine($"Model will be downloaded to: {modelPath}");
        Console.WriteLine($"Tokens will be downloaded to: {tokensPath}");

        try
        {
            // Download the tar.bz2 archive
            Console.WriteLine($"Downloading archive from {model.Url}");
            using (var client = new HttpClient())
            {
                var response = await client.GetAsync(model.Url);
                response.EnsureSuccessStatusCode();
                using (var fs = new FileStream(archivePath, FileMode.Create))
                {
                    await response.Content.CopyToAsync(fs);
                }
            }

            // Extract files from tar.bz2
            Console.WriteLine("Extracting files from archive...");
            using (var fs = File.OpenRead(archivePath))
            using (var bz2 = new BZip2InputStream(fs))
            using (var tar = new TarInputStream(bz2))
            {
                TarEntry entry;
                bool foundModel = false;
                bool foundTokens = false;

                while ((entry = tar.GetNextEntry()) != null)
                {
                    string entryName = entry.Name.ToLower();
                    Console.WriteLine($"Found entry: {entry.Name}");

                    if (entryName.EndsWith(".onnx") && !entryName.EndsWith(".onnx.json"))
                    {
                        using (var outStream = File.Create(modelPath))
                        {
                            tar.CopyEntryContents(outStream);
                        }
                        Console.WriteLine($"Extracted model file to {modelPath}");
                        foundModel = true;
                    }
                    else if (entryName.EndsWith("tokens.txt"))
                    {
                        using (var outStream = File.Create(tokensPath))
                        {
                            tar.CopyEntryContents(outStream);
                        }
                        Console.WriteLine($"Extracted tokens.txt to {tokensPath}");
                        foundTokens = true;
                    }

                    if (foundModel && foundTokens)
                        break;
                }

                if (!foundModel)
                    throw new Exception("ONNX model file not found in archive");
                if (!foundTokens)
                    throw new Exception("tokens.txt not found in archive");
            }

            // Clean up temp file
            try
            {
                File.Delete(archivePath);
            }
            catch { }

            // Verify files exist and have content
            if (!File.Exists(modelPath) || new FileInfo(modelPath).Length == 0)
                throw new Exception($"Model file not found or empty at {modelPath}");
            if (!File.Exists(tokensPath) || new FileInfo(tokensPath).Length == 0)
                throw new Exception($"Tokens file not found or empty at {tokensPath}");

            // Set the paths in the model object
            model.ModelPath = modelPath;
            model.TokensPath = tokensPath;
            model.LexiconPath = ""; // Leave empty for MMS models

            Console.WriteLine($"Successfully downloaded and extracted files for {model.Id}.");
            Console.WriteLine($"Model path: {model.ModelPath}");
            Console.WriteLine($"Tokens path: {model.TokensPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to download/extract model: {ex.Message}");
            Console.WriteLine("Falling back to local copy...");

            // Fallback to local copy
            string localModelPath = Path.Combine("local", model.Id, "model.onnx");
            string localTokensPath = Path.Combine("local", model.Id, "tokens.txt");

            if (File.Exists(localModelPath) && File.Exists(localTokensPath))
            {
                File.Copy(localModelPath, modelPath, overwrite: true);
                File.Copy(localTokensPath, tokensPath, overwrite: true);

                // Set the paths in the model object
                model.ModelPath = modelPath;
                model.TokensPath = tokensPath;
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
