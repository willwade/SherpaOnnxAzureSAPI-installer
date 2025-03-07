using System;
using System.Diagnostics;
using System.IO;

namespace KeyGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Generating strong name key file...");
            
            try
            {
                // Use the sn.exe tool from the .NET SDK to generate a key file
                string keyFilePath = Path.Combine(Directory.GetCurrentDirectory(), "OpenSpeechKey.snk");
                
                // Create a process to run the sn.exe tool
                var startInfo = new ProcessStartInfo
                {
                    FileName = "cmd.exe",
                    Arguments = $"/c dotnet build -t:GenerateKeyFile /p:KeyFile={keyFilePath}",
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };
                
                using (var process = Process.Start(startInfo))
                {
                    string output = process.StandardOutput.ReadToEnd();
                    string error = process.StandardError.ReadToEnd();
                    process.WaitForExit();
                    
                    if (process.ExitCode == 0)
                    {
                        Console.WriteLine($"Key file '{keyFilePath}' created successfully.");
                    }
                    else
                    {
                        Console.WriteLine($"Error generating key: {error}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error generating key: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
