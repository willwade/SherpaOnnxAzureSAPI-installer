using System;
using System.Diagnostics;
using System.IO;

namespace SignAssembly
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Signing sherpa-onnx assembly...");
            
            try
            {
                // Get the path to the sherpa-onnx assembly
                string assemblyPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                    "OpenAssistive", "OpenSpeech", "sherpa-onnx.dll");
                
                // Get the path to the key file
                string keyFilePath = Path.Combine(
                    Directory.GetParent(Directory.GetCurrentDirectory()).FullName,
                    "KeyGenerator", "OpenSpeechKey.snk");
                
                Console.WriteLine($"Assembly path: {assemblyPath}");
                Console.WriteLine($"Key file path: {keyFilePath}");
                
                if (!File.Exists(assemblyPath))
                {
                    Console.WriteLine("Error: sherpa-onnx.dll not found.");
                    return;
                }
                
                if (!File.Exists(keyFilePath))
                {
                    Console.WriteLine("Error: OpenSpeechKey.snk not found.");
                    return;
                }
                
                // Create a backup of the original assembly
                string backupPath = assemblyPath + ".bak";
                File.Copy(assemblyPath, backupPath, true);
                Console.WriteLine($"Created backup at: {backupPath}");
                
                // Use ildasm to disassemble the assembly
                string ilPath = Path.ChangeExtension(assemblyPath, ".il");
                string resPath = Path.ChangeExtension(assemblyPath, ".res");
                
                RunProcess("ildasm", $"/out=\"{ilPath}\" \"{assemblyPath}\" /nobar");
                Console.WriteLine("Disassembled assembly to IL");
                
                // Use ilasm to reassemble the assembly with the strong name key
                RunProcess("ilasm", $"\"{ilPath}\" /dll /key=\"{keyFilePath}\" /output=\"{assemblyPath}\" /resource=\"{resPath}\"");
                Console.WriteLine("Reassembled assembly with strong name key");
                
                // Clean up temporary files
                if (File.Exists(ilPath))
                    File.Delete(ilPath);
                
                if (File.Exists(resPath))
                    File.Delete(resPath);
                
                Console.WriteLine("Assembly signed successfully!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error signing assembly: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
        
        static void RunProcess(string fileName, string arguments)
        {
            var startInfo = new ProcessStartInfo
            {
                FileName = fileName,
                Arguments = arguments,
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
                
                if (process.ExitCode != 0)
                {
                    Console.WriteLine($"Error running {fileName}: {error}");
                    throw new Exception($"Process {fileName} exited with code {process.ExitCode}");
                }
                
                if (!string.IsNullOrEmpty(output))
                    Console.WriteLine(output);
            }
        }
    }
}
