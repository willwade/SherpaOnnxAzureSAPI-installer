using System;
using System.Runtime.InteropServices;

namespace DllExports
{
    class Program
    {
        [DllImport("kernel32.dll", SetLastError = true)]
        static extern IntPtr LoadLibrary(string lpFileName);

        [DllImport("kernel32.dll")]
        static extern IntPtr GetProcAddress(IntPtr hModule, string lpProcName);

        static void Main(string[] args)
        {
            string dllPath = "sherpa-onnx.dll";
            IntPtr hModule = LoadLibrary(dllPath);
            if (hModule == IntPtr.Zero)
            {
                int error = Marshal.GetLastWin32Error();
                Console.WriteLine($"Failed to load DLL. Error: {error}");
                return;
            }

            Console.WriteLine("DLL loaded successfully. Looking for TTS-related exports...");

            // Try different variations of the function name
            string[] possibleNames = new[]
            {
                "SherpaOnnxCreateOfflineTts",
                "sherpa_onnx_offline_tts_create",
                "sherpa_onnx_create_offline_tts",
                "CreateOfflineTts",
                "create_offline_tts"
            };

            foreach (var name in possibleNames)
            {
                IntPtr procAddress = GetProcAddress(hModule, name);
                if (procAddress != IntPtr.Zero)
                {
                    Console.WriteLine($"Found export: {name}");
                }
            }
        }
    }
}
