using System;
using System.IO;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.Win32.SafeHandles;
using System.Runtime.InteropServices;
using System.Collections.Generic;

namespace Installer
{
    /// <summary>
    /// Bridge for communicating with AACSpeakHelper pipe service
    /// Handles sending TTS requests and receiving audio data
    /// </summary>
    public class PipeServiceBridge : IDisposable
    {
        private const string PipeName = @"\\.\pipe\AACSpeakHelper";
        private const int BufferSize = 64 * 1024; // 64KB buffer

        // Win32 API imports for named pipe communication
        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern SafeFileHandle CreateFile(
            string lpFileName,
            uint dwDesiredAccess,
            uint dwShareMode,
            IntPtr lpSecurityAttributes,
            uint dwCreationDisposition,
            uint dwFlagsAndAttributes,
            IntPtr hTemplateFile);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool WriteFile(
            SafeFileHandle hFile,
            byte[] lpBuffer,
            uint nNumberOfBytesToWrite,
            out uint lpNumberOfBytesWritten,
            IntPtr lpOverlapped);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool ReadFile(
            SafeFileHandle hFile,
            byte[] lpBuffer,
            uint nNumberOfBytesToRead,
            out uint lpNumberOfBytesRead,
            IntPtr lpOverlapped);

        private const uint GENERIC_READ = 0x80000000;
        private const uint GENERIC_WRITE = 0x40000000;
        private const uint OPEN_EXISTING = 3;

        /// <summary>
        /// Message structure for pipe communication
        /// Matches the format expected by AACSpeakHelper
        /// </summary>
        public class PipeMessage
        {
            public Dictionary<string, object> args { get; set; } = new Dictionary<string, object>();
            public Dictionary<string, Dictionary<string, string>> config { get; set; } = new Dictionary<string, Dictionary<string, string>>();
            public string clipboard_text { get; set; }
        }

        /// <summary>
        /// Sends a TTS request to AACSpeakHelper pipe service
        /// </summary>
        /// <param name="text">Text to synthesize</param>
        /// <param name="voiceConfig">Voice configuration</param>
        /// <returns>True if request was sent successfully</returns>
        public async Task<bool> SynthesizeTextAsync(string text, ConfigBasedVoiceManager.PipeVoiceConfig voiceConfig)
        {
            try
            {
                // Create pipe message
                var message = CreatePipeMessage(text, voiceConfig);
                
                // Serialize to JSON
                string jsonMessage = JsonSerializer.Serialize(message, new JsonSerializerOptions 
                { 
                    WriteIndented = false,
                    PropertyNamingPolicy = JsonNamingPolicy.CamelCase
                });

                // Send to pipe
                return await SendToPipeAsync(jsonMessage);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error synthesizing text: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Creates a pipe message from voice configuration
        /// </summary>
        private PipeMessage CreatePipeMessage(string text, ConfigBasedVoiceManager.PipeVoiceConfig voiceConfig)
        {
            var message = new PipeMessage
            {
                clipboard_text = text,
                args = new Dictionary<string, object>
                {
                    { "listvoices", false },
                    { "preview", false },
                    { "style", voiceConfig.TtsConfig?.AzureTTS?.Style ?? "" },
                    { "styledegree", null },
                    { "text", text },
                    { "verbose", false }
                }
            };

            // Build configuration sections based on voice config
            if (voiceConfig.TtsConfig != null)
            {
                // TTS section
                if (voiceConfig.TtsConfig.TTS != null)
                {
                    message.config["TTS"] = new Dictionary<string, string>
                    {
                        { "engine", voiceConfig.TtsConfig.TTS.Engine ?? "azure" },
                        { "voice_id", voiceConfig.TtsConfig.TTS.VoiceId ?? "" },
                        { "bypass_tts", voiceConfig.TtsConfig.TTS.BypassTts.ToString().ToLower() }
                    };
                }

                // Azure TTS section
                if (voiceConfig.TtsConfig.AzureTTS != null)
                {
                    message.config["azureTTS"] = new Dictionary<string, string>
                    {
                        { "key", voiceConfig.TtsConfig.AzureTTS.Key ?? "" },
                        { "location", voiceConfig.TtsConfig.AzureTTS.Location ?? "uksouth" },
                        { "voice", voiceConfig.TtsConfig.AzureTTS.Voice ?? "" },
                        { "style", voiceConfig.TtsConfig.AzureTTS.Style ?? "" },
                        { "role", voiceConfig.TtsConfig.AzureTTS.Role ?? "" }
                    };
                }

                // Google TTS section
                if (voiceConfig.TtsConfig.GoogleTTS != null)
                {
                    message.config["googleTTS"] = new Dictionary<string, string>
                    {
                        { "creds", voiceConfig.TtsConfig.GoogleTTS.Creds ?? "" },
                        { "voice", voiceConfig.TtsConfig.GoogleTTS.Voice ?? "" },
                        { "lang", voiceConfig.TtsConfig.GoogleTTS.Lang ?? "" }
                    };
                }

                // Translation section
                if (voiceConfig.TtsConfig.Translate != null)
                {
                    message.config["translate"] = new Dictionary<string, string>
                    {
                        { "no_translate", voiceConfig.TtsConfig.Translate.NoTranslate.ToString().ToLower() },
                        { "provider", voiceConfig.TtsConfig.Translate.Provider ?? "" },
                        { "start_lang", voiceConfig.TtsConfig.Translate.StartLang ?? "auto" },
                        { "end_lang", voiceConfig.TtsConfig.Translate.EndLang ?? "en" },
                        { "replace_pb", voiceConfig.TtsConfig.Translate.ReplacePb.ToString().ToLower() }
                    };
                }
            }

            // Add App section with paths
            message.config["App"] = new Dictionary<string, string>
            {
                { "config_path", "" },
                { "audio_files_path", "" }
            };

            return message;
        }

        /// <summary>
        /// Sends data to the named pipe
        /// </summary>
        private async Task<bool> SendToPipeAsync(string jsonMessage)
        {
            const int maxRetries = 3;
            const int retryDelayMs = 1000;

            for (int attempt = 0; attempt < maxRetries; attempt++)
            {
                try
                {
                    using (var handle = CreateFile(
                        PipeName,
                        GENERIC_READ | GENERIC_WRITE,
                        0,
                        IntPtr.Zero,
                        OPEN_EXISTING,
                        0,
                        IntPtr.Zero))
                    {
                        if (handle.IsInvalid)
                        {
                            int error = Marshal.GetLastWin32Error();
                            Console.WriteLine($"Failed to connect to pipe (attempt {attempt + 1}): Error {error}");
                            
                            if (attempt < maxRetries - 1)
                            {
                                await Task.Delay(retryDelayMs);
                                continue;
                            }
                            return false;
                        }

                        // Send message
                        byte[] messageBytes = Encoding.UTF8.GetBytes(jsonMessage);
                        bool writeResult = WriteFile(handle, messageBytes, (uint)messageBytes.Length, out uint bytesWritten, IntPtr.Zero);
                        
                        if (!writeResult)
                        {
                            Console.WriteLine($"Failed to write to pipe: Error {Marshal.GetLastWin32Error()}");
                            return false;
                        }

                        Console.WriteLine($"Successfully sent {bytesWritten} bytes to AACSpeakHelper pipe service");
                        
                        // Try to read response (optional, for voice list requests)
                        byte[] responseBuffer = new byte[BufferSize];
                        bool readResult = ReadFile(handle, responseBuffer, BufferSize, out uint bytesRead, IntPtr.Zero);
                        
                        if (readResult && bytesRead > 0)
                        {
                            string response = Encoding.UTF8.GetString(responseBuffer, 0, (int)bytesRead);
                            Console.WriteLine($"Received response: {response.Substring(0, Math.Min(100, response.Length))}...");
                        }

                        return true;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error communicating with pipe (attempt {attempt + 1}): {ex.Message}");
                    
                    if (attempt < maxRetries - 1)
                    {
                        await Task.Delay(retryDelayMs);
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// Tests connection to AACSpeakHelper pipe service
        /// </summary>
        public async Task<bool> TestConnectionAsync()
        {
            try
            {
                // Create a simple test message
                var testMessage = new PipeMessage
                {
                    clipboard_text = "Test connection",
                    args = new Dictionary<string, object>
                    {
                        { "listvoices", true },
                        { "verbose", true }
                    },
                    config = new Dictionary<string, Dictionary<string, string>>
                    {
                        ["TTS"] = new Dictionary<string, string>
                        {
                            { "engine", "azure" },
                            { "bypass_tts", "true" }
                        },
                        ["translate"] = new Dictionary<string, string>
                        {
                            { "no_translate", "true" }
                        }
                    }
                };

                string jsonMessage = JsonSerializer.Serialize(testMessage);
                return await SendToPipeAsync(jsonMessage);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Connection test failed: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Checks if AACSpeakHelper pipe service is running
        /// </summary>
        public bool IsServiceRunning()
        {
            try
            {
                using (var handle = CreateFile(
                    PipeName,
                    GENERIC_READ,
                    0,
                    IntPtr.Zero,
                    OPEN_EXISTING,
                    0,
                    IntPtr.Zero))
                {
                    return !handle.IsInvalid;
                }
            }
            catch
            {
                return false;
            }
        }

        public void Dispose()
        {
            // No resources to dispose in this implementation
            GC.SuppressFinalize(this);
        }
    }
}
