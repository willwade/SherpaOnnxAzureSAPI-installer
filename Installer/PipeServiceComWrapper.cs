using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.Json;
using Microsoft.Win32;

namespace Installer
{
    /// <summary>
    /// COM wrapper for pipe service voices
    /// This class implements the basic COM interface needed for SAPI integration
    /// and forwards TTS requests to the AACSpeakHelper pipe service
    /// </summary>
    [ComVisible(true)]
    [Guid("4A8B9C2D-1E3F-4567-8901-234567890ABC")]
    [ClassInterface(ClassInterfaceType.None)]
    public class PipeServiceComWrapper : IDisposable
    {
        private PipeServiceBridge _bridge;
        private ConfigBasedVoiceManager.PipeVoiceConfig _currentVoiceConfig;
        private bool _disposed = false;

        public PipeServiceComWrapper()
        {
            _bridge = new PipeServiceBridge();
        }

        /// <summary>
        /// Initializes the COM wrapper with a specific voice configuration
        /// </summary>
        /// <param name="voiceName">Name of the voice to use</param>
        /// <returns>True if initialization successful</returns>
        public bool Initialize(string voiceName)
        {
            try
            {
                // Load voice configuration
                string configPath = Path.Combine(
                    @"C:\Program Files\OpenAssistive\OpenSpeech\voice_configs",
                    $"{voiceName}.json"
                );

                if (!File.Exists(configPath))
                {
                    Console.WriteLine($"Voice configuration not found: {configPath}");
                    return false;
                }

                string jsonContent = File.ReadAllText(configPath);
                _currentVoiceConfig = JsonSerializer.Deserialize<ConfigBasedVoiceManager.PipeVoiceConfig>(jsonContent);

                if (_currentVoiceConfig == null)
                {
                    Console.WriteLine($"Failed to parse voice configuration: {configPath}");
                    return false;
                }

                Console.WriteLine($"Initialized pipe service COM wrapper for voice: {_currentVoiceConfig.DisplayName}");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error initializing COM wrapper: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Synthesizes text using the configured voice
        /// </summary>
        /// <param name="text">Text to synthesize</param>
        /// <returns>True if synthesis was successful</returns>
        public bool SynthesizeText(string text)
        {
            try
            {
                if (_currentVoiceConfig == null)
                {
                    Console.WriteLine("COM wrapper not initialized with a voice configuration");
                    return false;
                }

                if (string.IsNullOrEmpty(text))
                {
                    Console.WriteLine("No text provided for synthesis");
                    return false;
                }

                // Check if pipe service is running
                if (!_bridge.IsServiceRunning())
                {
                    Console.WriteLine("AACSpeakHelper pipe service is not running");
                    return false;
                }

                // Send synthesis request
                var task = _bridge.SynthesizeTextAsync(text, _currentVoiceConfig);
                task.Wait(); // Synchronous call for COM compatibility

                return task.Result;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error synthesizing text: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Gets information about the current voice
        /// </summary>
        /// <returns>Voice information string</returns>
        public string GetVoiceInfo()
        {
            if (_currentVoiceConfig == null)
                return "No voice configured";

            return $"{_currentVoiceConfig.DisplayName} - {_currentVoiceConfig.Description}";
        }

        /// <summary>
        /// Tests connection to the pipe service
        /// </summary>
        /// <returns>True if connection test successful</returns>
        public bool TestConnection()
        {
            try
            {
                var task = _bridge.TestConnectionAsync();
                task.Wait();
                return task.Result;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Connection test failed: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Registers this COM component
        /// </summary>
        /// <param name="type">Type being registered</param>
        [ComRegisterFunction]
        public static void RegisterFunction(Type type)
        {
            try
            {
                // Register the COM class
                string clsid = "{4A8B9C2D-1E3F-4567-8901-234567890ABC}";
                
                using (var clsidKey = Registry.ClassesRoot.CreateSubKey($@"CLSID\{clsid}"))
                {
                    clsidKey.SetValue("", "PipeServiceComWrapper");
                    clsidKey.SetValue("AppID", clsid);

                    using (var inprocKey = clsidKey.CreateSubKey("InprocServer32"))
                    {
                        inprocKey.SetValue("", System.Reflection.Assembly.GetExecutingAssembly().Location);
                        inprocKey.SetValue("ThreadingModel", "Apartment");
                    }

                    using (var progIdKey = clsidKey.CreateSubKey("ProgId"))
                    {
                        progIdKey.SetValue("", "PipeServiceComWrapper.1");
                    }
                }

                // Register ProgID
                using (var progIdKey = Registry.ClassesRoot.CreateSubKey("PipeServiceComWrapper.1"))
                {
                    progIdKey.SetValue("", "PipeServiceComWrapper");
                    
                    using (var clsidRefKey = progIdKey.CreateSubKey("CLSID"))
                    {
                        clsidRefKey.SetValue("", clsid);
                    }
                }

                Console.WriteLine("PipeServiceComWrapper registered successfully");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error registering COM component: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Unregisters this COM component
        /// </summary>
        /// <param name="type">Type being unregistered</param>
        [ComUnregisterFunction]
        public static void UnregisterFunction(Type type)
        {
            try
            {
                string clsid = "{4A8B9C2D-1E3F-4567-8901-234567890ABC}";
                
                // Remove CLSID registration
                Registry.ClassesRoot.DeleteSubKeyTree($@"CLSID\{clsid}", false);
                
                // Remove ProgID registration
                Registry.ClassesRoot.DeleteSubKeyTree("PipeServiceComWrapper.1", false);

                Console.WriteLine("PipeServiceComWrapper unregistered successfully");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error unregistering COM component: {ex.Message}");
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    _bridge?.Dispose();
                }
                _disposed = true;
            }
        }

        ~PipeServiceComWrapper()
        {
            Dispose(false);
        }
    }
}
