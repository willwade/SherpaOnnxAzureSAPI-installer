using System;
using Microsoft.Win32;
using System.IO;

namespace Installer
{
    public class Sapi5Registrar
    {
        private const string RegistryBasePath = @"SOFTWARE\Microsoft\SPEECH\Voices\Tokens";

        public void RegisterVoice(TtsModel model, string dllPath)
        {
            string voiceRegistryPath = $@"{RegistryBasePath}\{model.Id}";
            string clsid = "3d8f5c5d-9d6b-4b92-a12b-1a6dff80b6b2";
            string clsidPath = $@"HKEY_CLASSES_ROOT\CLSID\{clsid}";
            string inprocServer32Path = $@"HKEY_CLASSES_ROOT\CLSID\{clsid}\InprocServer32";

            try
            {
                // 1. Register the SAPI voice
                Registry.SetValue($@"HKEY_LOCAL_MACHINE\{voiceRegistryPath}", "", model.Name);
                Registry.SetValue($@"HKEY_LOCAL_MACHINE\{voiceRegistryPath}", "Lang", "409"); // Example: LCID for English
                Registry.SetValue($@"HKEY_LOCAL_MACHINE\{voiceRegistryPath}", "Gender", "Neutral");
                Registry.SetValue($@"HKEY_LOCAL_MACHINE\{voiceRegistryPath}", "Age", "Adult");
                Registry.SetValue($@"HKEY_LOCAL_MACHINE\{voiceRegistryPath}", "Vendor", model.Developer);
                Registry.SetValue($@"HKEY_LOCAL_MACHINE\{voiceRegistryPath}", "Version", "1.0");

                // Add Attributes subkey with CLSID and VoicePath
                string attributesPath = $@"HKEY_LOCAL_MACHINE\{voiceRegistryPath}\Attributes";
                Registry.SetValue(attributesPath, "CLSID", clsid);
                Registry.SetValue(attributesPath, "VoicePath", dllPath);

                // 2. Register the COM class CLSID with InprocServer32
                Registry.SetValue(clsidPath, "", "Sapi5VoiceImpl");
                Registry.SetValue(inprocServer32Path, "", @"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\mscoree.dll");
                Registry.SetValue(inprocServer32Path, "ThreadingModel", "Both");

                Console.WriteLine($"Registered voice '{model.Name}' successfully with SAPI5 and COM.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error registering voice '{model.Name}': {ex.Message}");
            }
        }

        public void UnregisterVoice(string voiceId)
        {
            string voiceRegistryPath = $@"{RegistryBasePath}\{voiceId}";
            try
            {
                using (var key = Registry.LocalMachine.OpenSubKey(voiceRegistryPath, true))
                {
                    if (key != null)
                    {
                        Registry.LocalMachine.DeleteSubKeyTree(voiceRegistryPath);
                        Console.WriteLine($"Successfully unregistered voice: {voiceId}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error unregistering voice {voiceId}: {ex.Message}");
            }
        }
    }
}
