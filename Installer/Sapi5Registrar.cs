using System;
using Microsoft.Win32;
using System.IO;
using System.Globalization;
using System.Linq;

namespace Installer
{
    public class Sapi5Registrar
    {
        private const string RegistryBasePath = @"SOFTWARE\Microsoft\SPEECH\Voices\Tokens";

        private string GetLcidFromLanguage(string langCode)
        {
            try
            {
                // Convert ISO language code to CultureInfo
                var culture = CultureInfo.GetCultures(CultureTypes.AllCultures)
                    .FirstOrDefault(c => c.TwoLetterISOLanguageName.Equals(langCode, StringComparison.OrdinalIgnoreCase));

                if (culture != null)
                {
                    return culture.LCID.ToString();
                }
            }
            catch { }

            // Default to English (US) if not found
            return "409";
        }

        public void RegisterVoice(TtsModel model, string dllPath)
        {
            string voiceRegistryPath = $@"{RegistryBasePath}\{model.Id}";
            string clsid = "3d8f5c5d-9d6b-4b92-a12b-1a6dff80b6b2";
            string clsidPath = $@"HKEY_CLASSES_ROOT\CLSID\{clsid}";
            string inprocServer32Path = $@"HKEY_CLASSES_ROOT\CLSID\{clsid}\InprocServer32";

            try
            {
                // Get LCID from the first language in the model
                string lcid = GetLcidFromLanguage(model.Language.FirstOrDefault()?.LangCode ?? "en");

                // 1. Register the SAPI voice
                Registry.SetValue($@"HKEY_LOCAL_MACHINE\{voiceRegistryPath}", "", model.Name);
                Registry.SetValue($@"HKEY_LOCAL_MACHINE\{voiceRegistryPath}", "Lang", lcid);
                Registry.SetValue($@"HKEY_LOCAL_MACHINE\{voiceRegistryPath}", "Gender", "Neutral");
                Registry.SetValue($@"HKEY_LOCAL_MACHINE\{voiceRegistryPath}", "Age", "Adult");
                Registry.SetValue($@"HKEY_LOCAL_MACHINE\{voiceRegistryPath}", "Vendor", model.Developer);
                Registry.SetValue($@"HKEY_LOCAL_MACHINE\{voiceRegistryPath}", "Version", "1.0");

                // Add Attributes subkey with CLSID, VoicePath, and model paths
                string attributesPath = $@"HKEY_LOCAL_MACHINE\{voiceRegistryPath}\Attributes";
                Registry.SetValue(attributesPath, "CLSID", clsid);
                Registry.SetValue(attributesPath, "VoicePath", dllPath);
                Registry.SetValue(attributesPath, "ModelPath", model.ModelPath);
                Registry.SetValue(attributesPath, "TokensPath", model.TokensPath);
                Registry.SetValue(attributesPath, "LexiconPath", model.LexiconPath ?? "");
                Registry.SetValue(attributesPath, "ModelType", model.ModelType ?? "vits");

                // 2. Register the COM class CLSID with InprocServer32
                Registry.SetValue(clsidPath, "", "Sapi5VoiceImpl");
                Registry.SetValue(inprocServer32Path, "", @"C:\Windows\Microsoft.NET\Framework64\v4.0.30319\mscoree.dll");
                Registry.SetValue(inprocServer32Path, "ThreadingModel", "Both");

                Console.WriteLine($"Registered voice '{model.Name}' successfully with SAPI5 and COM.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error registering voice '{model.Name}': {ex.Message}");
                throw;
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
                throw;
            }
        }
    }
}
