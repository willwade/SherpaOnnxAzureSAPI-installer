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
                Registry.SetValue(attributesPath, "VoicePath", Path.GetFullPath(dllPath)); // Use full path
                Registry.SetValue(attributesPath, "ModelPath", Path.GetFullPath(model.ModelPath)); // Use full path
                Registry.SetValue(attributesPath, "TokensPath", Path.GetFullPath(model.TokensPath)); // Use full path
                Registry.SetValue(attributesPath, "LexiconPath", model.LexiconPath ?? "");
                Registry.SetValue(attributesPath, "ModelType", model.ModelType ?? "vits");

                // 2. Register the COM class CLSID with InprocServer32
                Registry.SetValue(clsidPath, "", "OpenSpeechTTS.Sapi5VoiceImpl");
                Registry.SetValue(inprocServer32Path, "", @"mscoree.dll"); // Just use mscoree.dll, Windows will find it
                Registry.SetValue(inprocServer32Path, "ThreadingModel", "Both");
                Registry.SetValue(inprocServer32Path, "Class", "OpenSpeechTTS.Sapi5VoiceImpl");
                Registry.SetValue(inprocServer32Path, "Assembly", "OpenSpeechTTS, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null");
                Registry.SetValue(inprocServer32Path, "RuntimeVersion", "v4.0.30319");
                Registry.SetValue(inprocServer32Path, "CodeBase", $"file:///{Path.GetFullPath(dllPath)}"); // Use file:/// format

                // 3. Register the ProgID
                string progId = "OpenSpeechTTS.Sapi5VoiceImpl";
                Registry.SetValue($@"HKEY_CLASSES_ROOT\{progId}", "", "OpenSpeechTTS SAPI5 Voice Implementation");
                Registry.SetValue($@"HKEY_CLASSES_ROOT\{progId}\CLSID", "", clsid);
                Registry.SetValue(clsidPath, "ProgId", progId);

                Console.WriteLine($"Registered voice '{model.Name}' successfully with SAPI5 and COM.");
                Console.WriteLine($"Model path: {Path.GetFullPath(model.ModelPath)}");
                Console.WriteLine($"Tokens path: {Path.GetFullPath(model.TokensPath)}");
                Console.WriteLine($"DLL path: {Path.GetFullPath(dllPath)}");
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
