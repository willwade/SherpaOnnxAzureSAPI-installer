using System;
using Microsoft.Win32;
using System.IO;
using System.Globalization;
using System.Linq;
using Installer.Shared;

namespace Installer
{
    public class Sapi5Registrar
    {
        private const string RegistryBasePath = @"SOFTWARE\Microsoft\SPEECH\Voices\Tokens";

        private string GetLcidFromLanguage(LanguageInfo language)
        {
            try
            {
                if (language == null) return "409"; // Default to US English

                // Combine language code and country for full culture code
                string cultureName = $"{language.LangCode}-{language.Country}";
                var culture = CultureInfo.GetCultures(CultureTypes.AllCultures)
                    .FirstOrDefault(c => c.Name.Equals(cultureName, StringComparison.OrdinalIgnoreCase));

                if (culture != null)
                {
                    return culture.LCID.ToString();
                }

                // Fallback to just language code if country specific culture not found
                culture = CultureInfo.GetCultures(CultureTypes.AllCultures)
                    .FirstOrDefault(c => c.TwoLetterISOLanguageName.Equals(language.LangCode, StringComparison.OrdinalIgnoreCase));

                if (culture != null)
                {
                    return culture.LCID.ToString();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error getting LCID: {ex.Message}");
            }

            return "409"; // Default to US English if not found
        }

        public void RegisterVoice(TtsModel model, string dllPath)
        {
            // Use model.Name instead of model.Id for the registry key
            string voiceRegistryPath = $@"{RegistryBasePath}\{model.Name}";
            string clsid = "3d8f5c5d-9d6b-4b92-a12b-1a6dff80b6b2";
            string clsidPath = $@"HKEY_CLASSES_ROOT\CLSID\{clsid}";
            string inprocServer32Path = $@"HKEY_CLASSES_ROOT\CLSID\{clsid}\InprocServer32";

            try
            {
                // Get LCID from the first language in the model
                string lcid = GetLcidFromLanguage(model.Language.FirstOrDefault());

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

        public void UnregisterVoice(string voiceId, string voiceName)
        {
            string voiceRegistryPath = $@"{RegistryBasePath}\{voiceName}";
            try
            {
                using (var key = Registry.LocalMachine.OpenSubKey(voiceRegistryPath, true))
                {
                    if (key != null)
                    {
                        Registry.LocalMachine.DeleteSubKeyTree(voiceRegistryPath);
                        Console.WriteLine($"Unregistered voice from registry: {voiceName}");
                    }
                    else
                    {
                        Console.WriteLine($"Voice not found in registry: {voiceName}");
                    }
                }

                // Also try to clean up model files
                string modelDir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                    "OpenSpeech",
                    "models",
                    voiceId
                );

                if (Directory.Exists(modelDir))
                {
                    Directory.Delete(modelDir, true);
                    Console.WriteLine($"Deleted model directory: {modelDir}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error unregistering voice '{voiceName}': {ex.Message}");
                throw;
            }
        }
    }
}
