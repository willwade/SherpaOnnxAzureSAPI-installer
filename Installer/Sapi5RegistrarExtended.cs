using System;
using System.Linq;
using System.Collections.Generic;
using Microsoft.Win32;
using System.Globalization;
using System.IO;
using Installer.Shared;

namespace Installer
{
    public class Sapi5RegistrarExtended
    {
        private const string RegistryBasePath = @"SOFTWARE\Microsoft\SPEECH\Voices\Tokens";

        // CLSID for the SAPI5 TTS Engine COM object (coclass)
        // See SAPI5_CLSID_TRACKING.md for details
        private const string TTS_ENGINE_CLSID = "{A1B2C3D4-E5F6-4A5B-8C9D-1E2F3A4B5C6D}";

        private string GetLcidFromLanguage(LanguageInfo language)
        {
            try
            {
                if (language == null) return "0409"; // Default to en-US if no language specified

                // Special handling for British English
                if (language.LangCode?.Equals("en", StringComparison.OrdinalIgnoreCase) == true &&
                    (language.Country?.Equals("GB", StringComparison.OrdinalIgnoreCase) == true ||
                     language.Country?.Equals("UK", StringComparison.OrdinalIgnoreCase) == true))
                {
                    return "0809"; // British English
                }

                // Create culture name from language code and country
                string cultureName = !string.IsNullOrEmpty(language.Country) 
                    ? $"{language.LangCode}-{language.Country}"
                    : language.LangCode;

                try
                {
                    var culture = CultureInfo.GetCultureInfo(cultureName);
                    return culture.LCID.ToString("X4"); // Format as 4-digit hex
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: Could not find culture for {cultureName}: {ex.Message}");
                    
                    // Try just the language code if country-specific lookup failed
                    try
                    {
                        var culture = CultureInfo.GetCultureInfo(language.LangCode);
                        return culture.LCID.ToString("X4");
                    }
                    catch
                    {
                        Console.WriteLine($"Warning: Could not find culture for language code {language.LangCode}");
                        return "0409"; // Default to en-US if all lookups fail
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error getting LCID: {ex.Message}");
                return "0409"; // Default to en-US if any error occurs
            }
        }

        private string GetLcidFromLocale(string locale)
        {
            try
            {
                if (string.IsNullOrEmpty(locale)) return "0409"; // Default to en-US if no locale specified

                try
                {
                    var culture = CultureInfo.GetCultureInfo(locale);
                    return culture.LCID.ToString("X4"); // Format as 4-digit hex
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: Could not find culture for {locale}: {ex.Message}");
                    return "0409"; // Default to en-US if lookup fails
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error getting LCID: {ex.Message}");
                return "0409"; // Default to en-US if any error occurs
            }
        }

        private string GetGenderAttribute(string gender)
        {
            if (string.IsNullOrEmpty(gender))
                return "Male"; // Default to Male if not specified

            // Handle common variations
            switch (gender.Trim().ToLowerInvariant())
            {
                case "f":
                case "female":
                    return "Female";
                case "m":
                case "male":
                    return "Male";
                case "n":
                case "neutral":
                    return "Neutral";
                default:
                    return "Male"; // Default to Male for unknown values
            }
        }

        public void RegisterSherpaVoice(TtsModel model, string dllPath)
        {
            string voiceRegistryPath = $@"{RegistryBasePath}\{model.Name}";

            try
            {
                // Get LCID from the first language in the model
                var language = model.Language.FirstOrDefault();
                string lcid = GetLcidFromLanguage(language);

                Console.WriteLine($"Registering Sherpa ONNX voice with LCID: {lcid}");

                using (var voiceKey = Registry.LocalMachine.CreateSubKey(voiceRegistryPath))
                {
                    // 1. Register the SAPI voice token - all voices use the same CLSID
                    voiceKey.SetValue("", model.Name);
                    voiceKey.SetValue(lcid, model.Name); // Register for the specific language
                    voiceKey.SetValue("CLSID", TTS_ENGINE_CLSID);

                    // 2. Set voice attributes
                    using (var attributesKey = voiceKey.CreateSubKey("Attributes"))
                    {
                        attributesKey.SetValue("Language", lcid);
                        attributesKey.SetValue("Gender", GetGenderAttribute(model.Gender));
                        attributesKey.SetValue("Age", "Adult");
                        attributesKey.SetValue("Vendor", model.Developer);
                        attributesKey.SetValue("Version", "1.0");
                        attributesKey.SetValue("Name", model.Name);
                        attributesKey.SetValue("VoiceType", "SherpaOnnx");

                        // 3. Set model paths with consistent naming
                        attributesKey.SetValue("Model Path", model.ModelPath);
                        attributesKey.SetValue("Tokens Path", model.TokensPath);

                        // Add data directory path
                        string dataDir = Path.GetDirectoryName(model.ModelPath);
                        attributesKey.SetValue("Data Directory", dataDir);
                    }
                }

                Console.WriteLine($"Registered Sherpa ONNX voice '{model.Name}' successfully with SAPI5 and COM.");
                Console.WriteLine($"Language ID: {lcid}");
                Console.WriteLine($"Model path: {model.ModelPath}");
                Console.WriteLine($"Tokens path: {model.TokensPath}");
                Console.WriteLine($"DLL path: {dllPath}");
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to register Sherpa ONNX voice: {ex.Message}", ex);
            }
        }

        public void RegisterAzureVoice(AzureTtsModel model, string dllPath)
        {
            string voiceRegistryPath = $@"{RegistryBasePath}\{model.Name}";

            try
            {
                // Get LCID from the locale
                string lcid = GetLcidFromLocale(model.Locale);

                Console.WriteLine($"Registering Azure TTS voice with LCID: {lcid}");

                using (var voiceKey = Registry.LocalMachine.CreateSubKey(voiceRegistryPath))
                {
                    // 1. Register the SAPI voice token - all voices use the same CLSID
                    voiceKey.SetValue("", model.Name);
                    voiceKey.SetValue(lcid, model.Name); // Register for the specific language
                    voiceKey.SetValue("CLSID", TTS_ENGINE_CLSID);

                    // 2. Set voice attributes
                    using (var attributesKey = voiceKey.CreateSubKey("Attributes"))
                    {
                        attributesKey.SetValue("Language", lcid);
                        attributesKey.SetValue("Gender", GetGenderAttribute(model.Gender));
                        attributesKey.SetValue("Age", "Adult");
                        attributesKey.SetValue("Vendor", "Microsoft");
                        attributesKey.SetValue("Version", "1.0");
                        attributesKey.SetValue("Name", model.Name);
                        attributesKey.SetValue("VoiceType", "AzureTTS");

                        // 3. Set Azure-specific attributes
                        attributesKey.SetValue("SubscriptionKey", model.SubscriptionKey);
                        attributesKey.SetValue("Region", model.Region);
                        attributesKey.SetValue("VoiceName", model.ShortName);

                        // Set style and role lists if available
                        if (model.StyleList != null && model.StyleList.Count > 0)
                        {
                            attributesKey.SetValue("StyleList", string.Join(",", model.StyleList));
                        }

                        if (model.RoleList != null && model.RoleList.Count > 0)
                        {
                            attributesKey.SetValue("RoleList", string.Join(",", model.RoleList));
                        }

                        // Set selected style and role if specified
                        if (!string.IsNullOrEmpty(model.SelectedStyle))
                        {
                            attributesKey.SetValue("SelectedStyle", model.SelectedStyle);
                        }

                        if (!string.IsNullOrEmpty(model.SelectedRole))
                        {
                            attributesKey.SetValue("SelectedRole", model.SelectedRole);
                        }
                    }
                }

                Console.WriteLine($"Registered Azure TTS voice '{model.Name}' successfully with SAPI5 and COM.");
                Console.WriteLine($"Language ID: {lcid}");
                Console.WriteLine($"Azure Voice: {model.ShortName}");
                Console.WriteLine($"Region: {model.Region}");
                Console.WriteLine($"DLL path: {dllPath}");
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to register Azure TTS voice: {ex.Message}", ex);
            }
        }

        // Generic RegisterVoice method for TtsModel (SherpaOnnx voices)
        public void RegisterVoice(TtsModel model, string dllPath)
        {
            RegisterSherpaVoice(model, dllPath);
        }

        // Generic RegisterVoice method for AzureTtsModel (Azure voices)
        public void RegisterVoice(AzureTtsModel model, string dllPath)
        {
            RegisterAzureVoice(model, dllPath);
        }

        public void UnregisterVoice(string voiceId)
        {
            try
            {
                using (var voicesKey = Registry.LocalMachine.OpenSubKey(RegistryBasePath, true))
                {
                    if (voicesKey != null)
                    {
                        voicesKey.DeleteSubKeyTree(voiceId, false);
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
