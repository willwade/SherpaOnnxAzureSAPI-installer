using System;
using System.Linq;
using System.Collections.Generic;
using Microsoft.Win32;
using System.Globalization;
using System.IO;

namespace Installer
{
    public class Sapi5RegistrarExtended
    {
        private const string RegistryBasePath = @"SOFTWARE\Microsoft\SPEECH\Voices\Tokens";
        
        // GUID for Sherpa ONNX implementation
        private const string SherpaClsid = "{3d8f5c5d-9d6b-4b92-a12b-1a6dff80b6b2}";
        
        // GUID for Azure TTS implementation
        private const string AzureClsid = "{3d8f5c5e-9d6b-4b92-a12b-1a6dff80b6b3}";

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

        // Register a Sherpa ONNX voice
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
                    // 1. Register the SAPI voice
                    voiceKey.SetValue("", model.Name);
                    voiceKey.SetValue(lcid, model.Name); // Register for the specific language
                    voiceKey.SetValue("CLSID", SherpaClsid);
                    voiceKey.SetValue("Path", dllPath);

                    // 2. Set voice attributes
                    using (var attributesKey = voiceKey.CreateSubKey("Attributes"))
                    {
                        attributesKey.SetValue("Language", lcid);
                        attributesKey.SetValue("Gender", GetGenderAttribute(model.Gender));
                        attributesKey.SetValue("Age", "Adult");
                        attributesKey.SetValue("Vendor", model.Developer);
                        attributesKey.SetValue("Version", "1.0");
                        attributesKey.SetValue("Name", model.Name);

                        // 3. Set model paths with consistent naming
                        attributesKey.SetValue("Model Path", model.ModelPath);
                        attributesKey.SetValue("Tokens Path", model.TokensPath);
                        
                        // Add data directory path
                        string dataDir = Path.GetDirectoryName(model.ModelPath);
                        attributesKey.SetValue("Data Directory", dataDir);
                    }
                }

                // 4. Register the CLSID token for the voice
                using (var clsidKey = Registry.ClassesRoot.CreateSubKey($@"CLSID\{SherpaClsid}"))
                {
                    using (var tokenKey = clsidKey.CreateSubKey("Token"))
                    {
                        tokenKey.SetValue("", model.Name);
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

        // Register an Azure TTS voice
        public void RegisterAzureVoice(AzureTtsModel model, string dllPath, string subscriptionKey)
        {
            string voiceRegistryPath = $@"{RegistryBasePath}\{model.Name}";

            try
            {
                // Get LCID from the first language in the model
                var language = model.Language.FirstOrDefault();
                string lcid = GetLcidFromLanguage(language);

                Console.WriteLine($"Registering Azure TTS voice with LCID: {lcid}");

                using (var voiceKey = Registry.LocalMachine.CreateSubKey(voiceRegistryPath))
                {
                    // 1. Register the SAPI voice
                    voiceKey.SetValue("", model.Name);
                    voiceKey.SetValue(lcid, model.Name); // Register for the specific language
                    voiceKey.SetValue("CLSID", AzureClsid);
                    voiceKey.SetValue("Path", dllPath);

                    // 2. Set voice attributes
                    using (var attributesKey = voiceKey.CreateSubKey("Attributes"))
                    {
                        attributesKey.SetValue("Language", lcid);
                        attributesKey.SetValue("Gender", GetGenderAttribute(model.Gender));
                        attributesKey.SetValue("Age", "Adult");
                        attributesKey.SetValue("Vendor", model.Developer);
                        attributesKey.SetValue("Version", "1.0");
                        attributesKey.SetValue("Name", model.Name);

                        // 3. Set Azure-specific attributes
                        attributesKey.SetValue("SubscriptionKey", subscriptionKey);
                        attributesKey.SetValue("Region", model.Region);
                        attributesKey.SetValue("VoiceName", model.VoiceName);
                        
                        // Store style and role information if available
                        if (model.StyleList != null && model.StyleList.Count > 0)
                        {
                            attributesKey.SetValue("StyleList", string.Join(",", model.StyleList));
                        }
                        
                        if (model.RoleList != null && model.RoleList.Count > 0)
                        {
                            attributesKey.SetValue("RoleList", string.Join(",", model.RoleList));
                        }
                    }
                }

                // 4. Register the CLSID token for the voice
                using (var clsidKey = Registry.ClassesRoot.CreateSubKey($@"CLSID\{AzureClsid}"))
                {
                    using (var tokenKey = clsidKey.CreateSubKey("Token"))
                    {
                        tokenKey.SetValue("", model.Name);
                    }
                }

                Console.WriteLine($"Registered Azure TTS voice '{model.Name}' successfully with SAPI5 and COM.");
                Console.WriteLine($"Language ID: {lcid}");
                Console.WriteLine($"Region: {model.Region}");
                Console.WriteLine($"Voice Name: {model.VoiceName}");
                Console.WriteLine($"DLL path: {dllPath}");
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to register Azure TTS voice: {ex.Message}", ex);
            }
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
