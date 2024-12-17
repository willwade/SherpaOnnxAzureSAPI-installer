using System;
using Microsoft.Win32;
using Installer.Shared;

public class Sapi5Registrar
{
    public void RegisterVoice(TtsModel model, string dllPath)
    {
        string registryBasePath = $@"SOFTWARE\Microsoft\SPEECH\Voices\Tokens\{model.Id}";
        int lcid = LanguageCodeConverter.ConvertToLcid(model.Language[0].LangCode, model.Language[0].Country);

        try
        {
            // Set main registry values for the voice
            Registry.SetValue($@"HKEY_LOCAL_MACHINE\{registryBasePath}", "", model.Name);
            Registry.SetValue($@"HKEY_LOCAL_MACHINE\{registryBasePath}", "Lang", lcid.ToString("X"));
            Registry.SetValue($@"HKEY_LOCAL_MACHINE\{registryBasePath}", "Gender", "Neutral");
            Registry.SetValue($@"HKEY_LOCAL_MACHINE\{registryBasePath}", "Age", "Adult");
            Registry.SetValue($@"HKEY_LOCAL_MACHINE\{registryBasePath}", "Vendor", model.Developer);
            Registry.SetValue($@"HKEY_LOCAL_MACHINE\{registryBasePath}", "Version", "1.0");

            // Add required 'Attributes' subkey with CLSID and VoicePath
            string attributesPath = $@"HKEY_LOCAL_MACHINE\{registryBasePath}\Attributes";
            Registry.SetValue(attributesPath, "CLSID", "3d8f5c5d-9d6b-4b92-a12b-1a6dff80b6b2"); 
            Registry.SetValue(attributesPath, "VoicePath", dllPath);

            Console.WriteLine($"Registered voice '{model.Name}' with SAPI5.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error registering voice '{model.Name}': {ex.Message}");
        }
    }

    public void UnregisterVoice(string voiceId)
    {
        string registryPath = $@"SOFTWARE\Microsoft\SPEECH\Voices\Tokens\{voiceId}";
        try
        {
            using (var key = Registry.LocalMachine.OpenSubKey(registryPath, true))
            {
                if (key != null)
                {
                    Registry.LocalMachine.DeleteSubKeyTree(registryPath);
                    Console.WriteLine($"Successfully unregistered voice: {voiceId}");
                }
                else
                {
                    Console.WriteLine($"Voice '{voiceId}' not found in the registry.");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error unregistering voice '{voiceId}': {ex.Message}");
        }
    }
}
