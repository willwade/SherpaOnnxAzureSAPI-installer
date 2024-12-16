using Microsoft.Win32;

public class Sapi5Registrar
{
    public void RegisterVoice(TtsModel model, string dllPath)
    {
        string registryPath = $@"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\SPEECH\Voices\Tokens\{model.Id}";
        int lcid = LanguageCodeConverter.ConvertToLcid(model.Language[0].LangCode, model.Language[0].Country);

        Registry.SetValue(registryPath, "", model.Name);
        Registry.SetValue(registryPath, "Lang", lcid.ToString("X"));
        Registry.SetValue(registryPath, "Gender", "Neutral");
        Registry.SetValue(registryPath, "Age", "Adult");
        Registry.SetValue(registryPath, "Vendor", model.Developer);
        Registry.SetValue(registryPath, "Version", "1.0");
        Registry.SetValue($@"{registryPath}\Attributes", "CLSID", "{Your-CLSID-GUID}");
        Registry.SetValue($@"{registryPath}\Attributes", "VoicePath", dllPath);

        Console.WriteLine($"Registered {model.Name} with SAPI5.");
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
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error unregistering voice {voiceId}: {ex.Message}");
        }
    }
}
