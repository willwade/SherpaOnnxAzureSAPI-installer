using System;
using System.Speech.Synthesis;
using System.Linq;
using Microsoft.Win32;
using System.IO;
using System.Threading;
using System.Runtime.InteropServices;

namespace DirectTest
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Direct Voice Test Application");
            Console.WriteLine("=============================");
            
            // List all available voices
            using (var synth = new SpeechSynthesizer())
            {
                Console.WriteLine("\nAvailable voices:");
                foreach (var voice in synth.GetInstalledVoices())
                {
                    var info = voice.VoiceInfo;
                    Console.WriteLine($"- {info.Name} ({info.Gender}, {info.Age}, {info.Culture})");
                }
                
                // Test each voice that contains "joe" or "amy"
                var testVoices = synth.GetInstalledVoices()
                    .Where(v => v.VoiceInfo.Name.ToLower().Contains("joe") || 
                                v.VoiceInfo.Name.ToLower().Contains("amy"))
                    .ToList();
                
                if (testVoices.Count > 0)
                {
                    Console.WriteLine("\nTesting voices:");
                    
                    foreach (var voice in testVoices)
                    {
                        Console.WriteLine($"\nTesting voice: {voice.VoiceInfo.Name}");
                        
                        try
                        {
                            // Check registry for this voice
                            CheckRegistryForVoice(voice.VoiceInfo.Name);
                            
                            // Try to select and use the voice
                            synth.SelectVoice(voice.VoiceInfo.Name);
                            Console.WriteLine($"Successfully selected voice: {voice.VoiceInfo.Name}");
                            
                            // Test speaking
                            string testText = $"This is a test of the {voice.VoiceInfo.Name} voice.";
                            Console.WriteLine($"Speaking: \"{testText}\"");
                            
                            // Set output to a unique file for each voice
                            string outputPath = Path.Combine(Environment.CurrentDirectory, $"test_{voice.VoiceInfo.Name.Replace(" ", "_")}.wav");
                            synth.SetOutputToWaveFile(outputPath);
                            synth.Speak(testText);
                            synth.SetOutputToNull();
                            
                            Console.WriteLine($"Speech saved to: {outputPath}");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error with voice {voice.VoiceInfo.Name}: {ex.Message}");
                            Console.WriteLine($"Stack trace: {ex.StackTrace}");
                            
                            if (ex.InnerException != null)
                            {
                                Console.WriteLine($"Inner exception: {ex.InnerException.Message}");
                                Console.WriteLine($"Inner stack trace: {ex.InnerException.StackTrace}");
                            }
                        }
                    }
                }
                else
                {
                    Console.WriteLine("\nNo test voices found.");
                }
            }
            
            Console.WriteLine("\nPress any key to exit...");
            Console.ReadKey();
        }
        
        static void CheckRegistryForVoice(string voiceName)
        {
            Console.WriteLine($"Checking registry for voice: {voiceName}");
            
            try
            {
                // First try to find the voice token
                string? tokenKeyName = null;
                
                using (var voicesKey = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Speech\Voices\Tokens"))
                {
                    if (voicesKey != null)
                    {
                        foreach (var subKeyName in voicesKey.GetSubKeyNames())
                        {
                            using (var subKey = voicesKey.OpenSubKey(subKeyName))
                            {
                                if (subKey != null)
                                {
                                    var defaultValue = subKey.GetValue("");
                                    if (defaultValue != null && defaultValue.ToString() == voiceName)
                                    {
                                        tokenKeyName = subKeyName;
                                        Console.WriteLine($"Found voice token key: {tokenKeyName}");
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
                
                // If we found the token, check its attributes
                if (tokenKeyName != null)
                {
                    using (var tokenKey = Registry.LocalMachine.OpenSubKey($@"SOFTWARE\Microsoft\Speech\Voices\Tokens\{tokenKeyName}"))
                    {
                        if (tokenKey != null)
                        {
                            Console.WriteLine("Token key values:");
                            foreach (var valueName in tokenKey.GetValueNames())
                            {
                                Console.WriteLine($"  {valueName}: {tokenKey.GetValue(valueName)}");
                            }
                            
                            // Check for CLSID
                            var clsid = tokenKey.GetValue("CLSID") as string;
                            if (!string.IsNullOrEmpty(clsid))
                            {
                                Console.WriteLine($"CLSID: {clsid}");
                                
                                // Check if this CLSID exists in the registry
                                using (var clsidKey = Registry.ClassesRoot.OpenSubKey($@"CLSID\{clsid}"))
                                {
                                    if (clsidKey != null)
                                    {
                                        Console.WriteLine("CLSID exists in registry");
                                        
                                        // Check InprocServer32
                                        using (var inprocKey = clsidKey.OpenSubKey("InprocServer32"))
                                        {
                                            if (inprocKey != null)
                                            {
                                                var dllPath = inprocKey.GetValue("") as string;
                                                Console.WriteLine($"DLL path: {dllPath}");
                                                
                                                if (!string.IsNullOrEmpty(dllPath) && File.Exists(dllPath))
                                                {
                                                    Console.WriteLine("DLL file exists");
                                                }
                                                else
                                                {
                                                    Console.WriteLine("DLL file does not exist");
                                                }
                                            }
                                            else
                                            {
                                                Console.WriteLine("InprocServer32 key not found");
                                            }
                                        }
                                    }
                                    else
                                    {
                                        Console.WriteLine("CLSID not found in registry");
                                    }
                                }
                            }
                        }
                    }
                    
                    // Check attributes
                    using (var attributesKey = Registry.LocalMachine.OpenSubKey($@"SOFTWARE\Microsoft\Speech\Voices\Tokens\{tokenKeyName}\Attributes"))
                    {
                        if (attributesKey != null)
                        {
                            Console.WriteLine("\nAttributes:");
                            foreach (var valueName in attributesKey.GetValueNames())
                            {
                                Console.WriteLine($"  {valueName}: {attributesKey.GetValue(valueName)}");
                            }
                        }
                    }
                }
                else
                {
                    Console.WriteLine($"Voice token not found for: {voiceName}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error checking registry: {ex.Message}");
            }
        }
    }
}
