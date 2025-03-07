using System;
using System.Speech.Synthesis;
using System.Linq;
using Microsoft.Win32;
using System.IO;
using System.Threading;
using System.Collections.Generic;
using System.Diagnostics;

namespace SimpleTest
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("OpenSpeech TTS Test Application");
            Console.WriteLine("==============================");
            
            try
            {
                // First, list all available voices
                using (var synth = new SpeechSynthesizer())
                {
                    Console.WriteLine("\nAvailable voices:");
                    var installedVoices = synth.GetInstalledVoices().ToList();
                    
                    if (installedVoices.Count == 0)
                    {
                        Console.WriteLine("No voices found. Please install voices first.");
                        return;
                    }
                    
                    // Group voices by type
                    var sherpaVoices = new List<InstalledVoice>();
                    var azureVoices = new List<InstalledVoice>();
                    var otherVoices = new List<InstalledVoice>();
                    
                    foreach (var voice in installedVoices)
                    {
                        var info = voice.VoiceInfo;
                        string voiceType = GetVoiceType(info.Name);
                        
                        Console.WriteLine($"- {info.Name} ({voiceType}, {info.Gender}, {info.Age}, {info.Culture})");
                        
                        switch (voiceType)
                        {
                            case "Sherpa ONNX":
                                sherpaVoices.Add(voice);
                                break;
                            case "Azure TTS":
                                azureVoices.Add(voice);
                                break;
                            default:
                                otherVoices.Add(voice);
                                break;
                        }
                    }
                    
                    Console.WriteLine($"\nFound {sherpaVoices.Count} Sherpa ONNX voices, {azureVoices.Count} Azure TTS voices, and {otherVoices.Count} other voices.");
                    
                    // Test Sherpa ONNX voice if available
                    if (sherpaVoices.Count > 0)
                    {
                        Console.WriteLine("\n=== Testing Sherpa ONNX Voice ===");
                        TestVoice(synth, sherpaVoices[0], "This is a test of the Sherpa ONNX text-to-speech system.", "sherpa_output.wav");
                    }
                    
                    // Test Azure TTS voice if available
                    if (azureVoices.Count > 0)
                    {
                        Console.WriteLine("\n=== Testing Azure TTS Voice ===");
                        TestVoice(synth, azureVoices[0], "This is a test of the Azure text-to-speech system.", "azure_output.wav");
                        
                        // Try direct test with a standard voice first to verify SAPI is working
                        if (otherVoices.Count > 0)
                        {
                            Console.WriteLine("\n=== Testing Standard Voice for Comparison ===");
                            TestVoice(synth, otherVoices[0], "This is a test of a standard voice.", "standard_output.wav");
                        }
                        
                        // Try alternative approach for Azure voice
                        Console.WriteLine("\n=== Testing Azure TTS Voice (Alternative Method) ===");
                        TestAzureVoiceAlternative(azureVoices[0].VoiceInfo.Name);
                        
                        // Test Azure voice with style and role if available
                        TestAzureVoiceWithStyleAndRole(azureVoices[0].VoiceInfo.Name);
                    }
                    
                    Console.WriteLine("\nTest completed successfully!");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\nError: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
            
            Console.WriteLine("\nPress any key to exit...");
            Console.ReadKey();
        }
        
        static string GetVoiceType(string voiceName)
        {
            try
            {
                string registryPath = $@"SOFTWARE\Microsoft\SPEECH\Voices\Tokens\{voiceName}\Attributes";
                using (var key = Registry.LocalMachine.OpenSubKey(registryPath))
                {
                    if (key != null)
                    {
                        var voiceType = key.GetValue("VoiceType") as string;
                        if (voiceType != null)
                        {
                            if (voiceType == "SherpaOnnx")
                                return "Sherpa ONNX";
                            else if (voiceType == "AzureTTS")
                                return "Azure TTS";
                        }
                        
                        // Check for CLSID to determine type
                        using (var parentKey = Registry.LocalMachine.OpenSubKey($@"SOFTWARE\Microsoft\SPEECH\Voices\Tokens\{voiceName}"))
                        {
                            if (parentKey != null)
                            {
                                var clsid = parentKey.GetValue("CLSID") as string;
                                if (clsid == "{3d8f5c5d-9d6b-4b92-a12b-1a6dff80b6b2}")
                                    return "Sherpa ONNX";
                                else if (clsid == "{3d8f5c5e-9d6b-4b92-a12b-1a6dff80b6b3}")
                                    return "Azure TTS";
                            }
                        }
                    }
                }
                
                return "Standard";
            }
            catch
            {
                return "Unknown";
            }
        }
        
        static void TestVoice(SpeechSynthesizer synth, InstalledVoice voice, string testText, string outputFileName)
        {
            try
            {
                var info = voice.VoiceInfo;
                Console.WriteLine($"Testing voice: {info.Name}");
                
                // Print registry information for debugging
                string voiceType = GetVoiceType(info.Name);
                if (voiceType == "Sherpa ONNX" || voiceType == "Azure TTS")
                {
                    PrintRegistryInfo(info.Name);
                }
                
                try
                {
                    // For Azure voices, use SSML instead of SelectVoice
                    if (voiceType == "Azure TTS")
                    {
                        // Set output to audio file
                        string outputPath = Path.Combine(Environment.CurrentDirectory, outputFileName);
                        synth.SetOutputToWaveFile(outputPath);
                        
                        // Create SSML for Azure voice
                        string ssml = $@"<speak version='1.0' xmlns='http://www.w3.org/2001/10/synthesis' xml:lang='en-US'>
                            <voice name='{info.Name}'>
                                {testText}
                            </voice>
                        </speak>";
                        
                        // Speak using SSML
                        Console.WriteLine($"Speaking with SSML: \"{testText}\"");
                        synth.SpeakSsml(ssml);
                        
                        // Reset output to default
                        synth.SetOutputToDefaultAudioDevice();
                        
                        Console.WriteLine($"Speech saved to: {outputPath}");
                        
                        // Try speaking to the default audio device
                        Console.WriteLine("Now testing speech to default audio device with SSML...");
                        ssml = $@"<speak version='1.0' xmlns='http://www.w3.org/2001/10/synthesis' xml:lang='en-US'>
                            <voice name='{info.Name}'>
                                Hello, this is {info.Name}.
                            </voice>
                        </speak>";
                        synth.SpeakSsml(ssml);
                    }
                    else
                    {
                        // For non-Azure voices, use the standard approach
                        synth.SelectVoice(info.Name);
                        Console.WriteLine("Successfully selected the voice!");
                        
                        // Set output to audio file
                        string outputPath = Path.Combine(Environment.CurrentDirectory, outputFileName);
                        synth.SetOutputToWaveFile(outputPath);
                        
                        // Speak some text
                        Console.WriteLine($"Speaking: \"{testText}\"");
                        synth.Speak(testText);
                        
                        // Reset output to default
                        synth.SetOutputToDefaultAudioDevice();
                        
                        Console.WriteLine($"Speech saved to: {outputPath}");
                        
                        // Try speaking to the default audio device
                        Console.WriteLine("Now testing speech to default audio device...");
                        Console.WriteLine($"Speaking: \"Hello, this is {info.Name}.\"");
                        synth.Speak($"Hello, this is {info.Name}.");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error selecting or using voice: {ex.Message}");
                    Console.WriteLine($"Stack trace: {ex.StackTrace}");
                    
                    // Try to get more information about the voice
                    Console.WriteLine($"Voice info: ID={info.Id}, Name={info.Name}, Culture={info.Culture}, Gender={info.Gender}, Age={info.Age}, Enabled={voice.Enabled}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error testing voice {voice.VoiceInfo.Name}: {ex.Message}");
                Console.WriteLine($"Stack trace: {ex.StackTrace}");
            }
        }
        
        static void PrintRegistryInfo(string voiceName)
        {
            try
            {
                Console.WriteLine("Registry values read:");
                string registryPath = $@"SOFTWARE\Microsoft\SPEECH\Voices\Tokens\{voiceName}\Attributes";
                using (var key = Registry.LocalMachine.OpenSubKey(registryPath))
                {
                    if (key != null)
                    {
                        foreach (var valueName in key.GetValueNames())
                        {
                            var value = key.GetValue(valueName);
                            Console.WriteLine($"  {valueName}: {value}");
                            
                            // For Sherpa ONNX, check if files exist
                            if (valueName == "Model Path" || valueName == "Tokens Path" || valueName == "Data Directory")
                            {
                                string path = value.ToString();
                                bool exists = File.Exists(path) || Directory.Exists(path);
                                Console.WriteLine($"  {valueName} exists: {exists}");
                            }
                            
                            // For Azure TTS, check subscription key and region
                            if (valueName == "SubscriptionKey" || valueName == "Region" || valueName == "VoiceName")
                            {
                                Console.WriteLine($"  {valueName}: {value}");
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading registry: {ex.Message}");
            }
        }
        
        static void TestAzureVoiceAlternative(string voiceName)
        {
            try
            {
                Console.WriteLine($"Testing Azure voice using alternative method: {voiceName}");
                
                // Get Azure voice details from registry
                string subscriptionKey = null;
                string region = null;
                string azureVoiceName = null;
                
                string registryPath = $@"SOFTWARE\Microsoft\SPEECH\Voices\Tokens\{voiceName}\Attributes";
                using (var key = Registry.LocalMachine.OpenSubKey(registryPath))
                {
                    if (key != null)
                    {
                        subscriptionKey = key.GetValue("SubscriptionKey") as string;
                        region = key.GetValue("Region") as string;
                        azureVoiceName = key.GetValue("VoiceName") as string;
                    }
                }
                
                if (string.IsNullOrEmpty(subscriptionKey) || string.IsNullOrEmpty(region) || string.IsNullOrEmpty(azureVoiceName))
                {
                    Console.WriteLine("Missing required Azure voice information in registry.");
                    return;
                }
                
                Console.WriteLine($"Using Azure voice: {azureVoiceName}, Region: {region}");
                
                // Create a process to run a simple test using the OpenSpeechTTS.dll directly
                string dllPath = @"C:\Program Files\OpenAssistive\OpenSpeech\OpenSpeechTTS.dll";
                if (!File.Exists(dllPath))
                {
                    Console.WriteLine($"DLL not found at path: {dllPath}");
                    return;
                }
                
                Console.WriteLine("DLL exists, attempting to use it directly...");
                
                // Try to use the voice through SAPI directly
                using (var synth = new SpeechSynthesizer())
                {
                    // Try with SSML
                    string ssml = $@"<speak version='1.0' xmlns='http://www.w3.org/2001/10/synthesis' xml:lang='en-US'>
                        <voice name='{voiceName}'>
                            This is a direct test of the Azure voice.
                        </voice>
                    </speak>";
                    
                    try
                    {
                        Console.WriteLine("Testing with SSML...");
                        synth.SpeakSsml(ssml);
                        Console.WriteLine("SSML test completed successfully!");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error with SSML test: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in alternative Azure voice test: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
        
        static void TestAzureVoiceWithStyleAndRole(string voiceName)
        {
            try
            {
                string registryPath = $@"SOFTWARE\Microsoft\SPEECH\Voices\Tokens\{voiceName}\Attributes";
                using (var key = Registry.LocalMachine.OpenSubKey(registryPath))
                {
                    if (key != null)
                    {
                        var styleList = key.GetValue("StyleList") as string;
                        var roleList = key.GetValue("RoleList") as string;
                        
                        if (!string.IsNullOrEmpty(styleList))
                        {
                            Console.WriteLine($"Voice {voiceName} supports styles: {styleList}");
                            
                            // Test with first style
                            var styles = styleList.Split(',');
                            if (styles.Length > 0)
                            {
                                using (var synth = new SpeechSynthesizer())
                                {
                                    try
                                    {
                                        // Don't use SelectVoice, just use SSML directly
                                        string style = styles[0].Trim();
                                        string ssml = $@"<speak version='1.0' xmlns='http://www.w3.org/2001/10/synthesis' xmlns:mstts='https://www.w3.org/2001/mstts' xml:lang='en-US'>
                                            <voice name='{voiceName}'>
                                                <mstts:express-as style='{style}'>
                                                    This is a test of the Azure voice with the {style} style.
                                                </mstts:express-as>
                                            </voice>
                                        </speak>";
                                        
                                        Console.WriteLine($"Testing Azure voice with style: {style}");
                                        synth.SpeakSsml(ssml);
                                        Console.WriteLine("Style test completed successfully!");
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine($"Error testing style: {ex.Message}");
                                    }
                                }
                            }
                        }
                        
                        if (!string.IsNullOrEmpty(roleList))
                        {
                            Console.WriteLine($"Voice {voiceName} supports roles: {roleList}");
                            
                            // Test with first role
                            var roles = roleList.Split(',');
                            if (roles.Length > 0)
                            {
                                using (var synth = new SpeechSynthesizer())
                                {
                                    try
                                    {
                                        // Don't use SelectVoice, just use SSML directly
                                        string role = roles[0].Trim();
                                        string ssml = $@"<speak version='1.0' xmlns='http://www.w3.org/2001/10/synthesis' xmlns:mstts='https://www.w3.org/2001/mstts' xml:lang='en-US'>
                                            <voice name='{voiceName}'>
                                                <mstts:express-as role='{role}'>
                                                    This is a test of the Azure voice with the {role} role.
                                                </mstts:express-as>
                                            </voice>
                                        </speak>";
                                        
                                        Console.WriteLine($"Testing Azure voice with role: {role}");
                                        synth.SpeakSsml(ssml);
                                        Console.WriteLine("Role test completed successfully!");
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine($"Error testing role: {ex.Message}");
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error testing Azure voice with style/role: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
