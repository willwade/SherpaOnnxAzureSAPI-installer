using System;
using System.Speech.Synthesis;
using System.Linq;
using Microsoft.Win32;
using System.IO;

namespace SimpleTest
{
    internal class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Test registry access first
                Console.WriteLine("Testing registry access...");
                string modelPath = null;
                string tokensPath = null;
                string dataDirPath = null;

                using (var key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Speech\Voices\Tokens\OpenSpeechAmyVoice"))
                {
                    if (key != null)
                    {
                        Console.WriteLine("Found registry key. Values:");
                        foreach (var valueName in key.GetValueNames())
                        {
                            var value = key.GetValue(valueName);
                            Console.WriteLine($"  {valueName}: {value}");
                            if (valueName == "ModelPath") modelPath = value as string;
                            if (valueName == "TokensPath") tokensPath = value as string;
                            if (valueName == "DataDirPath") dataDirPath = value as string;
                        }

                        Console.WriteLine("\nChecking if files exist:");
                        if (modelPath != null)
                        {
                            Console.WriteLine($"Model file exists: {File.Exists(modelPath)} at {modelPath}");
                        }
                        if (tokensPath != null)
                        {
                            Console.WriteLine($"Tokens file exists: {File.Exists(tokensPath)} at {tokensPath}");
                        }
                        if (dataDirPath != null)
                        {
                            Console.WriteLine($"Data directory exists: {Directory.Exists(dataDirPath)} at {dataDirPath}");
                            if (Directory.Exists(dataDirPath))
                            {
                                Console.WriteLine("Data directory contents:");
                                foreach (var file in Directory.GetFiles(dataDirPath))
                                {
                                    Console.WriteLine($"  {Path.GetFileName(file)}");
                                }
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine("Registry key not found!");
                    }
                }

                Console.WriteLine("\nAvailable voices:");
                using (var synth = new SpeechSynthesizer())
                {
                    foreach (var voice in synth.GetInstalledVoices())
                    {
                        var info = voice.VoiceInfo;
                        Console.WriteLine($"Name: {info.Name}");
                        Console.WriteLine($"  Culture: {info.Culture}");
                        Console.WriteLine($"  Age: {info.Age}");
                        Console.WriteLine($"  Gender: {info.Gender}");
                        Console.WriteLine($"  Description: {info.Description}");
                        Console.WriteLine();
                    }

                    // Try to find our voice
                    var amyVoice = synth.GetInstalledVoices()
                        .FirstOrDefault(v => v.VoiceInfo.Name.Equals("OpenSpeech Amy", StringComparison.OrdinalIgnoreCase));

                    if (amyVoice != null)
                    {
                        Console.WriteLine("Found Amy voice, testing speech...");
                        synth.SelectVoice(amyVoice.VoiceInfo.Name);
                        synth.Rate = 0; // Normal speed
                        synth.Volume = 100; // Full volume
                        synth.Speak("Hello! This is a test of the OpenSpeech Amy voice using SAPI 5.");
                    }
                    else
                    {
                        Console.WriteLine("Amy voice not found in installed voices!");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                Console.WriteLine($"Stack trace: {ex.StackTrace}");
            }

            Console.WriteLine("\nPress any key to exit...");
            try { Console.ReadKey(); } catch { }
        }
    }
}
