using System;
using System.Speech.Synthesis;
using System.Linq;

namespace SimpleTest
{
    internal class Program
    {
        static void Main(string[] args)
        {
            try
            {
                using (var synth = new SpeechSynthesizer())
                {
                    // List all available voices
                    Console.WriteLine("Available voices:");
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
            Console.ReadKey();
        }
    }
}
