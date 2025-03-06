using System;
using System.Speech.Synthesis;
using System.Linq;
using Microsoft.Win32;
using System.IO;
using System.Threading;

namespace SimpleTest
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Sherpa ONNX TTS Test Application");
            Console.WriteLine("================================");
            
            try
            {
                // First, list all available voices
                using (var synth = new SpeechSynthesizer())
                {
                    Console.WriteLine("\nAvailable voices:");
                    foreach (var voice in synth.GetInstalledVoices())
                    {
                        var info = voice.VoiceInfo;
                        Console.WriteLine($"- {info.Name} ({info.Gender}, {info.Age}, {info.Culture})");
                    }
                    
                    // Try to find a Sherpa ONNX voice
                    var sherpaVoice = synth.GetInstalledVoices().FirstOrDefault(v => 
                        v.VoiceInfo.Name.Contains("sherpa", StringComparison.OrdinalIgnoreCase) || 
                        v.VoiceInfo.Name.Contains("openspeech", StringComparison.OrdinalIgnoreCase));
                    
                    if (sherpaVoice != null)
                    {
                        Console.WriteLine($"\nFound Sherpa ONNX voice: {sherpaVoice.VoiceInfo.Name}");
                        synth.SelectVoice(sherpaVoice.VoiceInfo.Name);
                    }
                    else
                    {
                        Console.WriteLine("\nNo Sherpa ONNX voice found. Using default voice.");
                    }
                    
                    // Test speech synthesis
                    Console.WriteLine("\nTesting speech synthesis...");
                    
                    // Set output to audio file
                    string outputPath = Path.Combine(Environment.CurrentDirectory, "test_output.wav");
                    synth.SetOutputToWaveFile(outputPath);
                    
                    // Speak some text
                    string testText = "This is a test of the Sherpa ONNX text-to-speech system.";
                    Console.WriteLine($"Speaking: \"{testText}\"");
                    synth.Speak(testText);
                    
                    // Reset output to default
                    synth.SetOutputToDefaultAudioDevice();
                    
                    Console.WriteLine($"\nSpeech saved to: {outputPath}");
                    
                    // Try speaking to the default audio device
                    Console.WriteLine("\nNow testing speech to default audio device...");
                    Console.WriteLine("Speaking: \"Hello, can you hear me?\"");
                    synth.Speak("Hello, can you hear me?");
                    
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
    }
}
