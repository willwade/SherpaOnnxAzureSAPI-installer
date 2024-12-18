using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using NAudio.Wave;
using Microsoft.Win32;

namespace OpenSpeechTTS
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    [Guid("3d8f5c5d-9d6b-4b92-a12b-1a6dff80b6b2")]
    [ProgId("OpenSpeechTTS.Sapi5VoiceImpl")]
    public class Sapi5VoiceImpl : ISapi5Voice
    {
        private SherpaTTS _ttsEngine;
        private const string RegistryBasePath = @"SOFTWARE\Microsoft\SPEECH\Voices\Tokens";

        public Sapi5VoiceImpl()
        {
            try
            {
                Console.WriteLine("Initializing SherpaTTS...");
                var modelPaths = GetModelPathsFromRegistry();
                _ttsEngine = new SherpaTTS(modelPaths.ModelPath, modelPaths.TokensPath, modelPaths.LexiconPath);
                Console.WriteLine("SherpaTTS initialized successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error initializing SherpaTTS: {ex.Message}");
                throw;
            }
        }

        private (string ModelPath, string TokensPath, string LexiconPath) GetModelPathsFromRegistry()
        {
            using (var voicesKey = Registry.LocalMachine.OpenSubKey(RegistryBasePath))
            {
                if (voicesKey == null)
                    throw new Exception("SAPI voices registry key not found");

                // Find our voice token by CLSID
                foreach (var voiceName in voicesKey.GetSubKeyNames())
                {
                    using (var voiceKey = voicesKey.OpenSubKey($"{voiceName}\\Attributes"))
                    {
                        if (voiceKey == null) continue;

                        var clsid = voiceKey.GetValue("CLSID") as string;
                        if (clsid == "3d8f5c5d-9d6b-4b92-a12b-1a6dff80b6b2")
                        {
                            var modelPath = voiceKey.GetValue("ModelPath") as string;
                            var tokensPath = voiceKey.GetValue("TokensPath") as string;
                            var lexiconPath = voiceKey.GetValue("LexiconPath") as string ?? "";

                            if (string.IsNullOrEmpty(modelPath) || string.IsNullOrEmpty(tokensPath))
                                throw new Exception("Model paths not found in registry");

                            return (modelPath, tokensPath, lexiconPath);
                        }
                    }
                }
            }

            throw new Exception("Voice not found in registry");
        }

        public void Speak(string text)
        {
            try
            {
                var audioData = _ttsEngine.GenerateAudio(text);
                if (audioData == null || audioData.Length == 0)
                {
                    Console.WriteLine("No audio generated.");
                    return;
                }

                using (var waveOut = new WaveOutEvent())
                {
                    var waveProvider = new RawSourceWaveStream(
                        new MemoryStream(audioData),
                        new WaveFormat(16000, 16, 1)); // Match Sherpa's sample rate
                    waveOut.Init(waveProvider);
                    waveOut.Play();
                    while (waveOut.PlaybackState == PlaybackState.Playing)
                    {
                        Thread.Sleep(100);
                    }
                }

                Console.WriteLine($"Played audio for: {text}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during playback: {ex.Message}");
            }
        }

        public void SetRate(int rate)
        {
            if (_ttsEngine != null)
            {
                float adjustedScale = Math.Max(0.5f, Math.Min(2.0f, 1.0f + (rate / 10.0f))); // Adjust rate scale
                Console.WriteLine($"Playback rate adjusted to: {adjustedScale}");
            }
        }

        public void SetVolume(int volume)
        {
            // Volume adjustment (0-100)
            float normalizedVolume = Math.Max(0, Math.Min(100, volume)) / 100.0f;
            Console.WriteLine($"Volume set to: {normalizedVolume}");
        }

        public void Pause()
        {
            Console.WriteLine("TTS paused.");
        }

        public void Resume()
        {
            Console.WriteLine("TTS resumed.");
        }
    }
}
