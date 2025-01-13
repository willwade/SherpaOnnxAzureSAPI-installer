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
    [ComDefaultInterface(typeof(ISpTTSEngine))]
    public class Sapi5VoiceImpl : ISapi5Voice, ISpObjectToken, ISpTTSEngine
    {
        private SherpaTTS _ttsEngine;
        private const string RegistryBasePath = @"SOFTWARE\Microsoft\SPEECH\Voices\Tokens\OpenSpeechAmyVoice";
        private float _speed = 1.0f;
        private bool _isPaused = false;
        private int _volume = 100;

        public Sapi5VoiceImpl()
        {
            try
            {
                Console.WriteLine("Initializing SherpaTTS...");
                var modelPaths = GetModelPathsFromRegistry();
                _ttsEngine = new SherpaTTS(modelPaths.ModelPath, modelPaths.TokensPath, modelPaths.LexiconPath, modelPaths.DataDirPath);
                Console.WriteLine("SherpaTTS initialized successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error initializing SherpaTTS: {ex.Message}");
                throw;
            }
        }

        private (string ModelPath, string TokensPath, string LexiconPath, string DataDirPath) GetModelPathsFromRegistry()
        {
            using (var voiceKey = Registry.LocalMachine.OpenSubKey(RegistryBasePath))
            {
                if (voiceKey == null)
                    throw new Exception("Voice registry key not found");

                var modelPath = voiceKey.GetValue("ModelPath") as string;
                var tokensPath = voiceKey.GetValue("TokensPath") as string;
                var lexiconPath = voiceKey.GetValue("LexiconPath") as string ?? "";
                var dataDirPath = voiceKey.GetValue("DataDirPath") as string ?? "";

                if (string.IsNullOrEmpty(modelPath) || string.IsNullOrEmpty(tokensPath))
                    throw new Exception("Model paths not found in registry");

                return (modelPath, tokensPath, lexiconPath, dataDirPath);
            }
        }

        public void Speak(string text)
        {
            try
            {
                if (_isPaused) return;

                var audioData = _ttsEngine.GenerateAudio(text, _speed);
                if (audioData == null || audioData.Length == 0)
                {
                    Console.WriteLine("No audio generated.");
                    return;
                }

                using (var waveOut = new WaveOutEvent())
                {
                    waveOut.Volume = _volume / 100f;
                    var waveProvider = new RawSourceWaveStream(
                        new MemoryStream(audioData),
                        new WaveFormat(_ttsEngine.SampleRate, 16, 1));
                    waveOut.Init(waveProvider);
                    waveOut.Play();
                    while (waveOut.PlaybackState == PlaybackState.Playing && !_isPaused)
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
            // Convert SAPI rate (-10 to 10) to speed multiplier (0.5 to 2.0)
            _speed = 1.0f + (rate / 10.0f);
            if (_speed < 0.5f) _speed = 0.5f;
            if (_speed > 2.0f) _speed = 2.0f;
        }

        public void SetVolume(int volume)
        {
            _volume = Math.Max(0, Math.Min(100, volume));
        }

        public void Pause()
        {
            _isPaused = true;
        }

        public void Resume()
        {
            _isPaused = false;
        }

        // ISpObjectToken implementation
        public void GetId(out string objectId)
        {
            objectId = RegistryBasePath;
        }

        public void GetDescription(uint locale, out string description)
        {
            description = "OpenSpeech Amy Voice";
        }

        // ISpTTSEngine implementation
        public void Speak(string text, uint flags, IntPtr reserved)
        {
            Speak(text);
        }

        public void GetOutputFormat(ref Guid targetFormatId, ref WaveFormatEx targetFormat, out Guid actualFormatId, out WaveFormatEx actualFormat)
        {
            actualFormatId = targetFormatId;
            actualFormat = new WaveFormatEx
            {
                wFormatTag = 1, // PCM
                nChannels = 1,
                nSamplesPerSec = (uint)_ttsEngine.SampleRate,
                wBitsPerSample = 16,
                nBlockAlign = 2,
                nAvgBytesPerSec = (uint)_ttsEngine.SampleRate * 2,
                cbSize = 0
            };
        }
    }
}
