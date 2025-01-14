using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.Win32;
using NAudio.Wave;

namespace OpenSpeechTTS
{
    [ComVisible(true)]
    [Guid("3d8f5c5d-9d6b-4b92-a12b-1a6dff80b6b2")]
    [ProgId("OpenSpeechTTS.Sapi5VoiceImpl")]
    [ClassInterface(ClassInterfaceType.None)]
    public class Sapi5VoiceImpl : ISapi5Voice, ISpObjectToken, ISpTTSEngine
    {
        private SherpaTTS _ttsEngine;
        private const string RegistryBasePath = @"SOFTWARE\Microsoft\Speech\Voices\Tokens\OpenSpeechAmyVoice";
        private float _speed = 1.0f;
        private bool _isPaused = false;
        private int _volume = 100;

        public Sapi5VoiceImpl()
        {
            try
            {
                using (var key = Registry.LocalMachine.OpenSubKey(RegistryBasePath))
                {
                    if (key == null)
                    {
                        throw new Exception($"Registry key not found: {RegistryBasePath}");
                    }

                    var modelPath = key.GetValue("ModelPath") as string;
                    var tokensPath = key.GetValue("TokensPath") as string;
                    var dataDirPath = key.GetValue("DataDirPath") as string;

                    Console.WriteLine($"Registry values read:");
                    Console.WriteLine($"  ModelPath: {modelPath}");
                    Console.WriteLine($"  TokensPath: {tokensPath}");
                    Console.WriteLine($"  DataDirPath: {dataDirPath}");

                    if (string.IsNullOrEmpty(modelPath))
                        throw new Exception("ModelPath not found in registry");
                    if (string.IsNullOrEmpty(tokensPath))
                        throw new Exception("TokensPath not found in registry");
                    if (string.IsNullOrEmpty(dataDirPath))
                        throw new Exception("DataDirPath not found in registry");

                    if (!File.Exists(modelPath))
                        throw new Exception($"Model file not found at: {modelPath}");
                    if (!File.Exists(tokensPath))
                        throw new Exception($"Tokens file not found at: {tokensPath}");
                    if (!Directory.Exists(dataDirPath))
                        throw new Exception($"Data directory not found at: {dataDirPath}");

                    Console.WriteLine("All files exist, creating SherpaTTS instance...");
                    _ttsEngine = new SherpaTTS(modelPath, tokensPath, "", dataDirPath);
                    Console.WriteLine("SherpaTTS initialized successfully.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in Sapi5VoiceImpl constructor: {ex.Message}");
                Console.WriteLine($"Stack trace: {ex.StackTrace}");
                throw;
            }
        }

        public void Speak(string text)
        {
            try
            {
                if (_ttsEngine == null)
                {
                    Console.WriteLine("Error: TTS engine is null");
                    return;
                }

                if (_isPaused)
                {
                    Console.WriteLine("Speech is paused");
                    return;
                }

                Console.WriteLine($"Generating audio for text: {text}");
                var audioData = _ttsEngine.GenerateAudio(text, _speed);
                if (audioData == null || audioData.Length == 0)
                {
                    Console.WriteLine("No audio generated.");
                    return;
                }

                Console.WriteLine($"Generated {audioData.Length} bytes of audio data");
                using (var waveOut = new WaveOutEvent())
                {
                    waveOut.Volume = _volume / 100f;
                    var waveProvider = new RawSourceWaveStream(
                        new MemoryStream(audioData),
                        new WaveFormat(_ttsEngine.SampleRate, 16, 1));
                    waveOut.Init(waveProvider);
                    waveOut.Play();
                    Console.WriteLine("Playing audio...");
                    while (waveOut.PlaybackState == PlaybackState.Playing && !_isPaused)
                    {
                        Thread.Sleep(100);
                    }
                }

                Console.WriteLine($"Finished playing audio for: {text}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during playback: {ex.Message}");
                Console.WriteLine($"Stack trace: {ex.StackTrace}");
                throw;
            }
        }

        public void SetRate(int rate)
        {
            Console.WriteLine($"Setting rate to: {rate}");
            // Convert SAPI rate (-10 to 10) to speed multiplier (0.5 to 2.0)
            _speed = 1.0f + (rate / 10.0f);
            if (_speed < 0.5f) _speed = 0.5f;
            if (_speed > 2.0f) _speed = 2.0f;
            Console.WriteLine($"Speed multiplier set to: {_speed}");
        }

        public void SetVolume(int volume)
        {
            Console.WriteLine($"Setting volume to: {volume}");
            _volume = Math.Max(0, Math.Min(100, volume));
            Console.WriteLine($"Volume set to: {_volume}");
        }

        public void Pause()
        {
            Console.WriteLine("Pausing speech...");
            _isPaused = true;
            Console.WriteLine("Speech paused.");
        }

        public void Resume()
        {
            Console.WriteLine("Resuming speech...");
            _isPaused = false;
            Console.WriteLine("Speech resumed.");
        }

        // ISpObjectToken implementation
        public void GetId(out string objectId)
        {
            Console.WriteLine("Getting object ID...");
            objectId = RegistryBasePath;
            Console.WriteLine($"Object ID: {objectId}");
        }

        public void GetDescription(uint locale, out string description)
        {
            Console.WriteLine("Getting description...");
            description = "OpenSpeech Amy Voice";
            Console.WriteLine($"Description: {description}");
        }

        // ISpTTSEngine implementation
        public void Speak(string text, uint flags, IntPtr reserved)
        {
            Console.WriteLine("Speaking text with flags and reserved...");
            Speak(text);
        }

        public void GetOutputFormat(ref Guid targetFormatId, ref WaveFormatEx targetFormat, out Guid actualFormatId, out WaveFormatEx actualFormat)
        {
            Console.WriteLine("Getting output format...");
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
            Console.WriteLine($"Actual format: {actualFormat}");
        }
    }
}
