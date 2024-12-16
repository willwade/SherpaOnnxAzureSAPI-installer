using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using NAudio.Wave;

[ComVisible(true)]
[ClassInterface(ClassInterfaceType.None)]
[Guid("3d8f5c5d-9d6b-4b92-a12b-1a6dff80b6b2")] 
public class Sapi5VoiceImpl : ISapi5Voice
{
    private SherpaTTS _ttsEngine;

    public Sapi5VoiceImpl()
    {
        string modelPath = @"C:\Path\To\Models\model.onnx";
        string tokensPath = @"C:\Path\To\Models\tokens.txt";
        string lexiconPath = ""; // Leave empty for MMS models

        try
        {
            Console.WriteLine("Initializing SherpaTTS...");
            _ttsEngine = new SherpaTTS(modelPath, tokensPath, lexiconPath);
            Console.WriteLine("SherpaTTS initialized successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error initializing SherpaTTS: {ex.Message}");
            throw;
        }
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
            float adjustedScale = Math.Max(0.5f, Math.Min(2.0f, 1.0f / rate)); // Example clamping
            _ttsEngine.SetLengthScale(adjustedScale);
            Console.WriteLine($"Playback rate adjusted to: {rate}");
        }
    }

    public void SetVolume(int volume)
    {
        // Stub implementation for volume adjustment
        Console.WriteLine($"Volume set to: {volume}");
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
