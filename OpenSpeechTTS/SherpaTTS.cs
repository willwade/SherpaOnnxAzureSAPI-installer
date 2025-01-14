using NAudio.Wave;
using System;
using System.IO;
using SherpaNative;

namespace OpenSpeechTTS
{
    public class SherpaTTS : IDisposable
    {
        private readonly SherpaWrapper _tts;
        private bool _disposed;

        public SherpaTTS(string modelPath, string tokensPath, string lexiconPath, string dataDirPath)
        {
            if (!File.Exists(modelPath))
                throw new FileNotFoundException($"Model file not found: {modelPath}");

            if (!File.Exists(tokensPath))
                throw new FileNotFoundException($"Tokens file not found: {tokensPath}");

            if (!Directory.Exists(dataDirPath))
                throw new DirectoryNotFoundException($"Data directory not found: {dataDirPath}");

            _tts = new SherpaWrapper(modelPath, tokensPath);
        }

        public void SpeakToWaveStream(string text, Stream stream)
        {
            if (_disposed)
                throw new ObjectDisposedException(nameof(SherpaTTS));

            var bytes = _tts.GenerateWaveform(text);

            using (var writer = new BinaryWriter(stream))
            {
                writer.Write(0x46464952); // "RIFF"
                writer.Write(36 + bytes.Length);
                writer.Write(0x45564157); // "WAVE"
                writer.Write(0x20746D66); // "fmt "
                writer.Write(16);
                writer.Write((short)1); // PCM
                writer.Write((short)1); // Mono
                writer.Write(22050); // Sample rate
                writer.Write(22050 * 2); // Bytes per second
                writer.Write((short)2); // Block align
                writer.Write((short)16); // Bits per sample
                writer.Write(0x61746164); // "data"
                writer.Write(bytes.Length);
                writer.Write(bytes);
            }
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                _tts?.Dispose();
                _disposed = true;
            }
        }
    }
}
