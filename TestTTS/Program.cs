using System;
using System.IO;
using OpenSpeechTTS;

class Program
{
    static void Main(string[] args)
    {
        try
        {
            // Get the path to the test files
            string baseDir = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", ".."));
            string modelPath = Path.Combine(baseDir, "..", "models", "piper_en_southern", "vits.onnx");
            string tokensPath = Path.Combine(baseDir, "..", "models", "piper_en_southern", "tokens.txt");

            Console.WriteLine($"Using model: {modelPath}");
            Console.WriteLine($"Using tokens: {tokensPath}");

            // Create TTS instance
            using (var tts = new SherpaTTS(modelPath, tokensPath))
            {
                // Test text to synthesize
                string text = "Hello, this is a test of the Sherpa TTS system.";
                Console.WriteLine($"\nGenerating audio for text: {text}");

                // Generate audio
                byte[] audioData = tts.GenerateAudio(text);
                Console.WriteLine($"Generated {audioData.Length} bytes of audio data");

                // Save to WAV file
                string outputPath = Path.Combine(baseDir, "test_output.wav");
                using (var fs = new FileStream(outputPath, FileMode.Create))
                {
                    // Write WAV header
                    WriteWavHeader(fs, audioData.Length);
                    // Write audio data
                    fs.Write(audioData, 0, audioData.Length);
                }

                Console.WriteLine($"\nAudio saved to: {outputPath}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
        }
    }

    static void WriteWavHeader(Stream stream, int dataLength)
    {
        // RIFF header
        stream.Write(System.Text.Encoding.ASCII.GetBytes("RIFF"), 0, 4);
        stream.Write(BitConverter.GetBytes(dataLength + 36), 0, 4);
        stream.Write(System.Text.Encoding.ASCII.GetBytes("WAVE"), 0, 4);

        // fmt chunk
        stream.Write(System.Text.Encoding.ASCII.GetBytes("fmt "), 0, 4);
        stream.Write(BitConverter.GetBytes(16), 0, 4); // Subchunk1Size
        stream.Write(BitConverter.GetBytes((short)1), 0, 2); // AudioFormat (PCM)
        stream.Write(BitConverter.GetBytes((short)1), 0, 2); // NumChannels
        stream.Write(BitConverter.GetBytes(24000), 0, 4); // SampleRate
        stream.Write(BitConverter.GetBytes(48000), 0, 4); // ByteRate
        stream.Write(BitConverter.GetBytes((short)2), 0, 2); // BlockAlign
        stream.Write(BitConverter.GetBytes((short)16), 0, 2); // BitsPerSample

        // data chunk
        stream.Write(System.Text.Encoding.ASCII.GetBytes("data"), 0, 4);
        stream.Write(BitConverter.GetBytes(dataLength), 0, 4);
    }
}
