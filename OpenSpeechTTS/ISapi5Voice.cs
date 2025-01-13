using System;
using System.Runtime.InteropServices;

namespace OpenSpeechTTS
{
    [ComVisible(true)]
    [Guid("74c1c045-6e6a-4623-9735-2326a149af54")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ISapi5Voice
    {
        void Speak(string text);
        void SetRate(int rate);
        void SetVolume(int volume);
        void Pause();
        void Resume();
    }

    [ComVisible(true)]
    [Guid("9B9F4DEE-8934-47C0-A135-6A6E87D71264")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface ISpObjectToken
    {
        void GetId([MarshalAs(UnmanagedType.LPWStr)] out string objectId);
        void GetDescription(uint locale, [MarshalAs(UnmanagedType.LPWStr)] out string description);
    }

    [ComVisible(true)]
    [Guid("5AEF0FDD-3E3E-4E0E-96E6-DA4B71F0E0F9")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface ISpTTSEngine
    {
        void Speak([MarshalAs(UnmanagedType.LPWStr)] string text, uint flags, IntPtr reserved);
        void GetOutputFormat(ref Guid targetFormatId, ref WaveFormatEx targetFormat, out Guid actualFormatId, out WaveFormatEx actualFormat);
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct WaveFormatEx
    {
        public ushort wFormatTag;
        public ushort nChannels;
        public uint nSamplesPerSec;
        public uint nAvgBytesPerSec;
        public ushort nBlockAlign;
        public ushort wBitsPerSample;
        public ushort cbSize;
    }
}
