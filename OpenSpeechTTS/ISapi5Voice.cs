using System;
using System.Runtime.InteropServices;

namespace OpenSpeechTTS
{
    // SAPI5 TTS Engine interface - this is the correct interface for TTS engines
    [ComVisible(true)]
    [Guid("5AEF0FDD-3E3E-4E0E-96E6-DA4B71F0E0F9")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface ISpTTSEngine
    {
        [PreserveSig]
        int Speak(uint dwSpeakFlags, ref Guid rguidFormatId, ref WaveFormatEx pWaveFormatEx,
                  ref SpTTSFragList pTextFragList, IntPtr pOutputSite);

        [PreserveSig]
        int GetOutputFormat(ref Guid pTargetFormatId, ref WaveFormatEx pTargetWaveFormatEx,
                           out Guid pOutputFormatId, out IntPtr ppCoMemOutputWaveFormatEx);
    }

    // SAPI5 Output Site interface for streaming audio
    [ComVisible(true)]
    [Guid("3504B7C3-A8B8-4C60-8B24-C7D8A1E5C147")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface ISpTTSEngineSite
    {
        [PreserveSig]
        int AddEvents(IntPtr pEventArray, uint ulCount);

        [PreserveSig]
        int GetEventInterest(out ulong pullEventInterest);

        [PreserveSig]
        int GetActions();

        [PreserveSig]
        int Write(IntPtr pBuff, uint cb, out uint pcbWritten);

        [PreserveSig]
        int GetRate(out int pRateAdjust);

        [PreserveSig]
        int GetVolume(out ushort pusVolume);

        [PreserveSig]
        int GetSkipInfo(out int peType, out int plNumItems);

        [PreserveSig]
        int CompleteSkip(int ulNumSkipped);
    }

    // Text fragment structure for SAPI5
    [StructLayout(LayoutKind.Sequential)]
    public struct SpTTSFragList
    {
        public IntPtr pNext;
        public int State;
        public IntPtr pTextStart;
        public uint ulTextLen;
        public uint ulTextSrcOffset;
    }

    // Wave format structure
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

    // SAPI5 Event structure
    [StructLayout(LayoutKind.Sequential)]
    public struct SpEvent
    {
        public ushort eEventId;
        public ushort elParamType;
        public uint ulStreamNum;
        public ulong ullAudioStreamOffset;
        public IntPtr wParam;
        public IntPtr lParam;
    }

    // SAPI5 Event types
    public enum SpEventIds : ushort
    {
        SPEI_START_INPUT_STREAM = 1,
        SPEI_END_INPUT_STREAM = 2,
        SPEI_VOICE_CHANGE = 3,
        SPEI_TTS_BOOKMARK = 4,
        SPEI_WORD_BOUNDARY = 5,
        SPEI_PHONEME = 6,
        SPEI_SENTENCE_BOUNDARY = 7,
        SPEI_VISEME = 8,
        SPEI_TTS_AUDIO_LEVEL = 9,
        SPEI_TTS_PRIVATE = 15,
        SPEI_MIN_TTS = 1,
        SPEI_MAX_TTS = 15
    }

    // SAPI5 Speak flags
    [Flags]
    public enum SpVoiceFlags : uint
    {
        SPF_DEFAULT = 0,
        SPF_ASYNC = 1,
        SPF_PURGEBEFORESPEAK = 2,
        SPF_IS_FILENAME = 4,
        SPF_IS_XML = 8,
        SPF_IS_NOT_XML = 16,
        SPF_PERSIST_XML = 32,
        SPF_NLP_SPEAK_PUNC = 64,
        SPF_PARSE_SAPI = 128,
        SPF_PARSE_SSML = 256
    }
}
