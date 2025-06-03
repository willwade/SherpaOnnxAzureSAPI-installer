using System;
using System.Runtime.InteropServices;

namespace OpenSpeechTTS
{
    // SAPI5 TTS Engine interface - OFFICIAL SAPI5 interface definition
    [ComVisible(true)]
    [Guid("A74D7C8E-4CC5-4F2F-A6EB-804DEE18500E")]  // Official SAPI5 ISpTTSEngine GUID
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface ISpTTSEngine
    {
        [PreserveSig]
        int Speak(
            [In] uint dwSpeakFlags,
            [In] ref Guid rguidFormatId,
            [In] ref WaveFormatEx pWaveFormatEx,
            [In] ref SPVTEXTFRAG pTextFragList,
            [In] IntPtr pOutputSite);

        [PreserveSig]
        int GetOutputFormat(
            [In] ref Guid pTargetFormatId,
            [In] ref WaveFormatEx pTargetWaveFormatEx,
            [Out] out Guid pOutputFormatId,
            [Out] out IntPtr ppCoMemOutputWaveFormatEx);
    }

    // SAPI5 Object with Token interface - REQUIRED for TTS engines
    [ComVisible(true)]
    [Guid("5B559F40-E952-11D2-BB91-00C04F8EE6C0")]  // Official SAPI5 ISpObjectWithToken GUID
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface ISpObjectWithToken
    {
        [PreserveSig]
        int SetObjectToken([In, MarshalAs(UnmanagedType.Interface)] object pToken);

        [PreserveSig]
        int GetObjectToken([Out, MarshalAs(UnmanagedType.Interface)] out object ppToken);
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

    // OFFICIAL SAPI5 Text fragment structure - SPVTEXTFRAG
    [StructLayout(LayoutKind.Sequential)]
    public struct SPVTEXTFRAG
    {
        public IntPtr pNext;           // struct SPVTEXTFRAG *pNext;
        public SPVSTATE State;         // SPVSTATE State;
        public IntPtr pTextStart;      // LPCWSTR pTextStart;
        public uint ulTextLen;         // ULONG ulTextLen;
        public uint ulTextSrcOffset;   // ULONG ulTextSrcOffset;
    }

    // OFFICIAL SAPI5 Voice state structure - SPVSTATE
    [StructLayout(LayoutKind.Sequential)]
    public struct SPVSTATE
    {
        public SPVACTIONS eAction;          // SPVACTIONS eAction;
        public ushort LangID;               // LANGID LangID;
        public ushort wReserved;            // WORD wReserved;
        public int EmphAdj;                 // long EmphAdj;
        public int RateAdj;                 // long RateAdj;
        public uint Volume;                 // ULONG Volume;
        public SPVPITCH PitchAdj;           // SPVPITCH PitchAdj;
        public uint SilenceMSecs;           // ULONG SilenceMSecs;
        public IntPtr pPhoneIds;            // SPPHONEID *pPhoneIds;
        public SPPARTOFSPEECH ePartOfSpeech; // SPPARTOFSPEECH ePartOfSpeech;
        public SPVCONTEXT Context;          // SPVCONTEXT Context;
    }

    // SAPI5 Voice actions enumeration
    public enum SPVACTIONS
    {
        SPVA_Speak = 0,
        SPVA_Silence,
        SPVA_Pronounce,
        SPVA_Bookmark,
        SPVA_SpellOut,
        SPVA_Section,
        SPVA_ParseUnknownTag
    }

    // SAPI5 Pitch adjustment structure
    [StructLayout(LayoutKind.Sequential)]
    public struct SPVPITCH
    {
        public int MiddleAdj;    // long MiddleAdj;
        public int RangeAdj;     // long RangeAdj;
    }

    // SAPI5 Part of speech enumeration (simplified)
    public enum SPPARTOFSPEECH
    {
        SPPS_Unknown = 0,
        SPPS_Noun = 0x1000,
        SPPS_Verb = 0x2000,
        SPPS_Modifier = 0x3000,
        SPPS_Function = 0x4000,
        SPPS_Interjection = 0x5000
    }

    // SAPI5 Context structure
    [StructLayout(LayoutKind.Sequential)]
    public struct SPVCONTEXT
    {
        public IntPtr pCategory;    // LPCWSTR pCategory;
        public IntPtr pBefore;      // LPCWSTR pBefore;
        public IntPtr pAfter;       // LPCWSTR pAfter;
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
