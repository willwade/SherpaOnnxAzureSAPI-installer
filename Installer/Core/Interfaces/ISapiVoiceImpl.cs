using System;
using System.Runtime.InteropServices;

namespace Installer.Core.Interfaces
{
    /// <summary>
    /// Interface for SAPI voice implementations
    /// </summary>
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface ISapiVoiceImpl
    {
        /// <summary>
        /// Speaks the provided text
        /// </summary>
        /// <param name="text">Text to speak</param>
        /// <param name="flags">SAPI flags</param>
        /// <param name="reserved">Reserved parameter</param>
        void Speak([MarshalAs(UnmanagedType.LPWStr)] string text, uint flags, IntPtr reserved);
        
        /// <summary>
        /// Gets the output format for the voice
        /// </summary>
        /// <param name="targetFormatId">Target format ID</param>
        /// <param name="targetFormat">Target format</param>
        /// <param name="actualFormatId">Actual format ID</param>
        /// <param name="actualFormat">Actual format</param>
        void GetOutputFormat(ref Guid targetFormatId, ref WaveFormatEx targetFormat, out Guid actualFormatId, out WaveFormatEx actualFormat);
    }
    
    /// <summary>
    /// Wave format structure for SAPI integration
    /// </summary>
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