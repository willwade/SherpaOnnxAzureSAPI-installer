using System;
using System.IO;
using System.Runtime.InteropServices;
using Installer.Core.Interfaces;
using Microsoft.Win32;

namespace Installer.Core.Base
{
    /// <summary>
    /// Base class for SAPI voice implementations
    /// </summary>
    public abstract class SapiVoiceImplBase : ISapiVoiceImpl
    {
        private readonly string _logDir = "C:\\OpenSpeech\\Logs";
        protected bool _initialized;
        
        /// <summary>
        /// Gets the voice token from the registry
        /// </summary>
        /// <param name="clsid">The CLSID of the voice</param>
        /// <returns>The voice token</returns>
        protected string GetVoiceToken(Guid clsid)
        {
            string voiceToken = null;
            using (var key = Registry.ClassesRoot.OpenSubKey($@"CLSID\{clsid.ToString("B")}\Token"))
            {
                if (key != null)
                {
                    voiceToken = (string)key.GetValue("");
                }
            }
            
            if (string.IsNullOrEmpty(voiceToken))
            {
                throw new Exception("Voice token not found in registry");
            }
            
            return voiceToken;
        }
        
        /// <summary>
        /// Gets the voice attributes from the registry
        /// </summary>
        /// <param name="voiceToken">The voice token</param>
        /// <returns>The registry key containing the voice attributes</returns>
        protected RegistryKey GetVoiceAttributes(string voiceToken)
        {
            string registryPath = $@"SOFTWARE\Microsoft\Speech\Voices\Tokens\{voiceToken}";
            
            var voiceKey = Registry.LocalMachine.OpenSubKey(registryPath);
            if (voiceKey == null)
                throw new Exception($"Voice registry key not found: {registryPath}");
            
            var attributesKey = voiceKey.OpenSubKey("Attributes");
            if (attributesKey == null)
                throw new Exception("Voice attributes not found in registry");
            
            return attributesKey;
        }
        
        /// <summary>
        /// Speaks the provided text
        /// </summary>
        /// <param name="text">Text to speak</param>
        /// <param name="flags">SAPI flags</param>
        /// <param name="reserved">Reserved parameter</param>
        public abstract void Speak(string text, uint flags, IntPtr reserved);
        
        /// <summary>
        /// Gets the output format for the voice
        /// </summary>
        /// <param name="targetFormatId">Target format ID</param>
        /// <param name="targetFormat">Target format</param>
        /// <param name="actualFormatId">Actual format ID</param>
        /// <param name="actualFormat">Actual format</param>
        public virtual void GetOutputFormat(ref Guid targetFormatId, ref WaveFormatEx targetFormat, out Guid actualFormatId, out WaveFormatEx actualFormat)
        {
            // Default implementation - use PCM 16-bit 22050Hz mono
            actualFormatId = targetFormatId;
            actualFormat = new WaveFormatEx
            {
                wFormatTag = 1, // PCM
                nChannels = 1, // Mono
                nSamplesPerSec = 22050, // 22.05kHz
                wBitsPerSample = 16, // 16-bit
                nBlockAlign = 2, // 2 bytes per sample (16-bit mono)
                nAvgBytesPerSec = 44100, // 22050 * 2
                cbSize = 0
            };
        }
        
        /// <summary>
        /// Logs a message to the log file
        /// </summary>
        /// <param name="message">The message to log</param>
        protected void LogMessage(string message)
        {
            try
            {
                // Create log directory if it doesn't exist
                if (!Directory.Exists(_logDir))
                {
                    Directory.CreateDirectory(_logDir);
                }
                
                string logFile = Path.Combine(_logDir, "SapiVoice.log");
                File.AppendAllText(logFile, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] INFO: {message}{Environment.NewLine}");
            }
            catch
            {
                // Ignore logging errors
            }
        }
        
        /// <summary>
        /// Logs an error message to the log file
        /// </summary>
        /// <param name="message">The error message</param>
        /// <param name="ex">The exception, if any</param>
        protected void LogError(string message, Exception ex = null)
        {
            try
            {
                // Create log directory if it doesn't exist
                if (!Directory.Exists(_logDir))
                {
                    Directory.CreateDirectory(_logDir);
                }
                
                string logFile = Path.Combine(_logDir, "SapiVoice.log");
                string errorMessage = ex != null ? $"{message} - {ex.Message}{Environment.NewLine}{ex.StackTrace}" : message;
                File.AppendAllText(logFile, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] ERROR: {errorMessage}{Environment.NewLine}");
            }
            catch
            {
                // Ignore logging errors
            }
        }
    }
} 