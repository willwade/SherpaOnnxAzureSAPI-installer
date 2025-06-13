using System;
using Microsoft.Win32;
using Installer.Shared;

namespace Installer
{
    /// <summary>
    /// Base class for SAPI 5 voice registration
    /// Provides basic functionality for registering and unregistering SAPI voices
    /// </summary>
    public class Sapi5Registrar
    {
        protected const string RegistryBasePath = @"SOFTWARE\Microsoft\SPEECH\Voices\Tokens";

        /// <summary>
        /// Registers a TTS voice with SAPI 5
        /// </summary>
        /// <param name="model">TTS model to register</param>
        /// <param name="dllPath">Path to the COM DLL</param>
        public virtual void RegisterVoice(TtsModel model, string dllPath)
        {
            // This is a base implementation - derived classes should override
            Console.WriteLine($"Base RegisterVoice called for {model.Name}");
        }

        /// <summary>
        /// Unregisters a voice from SAPI 5
        /// </summary>
        /// <param name="voiceId">Voice ID to unregister</param>
        public virtual void UnregisterVoice(string voiceId)
        {
            try
            {
                string voiceRegistryPath = $@"{RegistryBasePath}\{voiceId}";
                Registry.LocalMachine.DeleteSubKeyTree(voiceRegistryPath, false);
                Console.WriteLine($"Unregistered voice: {voiceId}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error unregistering voice {voiceId}: {ex.Message}");
            }
        }

        /// <summary>
        /// Checks if a voice is registered
        /// </summary>
        /// <param name="voiceId">Voice ID to check</param>
        /// <returns>True if voice is registered</returns>
        public virtual bool IsVoiceRegistered(string voiceId)
        {
            try
            {
                string voiceRegistryPath = $@"{RegistryBasePath}\{voiceId}";
                using (var key = Registry.LocalMachine.OpenSubKey(voiceRegistryPath))
                {
                    return key != null;
                }
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Gets the registry path for a voice
        /// </summary>
        /// <param name="voiceId">Voice ID</param>
        /// <returns>Full registry path</returns>
        protected string GetVoiceRegistryPath(string voiceId)
        {
            return $@"{RegistryBasePath}\{voiceId}";
        }
    }
}
