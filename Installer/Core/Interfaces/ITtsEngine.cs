using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Installer.Core.Models;

namespace Installer.Core.Interfaces
{
    /// <summary>
    /// Interface defining common functionality for all TTS engines
    /// </summary>
    public interface ITtsEngine
    {
        // Basic properties
        string EngineName { get; }
        string EngineVersion { get; }
        string EngineDescription { get; }
        bool RequiresSsml { get; }
        bool RequiresAuthentication { get; }
        bool SupportsOfflineUsage { get; }
        
        // Configuration
        IEnumerable<ConfigurationParameter> GetRequiredParameters();
        bool ValidateConfiguration(Dictionary<string, string> config);
        
        // Voice management
        Task<IEnumerable<TtsVoiceInfo>> GetAvailableVoicesAsync(Dictionary<string, string> config);
        Task<bool> TestVoiceAsync(string voiceId, Dictionary<string, string> config);
        
        // Speech synthesis
        Task<byte[]> SynthesizeSpeechAsync(string text, string voiceId, Dictionary<string, string> parameters);
        Task<Stream> SynthesizeSpeechToStreamAsync(string text, string voiceId, Dictionary<string, string> parameters);
        
        // SAPI integration
        Guid GetEngineClsid();
        Type GetSapiImplementationType();
        void RegisterVoice(TtsVoiceInfo voice, Dictionary<string, string> config, string dllPath);
        void UnregisterVoice(string voiceId);
        
        // Lifecycle
        void Initialize();
        void Shutdown();
    }
} 