using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Installer.Core.Interfaces;
using Installer.Core.Models;
using Microsoft.Win32;

namespace Installer.Core.Base
{
    /// <summary>
    /// Base class for TTS engines with common implementation
    /// </summary>
    public abstract class TtsEngineBase : ITtsEngine
    {
        private readonly string _logDir = "C:\\OpenSpeech\\Logs";
        
        // Basic properties
        public abstract string EngineName { get; }
        public virtual string EngineVersion => "1.0";
        public virtual string EngineDescription => $"{EngineName} Text-to-Speech Engine";
        public virtual bool RequiresSsml => false;
        public virtual bool RequiresAuthentication => true;
        public virtual bool SupportsOfflineUsage => false;
        
        // Configuration
        public virtual IEnumerable<ConfigurationParameter> GetRequiredParameters() => new List<ConfigurationParameter>();
        
        public virtual bool ValidateConfiguration(Dictionary<string, string> config)
        {
            // Basic validation - check that all required parameters are present
            foreach (var param in GetRequiredParameters())
            {
                if (param.IsRequired && (!config.ContainsKey(param.Name) || string.IsNullOrEmpty(config[param.Name])))
                {
                    LogError($"Missing required parameter: {param.Name}");
                    return false;
                }
            }
            
            return true;
        }
        
        // Voice management
        public abstract Task<IEnumerable<TtsVoiceInfo>> GetAvailableVoicesAsync(Dictionary<string, string> config);
        
        public virtual async Task<bool> TestVoiceAsync(string voiceId, Dictionary<string, string> config)
        {
            try
            {
                // Default implementation - try to synthesize a short text
                var result = await SynthesizeSpeechAsync("This is a test.", voiceId, config);
                return result != null && result.Length > 0;
            }
            catch (Exception ex)
            {
                LogError($"Error testing voice {voiceId}", ex);
                return false;
            }
        }
        
        // Speech synthesis
        public abstract Task<byte[]> SynthesizeSpeechAsync(string text, string voiceId, Dictionary<string, string> parameters);
        
        public virtual async Task<Stream> SynthesizeSpeechToStreamAsync(string text, string voiceId, Dictionary<string, string> parameters)
        {
            var data = await SynthesizeSpeechAsync(text, voiceId, parameters);
            return new MemoryStream(data);
        }
        
        // SAPI integration
        public abstract Guid GetEngineClsid();
        
        public abstract Type GetSapiImplementationType();
        
        public virtual void RegisterVoice(TtsVoiceInfo voice, Dictionary<string, string> config, string dllPath)
        {
            try
            {
                // Create registry key for the voice
                string voiceToken = voice.Name.Replace(" ", "");
                string registryPath = $@"SOFTWARE\Microsoft\Speech\Voices\Tokens\{voiceToken}";
                
                using (var key = Registry.LocalMachine.CreateSubKey(registryPath))
                {
                    if (key == null)
                    {
                        throw new Exception($"Failed to create registry key: {registryPath}");
                    }
                    
                    // Set basic voice properties
                    key.SetValue("", voice.DisplayName);
                    key.SetValue("CLSID", GetEngineClsid().ToString("B"));
                    key.SetValue("Path", dllPath);
                    
                    // Create attributes subkey
                    using (var attributesKey = key.CreateSubKey("Attributes"))
                    {
                        if (attributesKey == null)
                        {
                            throw new Exception($"Failed to create attributes key: {registryPath}\\Attributes");
                        }
                        
                        // Set standard attributes
                        attributesKey.SetValue("Gender", voice.Gender);
                        attributesKey.SetValue("Age", voice.Age);
                        attributesKey.SetValue("Name", voice.Name);
                        attributesKey.SetValue("Language", ConvertLocaleToLcid(voice.Locale));
                        attributesKey.SetValue("Vendor", EngineName);
                        attributesKey.SetValue("Version", EngineVersion);
                        attributesKey.SetValue("VoiceType", EngineName);
                        
                        // Set engine-specific attributes
                        foreach (var attr in voice.AdditionalAttributes)
                        {
                            attributesKey.SetValue(attr.Key, attr.Value);
                        }
                        
                        // Set style and role if supported
                        if (voice.SupportsStyles && !string.IsNullOrEmpty(voice.SelectedStyle))
                        {
                            attributesKey.SetValue("SelectedStyle", voice.SelectedStyle);
                        }
                        
                        if (voice.SupportsRoles && !string.IsNullOrEmpty(voice.SelectedRole))
                        {
                            attributesKey.SetValue("SelectedRole", voice.SelectedRole);
                        }
                    }
                }
                
                // Create token link in CLSID
                string clsidPath = $@"CLSID\{GetEngineClsid().ToString("B")}\Token";
                using (var key = Registry.ClassesRoot.CreateSubKey(clsidPath))
                {
                    if (key == null)
                    {
                        throw new Exception($"Failed to create registry key: {clsidPath}");
                    }
                    
                    key.SetValue("", voiceToken);
                }
                
                LogMessage($"Successfully registered voice: {voice.Name}");
            }
            catch (Exception ex)
            {
                LogError($"Error registering voice: {voice.Name}", ex);
                throw;
            }
        }
        
        public virtual void UnregisterVoice(string voiceId)
        {
            try
            {
                string registryPath = $@"SOFTWARE\Microsoft\Speech\Voices\Tokens\{voiceId}";
                Registry.LocalMachine.DeleteSubKeyTree(registryPath, false);
                LogMessage($"Successfully unregistered voice: {voiceId}");
            }
            catch (Exception ex)
            {
                LogError($"Error unregistering voice: {voiceId}", ex);
                throw;
            }
        }
        
        // Lifecycle
        public virtual void Initialize()
        {
            // Create log directory if it doesn't exist
            if (!Directory.Exists(_logDir))
            {
                Directory.CreateDirectory(_logDir);
            }
            
            LogMessage($"Initializing {EngineName} engine");
        }
        
        public virtual void Shutdown()
        {
            LogMessage($"Shutting down {EngineName} engine");
        }
        
        // Helper methods
        protected virtual void LogMessage(string message)
        {
            try
            {
                string logFile = Path.Combine(_logDir, $"{EngineName}.log");
                File.AppendAllText(logFile, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] INFO: {message}{Environment.NewLine}");
            }
            catch
            {
                // Ignore logging errors
            }
        }
        
        protected virtual void LogError(string message, Exception ex = null)
        {
            try
            {
                string logFile = Path.Combine(_logDir, $"{EngineName}.log");
                string errorMessage = ex != null ? $"{message} - {ex.Message}{Environment.NewLine}{ex.StackTrace}" : message;
                File.AppendAllText(logFile, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] ERROR: {errorMessage}{Environment.NewLine}");
            }
            catch
            {
                // Ignore logging errors
            }
        }
        
        protected virtual string ConvertLocaleToLcid(string locale)
        {
            // Default implementation - this should be replaced with proper locale to LCID conversion
            // For now, just return a default value for English
            return "409"; // English (United States)
        }
    }
} 