using System;
using System.IO;
using System.Runtime.InteropServices;
using Installer.Core.Base;
using Installer.Core.Interfaces;
using Microsoft.Win32;

namespace Installer.Engines.SherpaOnnx
{
    /// <summary>
    /// SAPI voice implementation for Sherpa ONNX
    /// </summary>
    [Guid("3d8f5c5d-9d6b-4b92-a12b-1a6dff80b6b2")]
    [ComVisible(true)]
    public class SherpaOnnxSapi5VoiceImpl : SapiVoiceImplBase
    {
        private OpenSpeechTTS.SherpaTTS _sherpaTts;
        
        /// <summary>
        /// Creates a new instance of the SherpaOnnxSapi5VoiceImpl class
        /// </summary>
        public SherpaOnnxSapi5VoiceImpl()
        {
            try
            {
                // Get the voice token from the registry
                string voiceToken = GetVoiceToken(new Guid("3d8f5c5d-9d6b-4b92-a12b-1a6dff80b6b2"));
                
                // Get voice attributes
                var attributesKey = GetVoiceAttributes(voiceToken);
                
                // Get model paths
                var modelPath = (string)attributesKey.GetValue("Model Path");
                var tokensPath = (string)attributesKey.GetValue("Tokens Path");
                var lexiconPath = (string)attributesKey.GetValue("Lexicon Path");
                var dataDir = (string)attributesKey.GetValue("Data Directory");
                
                if (string.IsNullOrEmpty(modelPath))
                    throw new Exception("ModelPath not found in registry");
                if (string.IsNullOrEmpty(tokensPath))
                    throw new Exception("TokensPath not found in registry");
                if (string.IsNullOrEmpty(dataDir))
                    throw new Exception("DataDir not found in registry");
                
                // Create Sherpa TTS instance
                _sherpaTts = new OpenSpeechTTS.SherpaTTS(modelPath, tokensPath, lexiconPath, dataDir);
                _initialized = true;
                
                LogMessage($"Initialized SherpaOnnxSapi5VoiceImpl with model: {modelPath}");
            }
            catch (Exception ex)
            {
                LogError("Error initializing SherpaOnnxSapi5VoiceImpl", ex);
                throw;
            }
        }
        
        /// <summary>
        /// Speaks the provided text
        /// </summary>
        /// <param name="text">Text to speak</param>
        /// <param name="flags">SAPI flags</param>
        /// <param name="reserved">Reserved parameter</param>
        public override void Speak(string text, uint flags, IntPtr reserved)
        {
            try
            {
                if (!_initialized || _sherpaTts == null)
                {
                    LogError("SherpaOnnxSapi5VoiceImpl not initialized");
                    return;
                }
                
                LogMessage($"Speaking text: {text}");
                
                // Generate audio
                byte[] audioData = _sherpaTts.GenerateAudio(text);
                
                // Play audio
                _sherpaTts.PlayAudio(audioData);
            }
            catch (Exception ex)
            {
                LogError("Error speaking text", ex);
            }
        }
        
        /// <summary>
        /// Finalizes the instance
        /// </summary>
        ~SherpaOnnxSapi5VoiceImpl()
        {
            try
            {
                _sherpaTts?.Dispose();
                _sherpaTts = null;
            }
            catch
            {
                // Ignore errors during finalization
            }
        }
    }
} 