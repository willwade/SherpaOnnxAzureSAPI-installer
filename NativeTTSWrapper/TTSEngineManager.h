#pragma once

#include "ITTSEngine.h"
#include <windows.h>
#include <map>
#include <mutex>
#include <string>
#include <memory>

namespace NativeTTS {

    /// <summary>
    /// Manages multiple TTS engines and provides thread-safe access
    /// Keeps engines warm and ready for immediate synthesis
    /// </summary>
    class TTSEngineManager {
    public:
        TTSEngineManager();
        ~TTSEngineManager();

        /// <summary>
        /// Initialize an engine with the given configuration
        /// </summary>
        /// <param name="engineId">Unique identifier for this engine instance</param>
        /// <param name="type">Type of engine to create</param>
        /// <param name="config">JSON configuration for the engine</param>
        /// <returns>S_OK on success, error HRESULT on failure</returns>
        HRESULT InitializeEngine(const std::wstring& engineId, EngineType type, const std::wstring& config);

        /// <summary>
        /// Get an initialized engine by ID
        /// </summary>
        /// <param name="engineId">Engine identifier</param>
        /// <returns>Pointer to engine, or nullptr if not found</returns>
        ITTSEngine* GetEngine(const std::wstring& engineId);

        /// <summary>
        /// Shutdown a specific engine
        /// </summary>
        /// <param name="engineId">Engine identifier</param>
        /// <returns>S_OK on success, error HRESULT on failure</returns>
        HRESULT ShutdownEngine(const std::wstring& engineId);

        /// <summary>
        /// Shutdown all engines
        /// </summary>
        /// <returns>S_OK on success, error HRESULT on failure</returns>
        HRESULT ShutdownAllEngines();

        /// <summary>
        /// Get list of initialized engine IDs
        /// </summary>
        /// <returns>Vector of engine IDs</returns>
        std::vector<std::wstring> GetEngineIds() const;

        /// <summary>
        /// Check if an engine is initialized
        /// </summary>
        /// <param name="engineId">Engine identifier</param>
        /// <returns>true if engine exists and is initialized</returns>
        bool IsEngineInitialized(const std::wstring& engineId) const;

        /// <summary>
        /// Load configuration from JSON file
        /// </summary>
        /// <param name="configPath">Path to JSON configuration file</param>
        /// <returns>S_OK on success, error HRESULT on failure</returns>
        HRESULT LoadConfiguration(const std::wstring& configPath);

        /// <summary>
        /// Get engine ID for a voice name (from configuration mapping)
        /// </summary>
        /// <param name="voiceName">Voice name (e.g., "amy", "jenny")</param>
        /// <returns>Engine ID, or empty string if not found</returns>
        std::wstring GetEngineIdForVoice(const std::wstring& voiceName) const;

        /// <summary>
        /// Perform health check on all engines
        /// </summary>
        /// <returns>S_OK if all engines are healthy</returns>
        HRESULT PerformHealthCheck();
        HRESULT ParseConfiguration(const std::wstring& jsonConfig);

    private:
        mutable std::mutex m_enginesMutex;
        std::map<std::wstring, std::unique_ptr<ITTSEngine>> m_engines;
        std::map<std::wstring, std::wstring> m_voiceToEngineMap;  // voice name -> engine ID
        
        // Configuration
        std::wstring m_configPath;
        
        // Helper methods
        std::string WStringToUTF8(const std::wstring& wstr) const;
        void LogMessage(const std::wstring& message) const;
        void LogError(const std::wstring& message, HRESULT hr = E_FAIL) const;
    };

    /// <summary>
    /// Singleton instance of the engine manager
    /// </summary>
    class TTSEngineManagerSingleton {
    public:
        static TTSEngineManager& GetInstance();

    private:
        static void InitializeSpdlog();
        static std::unique_ptr<TTSEngineManager> s_instance;
        static std::mutex s_instanceMutex;
    };

} // namespace NativeTTS
