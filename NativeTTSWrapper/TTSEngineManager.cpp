#include "stdafx.h"
#include "TTSEngineManager.h"
#include "SherpaOnnxEngine.h"
#include <nlohmann/json.hpp>
#include <spdlog/spdlog.h>
#include <spdlog/sinks/basic_file_sink.h>
#include <fstream>
#include <sstream>

using json = nlohmann::json;

namespace NativeTTS {

    // Singleton implementation
    std::unique_ptr<TTSEngineManager> TTSEngineManagerSingleton::s_instance;
    std::mutex TTSEngineManagerSingleton::s_instanceMutex;

    TTSEngineManager& TTSEngineManagerSingleton::GetInstance() {
        std::lock_guard<std::mutex> lock(s_instanceMutex);
        if (!s_instance) {
            // Initialize spdlog BEFORE creating the TTSEngineManager instance
            InitializeSpdlog();
            s_instance = std::make_unique<TTSEngineManager>();
        }
        return *s_instance;
    }

    void TTSEngineManagerSingleton::InitializeSpdlog() {
        try {
            // Only initialize if not already initialized
            if (spdlog::default_logger() == nullptr || spdlog::default_logger()->name() == "default") {
                // Create log directory if it doesn't exist
                CreateDirectoryA("C:\\OpenSpeech", nullptr);

                auto logger = spdlog::basic_logger_mt("tts_engine_manager", "C:\\OpenSpeech\\engine_manager.log");
                spdlog::set_default_logger(logger);
                spdlog::set_level(spdlog::level::info);
                spdlog::flush_on(spdlog::level::info);

                spdlog::info("spdlog initialized successfully");
            }
        }
        catch (const std::exception&) {
            // Fallback if logging fails
            OutputDebugStringA("Failed to initialize spdlog\n");
        }
    }

    TTSEngineManager::TTSEngineManager() {
        // spdlog should already be initialized by GetInstance()
        LogMessage(L"TTSEngineManager initialized");
    }

    TTSEngineManager::~TTSEngineManager() {
        ShutdownAllEngines();
        LogMessage(L"TTSEngineManager destroyed");
    }

    HRESULT TTSEngineManager::InitializeEngine(const std::wstring& engineId, EngineType type, const std::wstring& config) {
        std::lock_guard<std::mutex> lock(m_enginesMutex);
        
        LogMessage(L"Initializing engine: " + engineId);
        
        try {
            // Check if engine already exists
            if (m_engines.find(engineId) != m_engines.end()) {
                LogMessage(L"Engine already exists: " + engineId);
                return S_FALSE; // Already exists
            }

            // Create engine based on type
            std::unique_ptr<ITTSEngine> engine;
            switch (type) {
                case EngineType::SherpaOnnx:
                    engine = std::make_unique<SherpaOnnxEngine>();
                    break;
                case EngineType::Azure:
                    // Azure TTS engine removed - using AACSpeakHelper pipe service instead
                    LogError(L"Azure engine not available - use AACSpeakHelper pipe service", E_NOTIMPL);
                    return E_NOTIMPL;
                case EngineType::Mock:
                    // For testing - we'll implement a mock engine later
                    LogError(L"Mock engine not implemented yet", E_NOTIMPL);
                    return E_NOTIMPL;
                default:
                    LogError(L"Unknown engine type", E_INVALIDARG);
                    return E_INVALIDARG;
            }

            // Initialize the engine
            HRESULT hr = engine->Initialize(config);
            if (FAILED(hr)) {
                LogError(L"Failed to initialize engine: " + engineId, hr);
                return hr;
            }

            // Store the engine
            m_engines[engineId] = std::move(engine);
            
            LogMessage(L"Engine initialized successfully: " + engineId);
            return S_OK;
        }
        catch (const std::exception& ex) {
            std::string error = "Exception in InitializeEngine: ";
            error += ex.what();
            LogError(std::wstring(error.begin(), error.end()));
            return E_FAIL;
        }
    }

    ITTSEngine* TTSEngineManager::GetEngine(const std::wstring& engineId) {
        std::lock_guard<std::mutex> lock(m_enginesMutex);
        
        auto it = m_engines.find(engineId);
        if (it != m_engines.end()) {
            return it->second.get();
        }
        
        LogMessage(L"Engine not found: " + engineId);
        return nullptr;
    }

    HRESULT TTSEngineManager::ShutdownEngine(const std::wstring& engineId) {
        std::lock_guard<std::mutex> lock(m_enginesMutex);
        
        LogMessage(L"Shutting down engine: " + engineId);
        
        auto it = m_engines.find(engineId);
        if (it != m_engines.end()) {
            HRESULT hr = it->second->Shutdown();
            m_engines.erase(it);
            
            LogMessage(L"Engine shutdown complete: " + engineId);
            return hr;
        }
        
        return S_FALSE; // Engine not found
    }

    HRESULT TTSEngineManager::ShutdownAllEngines() {
        std::lock_guard<std::mutex> lock(m_enginesMutex);
        
        LogMessage(L"Shutting down all engines");
        
        HRESULT result = S_OK;
        for (auto& pair : m_engines) {
            HRESULT hr = pair.second->Shutdown();
            if (FAILED(hr)) {
                result = hr; // Keep track of any failures
            }
        }
        
        m_engines.clear();
        m_voiceToEngineMap.clear();
        
        LogMessage(L"All engines shutdown complete");
        return result;
    }

    std::vector<std::wstring> TTSEngineManager::GetEngineIds() const {
        std::lock_guard<std::mutex> lock(m_enginesMutex);
        
        std::vector<std::wstring> ids;
        for (const auto& pair : m_engines) {
            ids.push_back(pair.first);
        }
        return ids;
    }

    bool TTSEngineManager::IsEngineInitialized(const std::wstring& engineId) const {
        std::lock_guard<std::mutex> lock(m_enginesMutex);
        
        auto it = m_engines.find(engineId);
        if (it != m_engines.end()) {
            return it->second->IsInitialized();
        }
        return false;
    }

    HRESULT TTSEngineManager::LoadConfiguration(const std::wstring& configPath) {
        LogMessage(L"Loading configuration from: " + configPath);
        
        try {
            // Read configuration file
            std::ifstream file(configPath);
            if (!file.is_open()) {
                LogError(L"Failed to open configuration file: " + configPath);
                return E_FAIL;
            }

            json config;
            file >> config;
            file.close();

            // Parse and initialize engines
            std::string configStr = config.dump();
            std::wstring configWStr(configStr.begin(), configStr.end());
            return ParseConfiguration(configWStr);
        }
        catch (const std::exception& ex) {
            std::string error = "Exception loading configuration: ";
            error += ex.what();
            LogError(std::wstring(error.begin(), error.end()));
            return E_FAIL;
        }
    }

    std::wstring TTSEngineManager::GetEngineIdForVoice(const std::wstring& voiceName) const {
        std::lock_guard<std::mutex> lock(m_enginesMutex);
        
        auto it = m_voiceToEngineMap.find(voiceName);
        if (it != m_voiceToEngineMap.end()) {
            return it->second;
        }
        return L""; // Not found
    }

    HRESULT TTSEngineManager::PerformHealthCheck() {
        std::lock_guard<std::mutex> lock(m_enginesMutex);
        
        LogMessage(L"Performing health check on all engines");
        
        bool allHealthy = true;
        for (const auto& pair : m_engines) {
            if (!pair.second->IsInitialized()) {
                LogError(L"Engine not healthy: " + pair.first);
                allHealthy = false;
            }
        }
        
        LogMessage(allHealthy ? L"All engines healthy" : L"Some engines unhealthy");
        return allHealthy ? S_OK : S_FALSE;
    }

    HRESULT TTSEngineManager::ParseConfiguration(const std::wstring& jsonConfig) {
        try {
            // Convert wstring to string for JSON parsing
            std::string configStr(jsonConfig.begin(), jsonConfig.end());
            json config = json::parse(configStr);

            // Clear existing voice mappings
            m_voiceToEngineMap.clear();

            // Parse engines section
            if (config.contains("engines")) {
                for (const auto& [engineId, engineConfig] : config["engines"].items()) {
                    std::string typeStr = engineConfig["type"];
                    EngineType type = TTSEngineFactory::GetEngineTypeFromString(std::wstring(typeStr.begin(), typeStr.end()));
                    
                    std::string engineConfigStr = engineConfig["config"].dump();
                    std::wstring engineConfigWStr(engineConfigStr.begin(), engineConfigStr.end());
                    
                    std::wstring engineIdWStr(engineId.begin(), engineId.end());
                    
                    HRESULT hr = InitializeEngine(engineIdWStr, type, engineConfigWStr);
                    if (FAILED(hr)) {
                        LogError(L"Failed to initialize engine from config: " + engineIdWStr, hr);
                    }
                }
            }

            // Parse voices section
            if (config.contains("voices")) {
                for (const auto& [voiceName, engineId] : config["voices"].items()) {
                    std::wstring voiceNameWStr(voiceName.begin(), voiceName.end());
                    std::wstring engineIdWStr(engineId.begin(), engineId.end());
                    m_voiceToEngineMap[voiceNameWStr] = engineIdWStr;
                }
            }

            LogMessage(L"Configuration parsed successfully");
            return S_OK;
        }
        catch (const std::exception& ex) {
            std::string error = "Exception parsing configuration: ";
            error += ex.what();
            LogError(std::wstring(error.begin(), error.end()));
            return E_FAIL;
        }
    }

    std::string TTSEngineManager::WStringToUTF8(const std::wstring& wstr) const {
        if (wstr.empty()) return std::string();
        int size_needed = WideCharToMultiByte(CP_UTF8, 0, &wstr[0], (int)wstr.size(), nullptr, 0, nullptr, nullptr);
        std::string strTo(size_needed, 0);
        WideCharToMultiByte(CP_UTF8, 0, &wstr[0], (int)wstr.size(), &strTo[0], size_needed, nullptr, nullptr);
        return strTo;
    }

    void TTSEngineManager::LogMessage(const std::wstring& message) const {
        try {
            // Check if spdlog is initialized before using it
            if (spdlog::default_logger() != nullptr) {
                std::string msg = WStringToUTF8(message);
                spdlog::info("[TTSEngineManager] {}", msg);
            } else {
                // Fallback to OutputDebugString if spdlog not initialized
                OutputDebugStringW((L"[TTSEngineManager] " + message + L"\n").c_str());
            }
        }
        catch (...) {
            // Fallback logging
            OutputDebugStringW((L"[TTSEngineManager] " + message + L"\n").c_str());
        }
    }

    void TTSEngineManager::LogError(const std::wstring& message, HRESULT hr) const {
        try {
            std::wstring fullMessage = L"ERROR: " + message;
            if (hr != E_FAIL) {
                fullMessage += L" (HRESULT: 0x" + std::to_wstring(hr) + L")";
            }

            // Check if spdlog is initialized before using it
            if (spdlog::default_logger() != nullptr) {
                std::string msg = WStringToUTF8(fullMessage);
                spdlog::error("[TTSEngineManager] {}", msg);
            } else {
                // Fallback to OutputDebugString if spdlog not initialized
                OutputDebugStringW((L"[TTSEngineManager] " + fullMessage + L"\n").c_str());
            }
        }
        catch (...) {
            // Fallback logging
            OutputDebugStringW((L"[TTSEngineManager] ERROR: " + message + L"\n").c_str());
        }
    }

} // namespace NativeTTS
