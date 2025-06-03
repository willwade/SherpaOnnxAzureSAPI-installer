#pragma once

#include "ITTSEngine.h"
#include <windows.h>
#include <string>
#include <vector>
#include <memory>

// Forward declarations for SherpaOnnx C API
// We'll include the actual headers in the .cpp file
struct SherpaOnnxOfflineTts;
struct SherpaOnnxOfflineTtsConfig;
struct SherpaOnnxGeneratedAudio;

namespace NativeTTS {

    /// <summary>
    /// SherpaOnnx TTS Engine implementation using the C API
    /// Provides direct integration with SherpaOnnx without .NET overhead
    /// </summary>
    class SherpaOnnxEngine : public ITTSEngine {
    public:
        SherpaOnnxEngine();
        virtual ~SherpaOnnxEngine();

        // ITTSEngine implementation
        HRESULT Initialize(const std::wstring& config) override;
        HRESULT Generate(const std::wstring& text, std::vector<float>& samples, int& sampleRate) override;
        HRESULT Shutdown() override;
        bool IsInitialized() const override;
        std::wstring GetEngineInfo() const override;
        HRESULT GetSupportedFormat(int& sampleRate, int& channels, int& bitsPerSample) const override;

    private:
        // SherpaOnnx objects
        const SherpaOnnxOfflineTts* m_tts;
        SherpaOnnxOfflineTtsConfig* m_config;
        
        // Configuration
        std::wstring m_modelPath;
        std::wstring m_tokensPath;
        std::wstring m_lexiconPath;  // Optional
        std::wstring m_dataDir;      // Optional

        // UTF-8 versions for SherpaOnnx C API
        std::string m_modelPathUtf8;
        std::string m_tokensPathUtf8;
        std::string m_lexiconPathUtf8;
        std::string m_dataDirUtf8;
        
        // Audio parameters
        float m_noiseScale;
        float m_noiseScaleW;
        float m_lengthScale;
        int m_numThreads;
        int m_sampleRate;
        
        // State
        bool m_initialized;
        
        // Helper methods
        HRESULT ParseConfiguration(const std::wstring& jsonConfig);
        std::string WStringToUTF8(const std::wstring& wstr) const;
        std::wstring UTF8ToWString(const std::string& str) const;
        void LogMessage(const std::wstring& message) const;
        void LogError(const std::wstring& message, HRESULT hr = E_FAIL) const;
        
        // SherpaOnnx helper methods
        bool ValidateModelFiles() const;
        HRESULT CreateSherpaConfig();
        void CleanupSherpaObjects();
    };

    /// <summary>
    /// Configuration structure for SherpaOnnx engine
    /// Matches the JSON configuration format
    /// </summary>
    struct SherpaOnnxConfig {
        std::wstring modelPath;
        std::wstring tokensPath;
        std::wstring lexiconPath;    // Optional
        std::wstring dataDir;        // Optional
        
        float noiseScale = 0.667f;
        float noiseScaleW = 0.8f;
        float lengthScale = 1.0f;
        int numThreads = 1;
        bool debug = false;
        std::wstring provider = L"cpu";  // "cpu" or "cuda"
        
        // Parse from JSON string
        static HRESULT FromJson(const std::wstring& json, SherpaOnnxConfig& config);
        
        // Convert to JSON string
        std::wstring ToJson() const;
        
        // Validate configuration
        bool IsValid() const;
    };

} // namespace NativeTTS
