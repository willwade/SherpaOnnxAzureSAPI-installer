#pragma once

#include "ITTSEngine.h"
#include <windows.h>
#include <string>
#include <vector>
#include <memory>

// Forward declarations for Azure Speech SDK
// We'll include the actual headers in the .cpp file
namespace Microsoft {
    namespace CognitiveServices {
        namespace Speech {
            class SpeechConfig;
            class SpeechSynthesizer;
            class SpeechSynthesisResult;
            class AudioConfig;
        }
    }
}

namespace NativeTTS {

    /// <summary>
    /// Azure TTS Engine implementation using the C++ Speech SDK
    /// Provides direct integration with Azure Cognitive Services
    /// </summary>
    class AzureTTSEngine : public ITTSEngine {
    public:
        AzureTTSEngine();
        virtual ~AzureTTSEngine();

        // ITTSEngine implementation
        HRESULT Initialize(const std::wstring& config) override;
        HRESULT Generate(const std::wstring& text, std::vector<float>& samples, int& sampleRate) override;
        HRESULT Shutdown() override;
        bool IsInitialized() const override;
        std::wstring GetEngineInfo() const override;
        HRESULT GetSupportedFormat(int& sampleRate, int& channels, int& bitsPerSample) const override;

    private:
        // Azure Speech SDK objects
        std::shared_ptr<Microsoft::CognitiveServices::Speech::SpeechConfig> m_speechConfig;
        std::shared_ptr<Microsoft::CognitiveServices::Speech::SpeechSynthesizer> m_synthesizer;
        std::shared_ptr<Microsoft::CognitiveServices::Speech::AudioConfig> m_audioConfig;
        
        // Configuration
        std::wstring m_subscriptionKey;
        std::wstring m_region;
        std::wstring m_voiceName;
        std::wstring m_language;
        std::wstring m_style;        // Optional
        std::wstring m_role;         // Optional
        
        // Audio parameters
        int m_sampleRate;
        int m_channels;
        int m_bitsPerSample;
        
        // State
        bool m_initialized;
        
        // Helper methods
        HRESULT ParseConfiguration(const std::wstring& jsonConfig);
        std::string WStringToUTF8(const std::wstring& wstr) const;
        std::wstring UTF8ToWString(const std::string& str) const;
        void LogMessage(const std::wstring& message) const;
        void LogError(const std::wstring& message, HRESULT hr = E_FAIL) const;
        
        // Azure-specific helper methods
        HRESULT CreateSpeechConfig();
        HRESULT CreateSynthesizer();
        std::wstring BuildSSML(const std::wstring& text) const;
        HRESULT ProcessSynthesisResult(
            std::shared_ptr<Microsoft::CognitiveServices::Speech::SpeechSynthesisResult> result,
            std::vector<float>& samples,
            int& sampleRate) const;
        
        // Audio format conversion
        HRESULT ConvertAudioToFloat(const uint8_t* audioData, size_t dataSize, 
                                   std::vector<float>& samples) const;
    };

    /// <summary>
    /// Configuration structure for Azure TTS engine
    /// Matches the JSON configuration format
    /// </summary>
    struct AzureTTSConfig {
        std::wstring subscriptionKey;
        std::wstring region;
        std::wstring voiceName;
        std::wstring language = L"en-US";
        std::wstring style;          // Optional
        std::wstring role;           // Optional
        
        // Audio format preferences
        int sampleRate = 24000;      // Azure default: 24kHz
        int channels = 1;            // Mono
        int bitsPerSample = 16;      // 16-bit PCM
        
        // Parse from JSON string
        static HRESULT FromJson(const std::wstring& json, AzureTTSConfig& config);
        
        // Convert to JSON string
        std::wstring ToJson() const;
        
        // Validate configuration
        bool IsValid() const;
    };

} // namespace NativeTTS
