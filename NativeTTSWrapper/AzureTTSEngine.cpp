#include "stdafx.h"
#include "AzureTTSEngine.h"
#include <nlohmann/json.hpp>
#include <spdlog/spdlog.h>
#include <cmath>
#include <algorithm>

// Real Azure Speech SDK integration
#include <speechapi_cxx.h>

using namespace Microsoft::CognitiveServices::Speech;
using namespace Microsoft::CognitiveServices::Speech::Audio;

using json = nlohmann::json;

namespace NativeTTS {

    AzureTTSEngine::AzureTTSEngine()
        : m_sampleRate(24000)
        , m_channels(1)
        , m_bitsPerSample(16)
        , m_initialized(false)
    {
        LogMessage(L"AzureTTSEngine created");
    }

    AzureTTSEngine::~AzureTTSEngine() {
        Shutdown();
        LogMessage(L"AzureTTSEngine destroyed");
    }

    HRESULT AzureTTSEngine::Initialize(const std::wstring& config) {
        LogMessage(L"Initializing AzureTTSEngine with config");
        
        try {
            // Parse configuration
            HRESULT hr = ParseConfiguration(config);
            if (FAILED(hr)) {
                LogError(L"Failed to parse Azure TTS configuration", hr);
                return hr;
            }

            // Validate required parameters
            if (m_subscriptionKey.empty()) {
                LogError(L"Azure subscription key is required");
                return E_INVALIDARG;
            }
            if (m_region.empty()) {
                LogError(L"Azure region is required");
                return E_INVALIDARG;
            }
            if (m_voiceName.empty()) {
                LogError(L"Azure voice name is required");
                return E_INVALIDARG;
            }

            // Create speech configuration (mock for now)
            hr = CreateSpeechConfig();
            if (FAILED(hr)) {
                LogError(L"Failed to create Azure speech config", hr);
                return hr;
            }

            // Create synthesizer (mock for now)
            hr = CreateSynthesizer();
            if (FAILED(hr)) {
                LogError(L"Failed to create Azure synthesizer", hr);
                return hr;
            }

            m_initialized = true;
            LogMessage(L"AzureTTSEngine initialized successfully");
            LogMessage(L"Voice: " + m_voiceName);
            LogMessage(L"Region: " + m_region);
            LogMessage(L"Sample rate: " + std::to_wstring(m_sampleRate));
            
            return S_OK;
        }
        catch (const std::exception& ex) {
            std::string error = "Exception in AzureTTSEngine::Initialize: ";
            error += ex.what();
            LogError(std::wstring(error.begin(), error.end()));
            return E_FAIL;
        }
    }

    HRESULT AzureTTSEngine::Generate(const std::wstring& text, std::vector<float>& samples, int& sampleRate) {
        if (!m_initialized) {
            LogError(L"AzureTTSEngine not initialized");
            return E_FAIL;
        }

        try {
            LogMessage(L"Generating Azure TTS audio for text: " + text);

            if (!m_synthesizer) {
                LogError(L"Azure synthesizer not created");
                return E_FAIL;
            }

            // Build SSML
            std::wstring ssml = BuildSSML(text);
            LogMessage(L"Generated SSML: " + ssml);

            // Convert SSML to UTF-8 for Azure SDK
            std::string ssmlUtf8 = WStringToUTF8(ssml);

            // Perform synthesis
            auto result = m_synthesizer->SpeakSsmlAsync(ssmlUtf8).get();
            if (!result) {
                LogError(L"Azure synthesis failed - no result");
                return E_FAIL;
            }

            // Check result reason
            if (result->Reason == ResultReason::SynthesizingAudioCompleted) {
                LogMessage(L"Azure synthesis completed successfully");

                // Process the synthesis result
                HRESULT hr = ProcessSynthesisResult(result, samples, sampleRate);
                if (FAILED(hr)) {
                    LogError(L"Failed to process Azure synthesis result", hr);
                    return hr;
                }

                LogMessage(L"Generated " + std::to_wstring(samples.size()) + L" samples at " + std::to_wstring(sampleRate) + L"Hz");
                return S_OK;
            }
            else if (result->Reason == ResultReason::Canceled) {
                auto cancellation = SpeechSynthesisCancellationDetails::FromResult(result);
                std::string errorDetails = "Azure synthesis canceled: " + cancellation->ErrorDetails;
                LogError(std::wstring(errorDetails.begin(), errorDetails.end()));
                return E_FAIL;
            }
            else {
                LogError(L"Azure synthesis failed with unknown reason");
                return E_FAIL;
            }
        }
        catch (const std::exception& ex) {
            std::string error = "Exception in AzureTTSEngine::Generate: ";
            error += ex.what();
            LogError(std::wstring(error.begin(), error.end()));
            return E_FAIL;
        }
    }

    HRESULT AzureTTSEngine::Shutdown() {
        LogMessage(L"Shutting down AzureTTSEngine");
        
        // Clean up Azure SDK objects (when we implement them)
        m_speechConfig.reset();
        m_synthesizer.reset();
        m_audioConfig.reset();
        
        m_initialized = false;
        
        LogMessage(L"AzureTTSEngine shutdown complete");
        return S_OK;
    }

    bool AzureTTSEngine::IsInitialized() const {
        return m_initialized;
    }

    std::wstring AzureTTSEngine::GetEngineInfo() const {
        return L"Azure TTS Engine (Mock Implementation) - Voice: " + m_voiceName + L", Region: " + m_region;
    }

    HRESULT AzureTTSEngine::GetSupportedFormat(int& sampleRate, int& channels, int& bitsPerSample) const {
        sampleRate = m_sampleRate;
        channels = m_channels;
        bitsPerSample = m_bitsPerSample;
        return S_OK;
    }

    HRESULT AzureTTSEngine::ParseConfiguration(const std::wstring& jsonConfig) {
        try {
            std::string configStr = WStringToUTF8(jsonConfig);
            json config = json::parse(configStr);

            // Parse required parameters
            if (config.contains("subscriptionKey")) {
                m_subscriptionKey = UTF8ToWString(config["subscriptionKey"]);
            }
            if (config.contains("region")) {
                m_region = UTF8ToWString(config["region"]);
            }
            if (config.contains("voiceName")) {
                m_voiceName = UTF8ToWString(config["voiceName"]);
            }

            // Parse optional parameters
            if (config.contains("language")) {
                m_language = UTF8ToWString(config["language"]);
            } else {
                m_language = L"en-US"; // Default
            }
            
            if (config.contains("style")) {
                m_style = UTF8ToWString(config["style"]);
            }
            if (config.contains("role")) {
                m_role = UTF8ToWString(config["role"]);
            }

            // Parse audio format parameters
            if (config.contains("sampleRate")) {
                m_sampleRate = config["sampleRate"];
            }
            if (config.contains("channels")) {
                m_channels = config["channels"];
            }
            if (config.contains("bitsPerSample")) {
                m_bitsPerSample = config["bitsPerSample"];
            }

            LogMessage(L"Azure TTS configuration parsed successfully");
            return S_OK;
        }
        catch (const std::exception& ex) {
            std::string error = "Exception parsing Azure TTS configuration: ";
            error += ex.what();
            LogError(std::wstring(error.begin(), error.end()));
            return E_FAIL;
        }
    }

    std::string AzureTTSEngine::WStringToUTF8(const std::wstring& wstr) const {
        if (wstr.empty()) return std::string();
        int size_needed = WideCharToMultiByte(CP_UTF8, 0, &wstr[0], (int)wstr.size(), nullptr, 0, nullptr, nullptr);
        std::string strTo(size_needed, 0);
        WideCharToMultiByte(CP_UTF8, 0, &wstr[0], (int)wstr.size(), &strTo[0], size_needed, nullptr, nullptr);
        return strTo;
    }

    std::wstring AzureTTSEngine::UTF8ToWString(const std::string& str) const {
        if (str.empty()) return std::wstring();
        int size_needed = MultiByteToWideChar(CP_UTF8, 0, &str[0], (int)str.size(), nullptr, 0);
        std::wstring wstrTo(size_needed, 0);
        MultiByteToWideChar(CP_UTF8, 0, &str[0], (int)str.size(), &wstrTo[0], size_needed);
        return wstrTo;
    }

    void AzureTTSEngine::LogMessage(const std::wstring& message) const {
        try {
            // Check if spdlog is initialized before using it
            if (spdlog::default_logger() != nullptr) {
                std::string msg = WStringToUTF8(message);
                spdlog::info("[AzureTTSEngine] {}", msg);
            } else {
                // Fallback to OutputDebugString if spdlog not initialized
                OutputDebugStringW((L"[AzureTTSEngine] " + message + L"\n").c_str());
            }
        }
        catch (...) {
            OutputDebugStringW((L"[AzureTTSEngine] " + message + L"\n").c_str());
        }
    }

    void AzureTTSEngine::LogError(const std::wstring& message, HRESULT hr) const {
        try {
            std::wstring fullMessage = L"ERROR: " + message;
            if (hr != E_FAIL) {
                fullMessage += L" (HRESULT: 0x" + std::to_wstring(hr) + L")";
            }

            // Check if spdlog is initialized before using it
            if (spdlog::default_logger() != nullptr) {
                std::string msg = WStringToUTF8(fullMessage);
                spdlog::error("[AzureTTSEngine] {}", msg);
            } else {
                // Fallback to OutputDebugString if spdlog not initialized
                OutputDebugStringW((L"[AzureTTSEngine] " + fullMessage + L"\n").c_str());
            }
        }
        catch (...) {
            OutputDebugStringW((L"[AzureTTSEngine] ERROR: " + message + L"\n").c_str());
        }
    }

    HRESULT AzureTTSEngine::CreateSpeechConfig() {
        LogMessage(L"Creating Azure speech configuration");

        try {
            // Convert subscription key and region to UTF-8
            std::string subscriptionKey = WStringToUTF8(m_subscriptionKey);
            std::string region = WStringToUTF8(m_region);

            // Create speech configuration
            m_speechConfig = SpeechConfig::FromSubscription(subscriptionKey, region);
            if (!m_speechConfig) {
                LogError(L"Failed to create Azure SpeechConfig");
                return E_FAIL;
            }

            // Set voice name
            std::string voiceName = WStringToUTF8(m_voiceName);
            m_speechConfig->SetSpeechSynthesisVoiceName(voiceName);

            // Set output format to PCM 24kHz 16-bit mono
            m_speechConfig->SetSpeechSynthesisOutputFormat(SpeechSynthesisOutputFormat::Riff24Khz16BitMonoPcm);

            LogMessage(L"Azure speech configuration created successfully");
            return S_OK;
        }
        catch (const std::exception& ex) {
            std::string error = "Exception creating Azure speech config: ";
            error += ex.what();
            LogError(std::wstring(error.begin(), error.end()));
            return E_FAIL;
        }
    }

    HRESULT AzureTTSEngine::CreateSynthesizer() {
        LogMessage(L"Creating Azure synthesizer");

        try {
            if (!m_speechConfig) {
                LogError(L"Speech config not created");
                return E_FAIL;
            }

            // Create audio configuration for in-memory synthesis
            m_audioConfig = AudioConfig::FromDefaultSpeakerOutput();
            if (!m_audioConfig) {
                LogError(L"Failed to create Azure AudioConfig");
                return E_FAIL;
            }

            // Create speech synthesizer
            m_synthesizer = SpeechSynthesizer::FromConfig(m_speechConfig, m_audioConfig);
            if (!m_synthesizer) {
                LogError(L"Failed to create Azure SpeechSynthesizer");
                return E_FAIL;
            }

            LogMessage(L"Azure synthesizer created successfully");
            return S_OK;
        }
        catch (const std::exception& ex) {
            std::string error = "Exception creating Azure synthesizer: ";
            error += ex.what();
            LogError(std::wstring(error.begin(), error.end()));
            return E_FAIL;
        }
    }

    std::wstring AzureTTSEngine::BuildSSML(const std::wstring& text) const {
        std::wstring ssml = L"<speak version='1.0' xmlns='http://www.w3.org/2001/10/synthesis' xmlns:mstts='https://www.w3.org/2001/mstts' xml:lang='";
        ssml += m_language;
        ssml += L"'><voice name='";
        ssml += m_voiceName;
        ssml += L"'>";

        // Add style if specified
        if (!m_style.empty()) {
            ssml += L"<mstts:express-as style='";
            ssml += m_style;
            ssml += L"'>";
        }

        // Add role if specified
        if (!m_role.empty()) {
            ssml += L"<mstts:express-as role='";
            ssml += m_role;
            ssml += L"'>";
        }

        ssml += text;

        // Close tags in reverse order
        if (!m_role.empty()) {
            ssml += L"</mstts:express-as>";
        }
        if (!m_style.empty()) {
            ssml += L"</mstts:express-as>";
        }

        ssml += L"</voice></speak>";

        return ssml;
    }

    HRESULT AzureTTSEngine::ProcessSynthesisResult(
        std::shared_ptr<Microsoft::CognitiveServices::Speech::SpeechSynthesisResult> result,
        std::vector<float>& samples,
        int& sampleRate) const {

        try {
            LogMessage(L"Processing Azure synthesis result");

            // Get audio data from result
            auto audioData = result->GetAudioData();
            if (audioData.empty()) {
                LogError(L"Azure synthesis result contains no audio data");
                return E_FAIL;
            }

            LogMessage(L"Received " + std::to_wstring(audioData.size()) + L" bytes of audio data from Azure");

            // Azure returns WAV format with header, we need to extract PCM data
            const uint8_t* data = audioData.data();
            size_t dataSize = audioData.size();

            // Skip WAV header (44 bytes) and convert PCM to float
            if (dataSize < 44) {
                LogError(L"Audio data too small to contain WAV header");
                return E_FAIL;
            }

            // Extract sample rate from WAV header (bytes 24-27)
            sampleRate = *reinterpret_cast<const uint32_t*>(data + 24);
            LogMessage(L"Audio sample rate: " + std::to_wstring(sampleRate) + L"Hz");

            // Convert PCM data to float samples
            HRESULT hr = ConvertAudioToFloat(data + 44, dataSize - 44, samples);
            if (FAILED(hr)) {
                LogError(L"Failed to convert Azure audio to float", hr);
                return hr;
            }

            LogMessage(L"Converted to " + std::to_wstring(samples.size()) + L" float samples");
            return S_OK;
        }
        catch (const std::exception& ex) {
            std::string error = "Exception processing Azure synthesis result: ";
            error += ex.what();
            LogError(std::wstring(error.begin(), error.end()));
            return E_FAIL;
        }
    }

    HRESULT AzureTTSEngine::ConvertAudioToFloat(const uint8_t* audioData, size_t dataSize,
                                               std::vector<float>& samples) const {
        try {
            LogMessage(L"Converting Azure audio to float");

            // Azure typically returns 16-bit PCM data
            if (dataSize % 2 != 0) {
                LogError(L"Audio data size is not aligned for 16-bit samples");
                return E_INVALIDARG;
            }

            size_t numSamples = dataSize / 2; // 16-bit = 2 bytes per sample
            samples.resize(numSamples);

            const int16_t* pcmData = reinterpret_cast<const int16_t*>(audioData);

            // Convert 16-bit PCM to normalized float (-1.0 to 1.0)
            for (size_t i = 0; i < numSamples; ++i) {
                samples[i] = static_cast<float>(pcmData[i]) / 32768.0f;
            }

            LogMessage(L"Converted " + std::to_wstring(numSamples) + L" PCM samples to float");
            return S_OK;
        }
        catch (const std::exception& ex) {
            std::string error = "Exception converting Azure audio to float: ";
            error += ex.what();
            LogError(std::wstring(error.begin(), error.end()));
            return E_FAIL;
        }
    }

} // namespace NativeTTS
