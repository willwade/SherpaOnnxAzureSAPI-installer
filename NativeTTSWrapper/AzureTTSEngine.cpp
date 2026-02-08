#include "stdafx.h"
#include "AzureTTSEngine.h"
#include <nlohmann/json.hpp>
#include <spdlog/spdlog.h>
#include <cmath>
#include <algorithm>

// Azure Speech SDK integration - TEMPORARILY DISABLED
// The Azure SDK headers are incomplete (missing inline implementations)
// TODO: Re-enable Azure SDK after fixing the installation or getting proper headers
// #include <speechapi_c.h>
// #include <speechapi_cxx.h>

using json = nlohmann::json;

namespace NativeTTS {

    AzureTTSEngine::AzureTTSEngine()
        : m_sampleRate(24000)
        , m_channels(1)
        , m_bitsPerSample(16)
        , m_initialized(false)
    {
        LogMessage(L"AzureTTSEngine created (stub implementation)");
    }

    AzureTTSEngine::~AzureTTSEngine() {
        Shutdown();
        LogMessage(L"AzureTTSEngine destroyed");
    }

    HRESULT AzureTTSEngine::Initialize(const std::wstring& config) {
        LogMessage(L"AzureTTSEngine::Initialize - Azure SDK integration not yet implemented");
        LogMessage(L"Config: " + config);
        return E_NOTIMPL;
    }

    HRESULT AzureTTSEngine::Generate(const std::wstring& text, std::vector<float>& samples, int& sampleRate) {
        LogMessage(L"AzureTTSEngine::Generate - Azure SDK integration not yet implemented");
        LogMessage(L"Text: " + text);
        return E_NOTIMPL;
    }

    HRESULT AzureTTSEngine::Shutdown() {
        LogMessage(L"AzureTTSEngine::Shutdown");
        m_initialized = false;
        return S_OK;
    }

    bool AzureTTSEngine::IsInitialized() const {
        return false;
    }

    std::wstring AzureTTSEngine::GetEngineInfo() const {
        return L"Azure TTS Engine (Not Yet Implemented)";
    }

    HRESULT AzureTTSEngine::GetSupportedFormat(int& sampleRate, int& channels, int& bitsPerSample) const {
        sampleRate = m_sampleRate;
        channels = m_channels;
        bitsPerSample = m_bitsPerSample;
        return S_OK;
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
            if (spdlog::default_logger() != nullptr) {
                std::string msg = WStringToUTF8(message);
                spdlog::info("[AzureTTSEngine] {}", msg);
            } else {
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

            if (spdlog::default_logger() != nullptr) {
                std::string msg = WStringToUTF8(fullMessage);
                spdlog::error("[AzureTTSEngine] {}", msg);
            } else {
                OutputDebugStringW((L"[AzureTTSEngine] " + fullMessage + L"\n").c_str());
            }
        }
        catch (...) {
            OutputDebugStringW((L"[AzureTTSEngine] ERROR: " + message + L"\n").c_str());
        }
    }

} // namespace NativeTTS
