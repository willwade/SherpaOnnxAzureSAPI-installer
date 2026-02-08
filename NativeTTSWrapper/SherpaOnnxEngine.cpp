#include "stdafx.h"
#include "SherpaOnnxEngine.h"
#include "sherpa-onnx/c-api/c-api.h"
#include <nlohmann/json.hpp>
#include <spdlog/spdlog.h>
#include <fstream>
#include <cmath>
#include <algorithm>

using json = nlohmann::json;

namespace NativeTTS {

    SherpaOnnxEngine::SherpaOnnxEngine()
        : m_tts(nullptr)
        , m_config(nullptr)
        , m_noiseScale(0.667f)
        , m_noiseScaleW(0.8f)
        , m_lengthScale(1.0f)
        , m_numThreads(1)
        , m_sampleRate(22050)
        , m_initialized(false)
    {
        LogMessage(L"SherpaOnnxEngine created");
    }

    SherpaOnnxEngine::~SherpaOnnxEngine() {
        Shutdown();
        LogMessage(L"SherpaOnnxEngine destroyed");
    }

    HRESULT SherpaOnnxEngine::Initialize(const std::wstring& config) {
        LogMessage(L"Initializing SherpaOnnxEngine with config");

        try {
            // Parse configuration
            HRESULT hr = ParseConfiguration(config);
            if (FAILED(hr)) {
                LogError(L"Failed to parse SherpaOnnx configuration", hr);
                return hr;
            }

            // Validate model files exist
            if (!ValidateModelFiles()) {
                LogError(L"SherpaOnnx model files validation failed");
                return E_FAIL;
            }

            // Create SherpaOnnx configuration
            hr = CreateSherpaConfig();
            if (FAILED(hr)) {
                LogError(L"Failed to create SherpaOnnx config", hr);
                return hr;
            }

            // Initialize SherpaOnnx TTS
            m_tts = SherpaOnnxCreateOfflineTts(m_config);
            if (!m_tts) {
                LogError(L"Failed to create SherpaOnnx TTS instance");
                return E_FAIL;
            }

            // Get actual sample rate from SherpaOnnx
            m_sampleRate = SherpaOnnxOfflineTtsSampleRate(m_tts);

            m_initialized = true;
            LogMessage(L"SherpaOnnxEngine initialized successfully");
            LogMessage(L"Sample rate: " + std::to_wstring(m_sampleRate));

            return S_OK;
        }
        catch (const std::exception& ex) {
            std::string error = "Exception in SherpaOnnxEngine::Initialize: ";
            error += ex.what();
            LogError(std::wstring(error.begin(), error.end()));
            return E_FAIL;
        }
    }

    HRESULT SherpaOnnxEngine::Generate(const std::wstring& text, std::vector<float>& samples, int& sampleRate) {
        if (!m_initialized || !m_tts) {
            LogError(L"SherpaOnnxEngine not initialized");
            return E_FAIL;
        }

        try {
            LogMessage(L"Generating audio for text: " + text);

            // Convert text to UTF-8
            std::string utf8Text = WStringToUTF8(text);

            // Generate audio using SherpaOnnx
            const SherpaOnnxGeneratedAudio* audio = SherpaOnnxOfflineTtsGenerate(
                m_tts, utf8Text.c_str(), 0, 1.0f);

            if (!audio || !audio->samples || audio->n <= 0) {
                LogError(L"SherpaOnnx generation failed");
                return E_FAIL;
            }

            // Copy samples to output vector
            samples.resize(audio->n);
            std::copy(audio->samples, audio->samples + audio->n, samples.begin());
            sampleRate = audio->sample_rate;

            // Clean up generated audio
            SherpaOnnxDestroyOfflineTtsGeneratedAudio(audio);

            LogMessage(L"Generated " + std::to_wstring(samples.size()) + L" samples at " + std::to_wstring(sampleRate) + L"Hz");
            return S_OK;
        }
        catch (const std::exception& ex) {
            std::string error = "Exception in SherpaOnnxEngine::Generate: ";
            error += ex.what();
            LogError(std::wstring(error.begin(), error.end()));
            return E_FAIL;
        }
    }

    HRESULT SherpaOnnxEngine::Shutdown() {
        LogMessage(L"Shutting down SherpaOnnxEngine");

        CleanupSherpaObjects();
        m_initialized = false;

        LogMessage(L"SherpaOnnxEngine shutdown complete");
        return S_OK;
    }

    bool SherpaOnnxEngine::IsInitialized() const {
        return m_initialized;
    }

    std::wstring SherpaOnnxEngine::GetEngineInfo() const {
        return L"SherpaOnnx TTS Engine v1.12.10";
    }

    HRESULT SherpaOnnxEngine::GetSupportedFormat(int& sampleRate, int& channels, int& bitsPerSample) const {
        sampleRate = m_sampleRate;
        channels = 1; // Mono
        bitsPerSample = 16; // 16-bit PCM
        return S_OK;
    }

    HRESULT SherpaOnnxEngine::ParseConfiguration(const std::wstring& jsonConfig) {
        try {
            std::string configStr = WStringToUTF8(jsonConfig);
            json config = json::parse(configStr);

            // Parse model paths
            if (config.contains("modelPath")) {
                m_modelPath = UTF8ToWString(config["modelPath"]);
            }
            if (config.contains("tokensPath")) {
                m_tokensPath = UTF8ToWString(config["tokensPath"]);
            }
            if (config.contains("lexiconPath")) {
                m_lexiconPath = UTF8ToWString(config["lexiconPath"]);
            }
            if (config.contains("dataDir")) {
                m_dataDir = UTF8ToWString(config["dataDir"]);
            }

            // Parse audio parameters
            if (config.contains("noiseScale")) {
                m_noiseScale = config["noiseScale"];
            }
            if (config.contains("noiseScaleW")) {
                m_noiseScaleW = config["noiseScaleW"];
            }
            if (config.contains("lengthScale")) {
                m_lengthScale = config["lengthScale"];
            }
            if (config.contains("numThreads")) {
                m_numThreads = config["numThreads"];
            }

            LogMessage(L"Configuration parsed successfully");
            LogMessage(L"Model path: " + m_modelPath);
            LogMessage(L"Tokens path: " + m_tokensPath);

            return S_OK;
        }
        catch (const std::exception& ex) {
            std::string error = "Exception parsing SherpaOnnx configuration: ";
            error += ex.what();
            LogError(std::wstring(error.begin(), error.end()));
            return E_FAIL;
        }
    }

    std::string SherpaOnnxEngine::WStringToUTF8(const std::wstring& wstr) const {
        if (wstr.empty()) return std::string();
        int size_needed = WideCharToMultiByte(CP_UTF8, 0, &wstr[0], (int)wstr.size(), nullptr, 0, nullptr, nullptr);
        std::string strTo(size_needed, 0);
        WideCharToMultiByte(CP_UTF8, 0, &wstr[0], (int)wstr.size(), &strTo[0], size_needed, nullptr, nullptr);
        return strTo;
    }

    std::wstring SherpaOnnxEngine::UTF8ToWString(const std::string& str) const {
        if (str.empty()) return std::wstring();
        int size_needed = MultiByteToWideChar(CP_UTF8, 0, &str[0], (int)str.size(), nullptr, 0);
        std::wstring wstrTo(size_needed, 0);
        MultiByteToWideChar(CP_UTF8, 0, &str[0], (int)str.size(), &wstrTo[0], size_needed);
        return wstrTo;
    }

    void SherpaOnnxEngine::LogMessage(const std::wstring& message) const {
        try {
            if (spdlog::default_logger() != nullptr) {
                std::string msg = WStringToUTF8(message);
                spdlog::info("[SherpaOnnxEngine] {}", msg);
            } else {
                OutputDebugStringW((L"[SherpaOnnxEngine] " + message + L"\n").c_str());
            }
        }
        catch (...) {
            OutputDebugStringW((L"[SherpaOnnxEngine] " + message + L"\n").c_str());
        }
    }

    void SherpaOnnxEngine::LogError(const std::wstring& message, HRESULT hr) const {
        try {
            std::wstring fullMessage = L"ERROR: " + message;
            if (hr != E_FAIL) {
                fullMessage += L" (HRESULT: 0x" + std::to_wstring(hr) + L")";
            }

            if (spdlog::default_logger() != nullptr) {
                std::string msg = WStringToUTF8(fullMessage);
                spdlog::error("[SherpaOnnxEngine] {}", msg);
            } else {
                OutputDebugStringW((L"[SherpaOnnxEngine] " + fullMessage + L"\n").c_str());
            }
        }
        catch (...) {
            OutputDebugStringW((L"[SherpaOnnxEngine] ERROR: " + message + L"\n").c_str());
        }
    }

    bool SherpaOnnxEngine::ValidateModelFiles() const {
        // Check if model file exists
        std::ifstream modelFile(m_modelPath);
        if (!modelFile.good()) {
            LogError(L"Model file not found: " + m_modelPath);
            return false;
        }

        // Check if tokens file exists
        std::ifstream tokensFile(m_tokensPath);
        if (!tokensFile.good()) {
            LogError(L"Tokens file not found: " + m_tokensPath);
            return false;
        }

        LogMessage(L"Model files validated successfully");
        return true;
    }

    HRESULT SherpaOnnxEngine::CreateSherpaConfig() {
        LogMessage(L"Creating SherpaOnnx configuration");

        try {
            // Allocate configuration structure
            m_config = new SherpaOnnxOfflineTtsConfig();
            memset(m_config, 0, sizeof(SherpaOnnxOfflineTtsConfig));

            // Convert paths to UTF-8
            std::string modelPathUtf8 = WStringToUTF8(m_modelPath);
            std::string tokensPathUtf8 = WStringToUTF8(m_tokensPath);
            std::string lexiconPathUtf8 = WStringToUTF8(m_lexiconPath);
            std::string dataDirUtf8 = WStringToUTF8(m_dataDir);

            // Store strings (need to keep them alive)
            m_modelPathUtf8 = modelPathUtf8;
            m_tokensPathUtf8 = tokensPathUtf8;
            m_lexiconPathUtf8 = lexiconPathUtf8;
            m_dataDirUtf8 = dataDirUtf8;

            // Configure VITS model (most common for Piper models)
            m_config->model.vits.model = m_modelPathUtf8.c_str();
            m_config->model.vits.tokens = m_tokensPathUtf8.c_str();
            m_config->model.vits.lexicon = m_lexiconPathUtf8.empty() ? nullptr : m_lexiconPathUtf8.c_str();
            m_config->model.vits.data_dir = m_dataDirUtf8.empty() ? nullptr : m_dataDirUtf8.c_str();
            m_config->model.vits.noise_scale = m_noiseScale;
            m_config->model.vits.noise_scale_w = m_noiseScaleW;
            m_config->model.vits.length_scale = m_lengthScale;

            // Configure general model settings
            m_config->model.num_threads = m_numThreads;
            m_config->model.debug = 0; // Disable debug for performance
            m_config->model.provider = "cpu"; // Use CPU provider

            // Configure TTS settings
            m_config->rule_fsts = nullptr;
            m_config->rule_fars = nullptr;
            m_config->max_num_sentences = 1; // Process one sentence at a time
            m_config->silence_scale = 1.0f;

            LogMessage(L"SherpaOnnx configuration created successfully");
            LogMessage(L"Model: " + m_modelPath);
            LogMessage(L"Tokens: " + m_tokensPath);
            if (!m_lexiconPath.empty()) {
                LogMessage(L"Lexicon: " + m_lexiconPath);
            }

            return S_OK;
        }
        catch (const std::exception& ex) {
            std::string error = "Exception creating SherpaOnnx config: ";
            error += ex.what();
            LogError(std::wstring(error.begin(), error.end()));
            return E_FAIL;
        }
    }

    void SherpaOnnxEngine::CleanupSherpaObjects() {
        if (m_tts) {
            SherpaOnnxDestroyOfflineTts(m_tts);
            m_tts = nullptr;
        }
        if (m_config) {
            delete m_config;
            m_config = nullptr;
        }

        // Clear UTF-8 strings
        m_modelPathUtf8.clear();
        m_tokensPathUtf8.clear();
        m_lexiconPathUtf8.clear();
        m_dataDirUtf8.clear();
    }

} // namespace NativeTTS
