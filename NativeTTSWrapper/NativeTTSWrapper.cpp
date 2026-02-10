#define NOMINMAX
#include "stdafx.h"
#include "resource.h"
#include "NativeTTSWrapper_i.h"
#include "dllmain.h"
#include "NativeTTSWrapper.h"
#include <windows.h>
#include <sapi.h>
#include <sphelper.h>
#include <comdef.h>
#include <iostream>
#include <fstream>
#include <string>
#include <memory>
#include <process.h>
#include <cstdint>
#include <vector>
#include <sstream>
#include <algorithm>

// Include SherpaOnnx C API for direct fallback
extern "C" {
    #include "sherpa-onnx-c-api.h"
}

// Include TTSEngineManager for native engine support
#include "ITTSEngine.h"
#include "TTSEngineManager.h"

// Implementation of CNativeTTSWrapper

CNativeTTSWrapper::CNativeTTSWrapper() : m_engineInitialized(false), m_actualSampleRate(16000)
{
    LogMessage(L"CNativeTTSWrapper constructor called");
}

CNativeTTSWrapper::~CNativeTTSWrapper()
{
    LogMessage(L"CNativeTTSWrapper destructor called");
}

// ISpTTSEngine implementation
STDMETHODIMP CNativeTTSWrapper::Speak(
    DWORD dwSpeakFlags,
    REFGUID rguidFormatId,
    const WAVEFORMATEX* pWaveFormatEx,
    const SPVTEXTFRAG* pTextFragList,
    ISpTTSEngineSite* pOutputSite)
    {
        LogMessage(L"*** NATIVE SPEAK METHOD CALLED ***");
        
        if (!pTextFragList || !pOutputSite)
        {
            LogMessage(L"Invalid parameters to Speak method");
            return E_INVALIDARG;
        }

        try
        {
            // Extract text from fragment list
            std::wstring text = ExtractTextFromFragList(pTextFragList);
            if (text.empty())
            {
                LogMessage(L"No text to speak");
                return S_OK;
            }

            LogMessage((L"Speaking text: " + text).c_str());

            // Generate audio using native engine
            std::vector<BYTE> audioData;
            HRESULT hr = GenerateAudioViaNativeEngine(text, audioData);
            if (FAILED(hr))
            {
                LogMessage(L"Failed to generate audio");
                return hr;
            }

            LogMessage((L"Generated " + std::to_wstring(audioData.size()) + L" bytes of audio").c_str());

            // Send start event
            SPEVENT startEvent = { 0 };
            startEvent.eEventId = SPEI_START_INPUT_STREAM;
            startEvent.elParamType = SPET_LPARAM_IS_UNDEFINED;
            startEvent.ullAudioStreamOffset = 0;
            pOutputSite->AddEvents(&startEvent, 1);

            // Write audio data
            ULONG bytesWritten = 0;
            HRESULT hrWrite = pOutputSite->Write(audioData.data(), (ULONG)audioData.size(), &bytesWritten);
            if (FAILED(hrWrite))
            {
                LogMessage((L"Failed to write audio data: " + std::to_wstring(hrWrite)).c_str());
                return hrWrite;
            }

            LogMessage((L"Successfully wrote " + std::to_wstring(bytesWritten) + L" bytes").c_str());

            // Send end event
            SPEVENT endEvent = { 0 };
            endEvent.eEventId = SPEI_END_INPUT_STREAM;
            endEvent.elParamType = SPET_LPARAM_IS_UNDEFINED;
            endEvent.ullAudioStreamOffset = audioData.size();
            pOutputSite->AddEvents(&endEvent, 1);

            return S_OK;
        }
        catch (...)
        {
            LogMessage(L"Exception in Speak method");
            return E_FAIL;
        }
    }

STDMETHODIMP CNativeTTSWrapper::GetOutputFormat(
    const GUID* pTargetFormatId,
    const WAVEFORMATEX* pTargetWaveFormatEx,
    GUID* pOutputFormatId,
    WAVEFORMATEX** ppCoMemOutputWaveFormatEx)
    {
        LogMessage(L"*** NATIVE GET OUTPUT FORMAT CALLED ***");

        if (!pOutputFormatId || !ppCoMemOutputWaveFormatEx)
            return E_INVALIDARG;

        // Return standard PCM format
        *pOutputFormatId = SPDFID_WaveFormatEx;

        WAVEFORMATEX* pFormat = (WAVEFORMATEX*)CoTaskMemAlloc(sizeof(WAVEFORMATEX));
        if (!pFormat)
            return E_OUTOFMEMORY;

        pFormat->wFormatTag = WAVE_FORMAT_PCM;
        pFormat->nChannels = 1;
        // Use the actual sample rate from the engine (e.g., 16000 for SherpaOnnx models)
        // This prevents pitch issues from mismatched sample rates
        pFormat->nSamplesPerSec = m_actualSampleRate;
        pFormat->wBitsPerSample = 16;
        pFormat->nBlockAlign = pFormat->nChannels * pFormat->wBitsPerSample / 8;
        pFormat->nAvgBytesPerSec = pFormat->nSamplesPerSec * pFormat->nBlockAlign;
        pFormat->cbSize = 0;

        *ppCoMemOutputWaveFormatEx = pFormat;

        LogMessage((L"Returned PCM format: " + std::to_wstring(m_actualSampleRate) + L"Hz, 16-bit, mono").c_str());
        return S_OK;
    }

// ISpObjectWithToken implementation
STDMETHODIMP CNativeTTSWrapper::SetObjectToken(ISpObjectToken* pToken)
{
    LogMessage(L"*** NATIVE SET OBJECT TOKEN CALLED ***");

    if (!pToken)
        return E_INVALIDARG;

    m_pToken = pToken;

    // Get token ID (voice name) from the token
    LPWSTR tokenId = nullptr;
    HRESULT hr = pToken->GetId(&tokenId);
    if (SUCCEEDED(hr) && tokenId)
    {
        LogMessage((L"SetObjectToken - Token ID: " + std::wstring(tokenId)).c_str());

        // Convert token ID to voice name (extract just the name part)
        std::wstring tokenStr(tokenId);
        // Token ID format: HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\SPEECH\Voices\Tokens\{VoiceName}
        size_t lastBackslash = tokenStr.find_last_of(L"\\");
        if (lastBackslash != std::wstring::npos)
        {
            std::wstring voiceName = tokenStr.substr(lastBackslash + 1);
            LogMessage((L"SetObjectToken - Voice Name: " + voiceName).c_str());

            // For now, use a default engine ID based on the voice name
            // In production, this should look up the engine ID from a mapping
            if (voiceName.find(L"TestSherpa") != std::wstring::npos)
            {
                m_currentEngineId = L"sherpa-amy";  // Use sherpa-amy for testing
                LogMessage(L"SetObjectToken - Using sherpa-amy engine for test voice");
            }
            else if (voiceName.find(L"Amy") != std::wstring::npos ||
                     voiceName.find(L"amy") != std::wstring::npos)
            {
                m_currentEngineId = L"sherpa-amy";
                LogMessage(L"SetObjectToken - Using sherpa-amy engine");
            }
            else if (voiceName.find(L"Jenny") != std::wstring::npos ||
                     voiceName.find(L"jenny") != std::wstring::npos)
            {
                m_currentEngineId = L"azure-jenny";
                LogMessage(L"SetObjectToken - Using azure-jenny engine");
            }
            else
            {
                // First try to look up the voice name in the voice-to-engine mapping
                // This handles cases like "mms_hat" -> "sherpa-mms_hat"
                NativeTTS::TTSEngineManager& manager = NativeTTS::TTSEngineManagerSingleton::GetInstance();
                std::wstring mappedEngineId = manager.GetEngineIdForVoice(voiceName);
                if (!mappedEngineId.empty())
                {
                    m_currentEngineId = mappedEngineId;
                    LogMessage((L"SetObjectToken - Voice '" + voiceName + L"' maps to engine '" + m_currentEngineId + L"'").c_str());
                }
                else
                {
                    // No mapping found, use voice name directly as engine ID
                    m_currentEngineId = voiceName;
                    LogMessage((L"SetObjectToken - Using voice name as engine ID: " + m_currentEngineId).c_str());
                }
            }
        }

        CoTaskMemFree(tokenId);
    }
    else
    {
        LogMessage(L"SetObjectToken - Failed to get token ID");
    }

    LogMessage((L"SetObjectToken completed - Engine ID: " + m_currentEngineId).c_str());
    return S_OK;
}

STDMETHODIMP CNativeTTSWrapper::GetObjectToken(ISpObjectToken** ppToken)
{
    LogMessage(L"GetObjectToken called");

    if (!ppToken)
        return E_INVALIDARG;

    return m_pToken.CopyTo(ppToken);
}

// Private helper methods
void CNativeTTSWrapper::LogMessage(const wchar_t* message)
    {
        // Also output to debugger for immediate visibility
        OutputDebugStringW(message);
        OutputDebugStringW(L"\n");

        try
        {
            // Use DLL directory for log file
            std::wstring logPath = GetModuleDirectory() + L"\\native_tts_debug.log";
            std::wofstream logFile(logPath, std::ios::app);
            if (logFile.is_open())
            {
                SYSTEMTIME st;
                GetLocalTime(&st);
                logFile << st.wYear << L"-" << st.wMonth << L"-" << st.wDay << L" "
                       << st.wHour << L":" << st.wMinute << L":" << st.wSecond << L"."
                       << st.wMilliseconds << L": " << message << std::endl;
                logFile.close();
            }
        }
        catch (...) {}
    }

std::wstring CNativeTTSWrapper::ExtractTextFromFragList(const SPVTEXTFRAG* pTextFragList)
    {
        std::wstring result;
        const SPVTEXTFRAG* pFrag = pTextFragList;

        while (pFrag)
        {
            if (pFrag->pTextStart && pFrag->ulTextLen > 0)
            {
                result.append(pFrag->pTextStart, pFrag->ulTextLen);
            }
            pFrag = pFrag->pNext;
        }

        return result;
    }

// NEW: Native Engine Integration Methods
HRESULT CNativeTTSWrapper::GenerateAudioViaNativeEngine(const std::wstring& text, std::vector<BYTE>& audioData)
{
    try
    {
        LogMessage(L"Starting native engine audio generation...");

        // Get the engine manager instance
        NativeTTS::TTSEngineManager& manager = NativeTTS::TTSEngineManagerSingleton::GetInstance();

        // Get the engine for this voice
        NativeTTS::ITTSEngine* engine = manager.GetEngine(m_currentEngineId);
        if (!engine)
        {
            LogMessage(L"No engine found for current voice, attempting initialization...");
            HRESULT hr = InitializeEngineFromToken(m_pToken);
            if (FAILED(hr))
            {
                LogMessage(L"Failed to initialize engine from token");
                return hr;
            }
            engine = manager.GetEngine(m_currentEngineId);
            if (!engine)
            {
                LogMessage(L"Still no engine after initialization");
                return E_FAIL;
            }
        }

        // Check if engine is initialized
        if (!engine->IsInitialized())
        {
            LogMessage(L"Engine not initialized");
            return E_FAIL;
        }

        // Generate audio using the engine
        std::vector<float> samples;
        int sampleRate;
        HRESULT hr = engine->Generate(text, samples, sampleRate);
        if (FAILED(hr))
        {
            LogMessage(L"Engine audio generation failed");
            return hr;
        }

        LogMessage((L"Generated " + std::to_wstring(samples.size()) + L" samples at " + std::to_wstring(sampleRate) + L"Hz").c_str());

        // Convert float samples to byte array (16-bit PCM)
        hr = ConvertFloatSamplesToBytes(samples, sampleRate, audioData);
        if (FAILED(hr))
        {
            LogMessage(L"Failed to convert samples to bytes");
            return hr;
        }

        LogMessage((L"Converted to " + std::to_wstring(audioData.size()) + L" bytes of audio data").c_str());
        return S_OK;
    }
    catch (const std::exception& ex)
    {
        std::string error = "Exception in GenerateAudioViaNativeEngine: ";
        error += ex.what();
        LogMessage(std::wstring(error.begin(), error.end()).c_str());
        return E_FAIL;
    }
    catch (...)
    {
        LogMessage(L"Unknown exception in GenerateAudioViaNativeEngine");
        return E_FAIL;
    }
}

HRESULT CNativeTTSWrapper::InitializeEngineFromToken(ISpObjectToken* pToken)
{
    try
    {
        LogMessage(L"Initializing engine from token...");

        if (!pToken)
        {
            LogMessage(L"No token provided");
            return E_INVALIDARG;
        }

        // Use the engine ID that was already set in SetObjectToken
        if (!m_currentEngineId.empty())
        {
            LogMessage((L"Using engine ID from SetObjectToken: " + m_currentEngineId).c_str());

            // Get the engine manager
            NativeTTS::TTSEngineManager& manager = NativeTTS::TTSEngineManagerSingleton::GetInstance();

            // Check if engine already exists
            if (manager.GetEngine(m_currentEngineId))
            {
                LogMessage(L"Engine already loaded");

                // Query the engine for its actual sample rate
                NativeTTS::ITTSEngine* engine = manager.GetEngine(m_currentEngineId);
                if (engine)
                {
                    int channels, bitsPerSample;
                    HRESULT hr = engine->GetSupportedFormat(m_actualSampleRate, channels, bitsPerSample);
                    if (SUCCEEDED(hr))
                    {
                        LogMessage((L"Engine sample rate: " + std::to_wstring(m_actualSampleRate) + L"Hz").c_str());
                    }
                    else
                    {
                        LogMessage(L"Could not query engine sample rate, using default 16000Hz");
                        m_actualSampleRate = 16000;
                    }
                }

                m_engineInitialized = true;
                return S_OK;
            }

            // Try to load configuration from module directory
            std::wstring moduleDir = GetModuleDirectory();
            std::wstring configPath = moduleDir + L"\\engines_config.json";
            LogMessage((L"Loading config from: " + configPath).c_str());
            HRESULT hr = manager.LoadConfiguration(configPath);
            if (FAILED(hr))
            {
                LogMessage(L"Failed to load configuration, using fallback...");

                // Use fallback configuration based on engine ID
                if (m_currentEngineId == L"sherpa-amy" || m_currentEngineId == L"piper-en-amy-medium")
                {
                    // Use relative paths from module directory
                    std::wstring modelPath = L"C:/github/SherpaOnnxAzureSAPI-installer/models/amy/vits-piper-en_US-amy-low/en_US-amy-low.onnx";
                    std::wstring tokensPath = L"C:/github/SherpaOnnxAzureSAPI-installer/models/amy/vits-piper-en_US-amy-low/tokens.txt";
                    std::wstring dataDir = L"C:/github/SherpaOnnxAzureSAPI-installer/models/amy/vits-piper-en_US-amy-low/espeak-ng-data";

                    std::wstring amyConfig = LR"({
                        "engines": {
                            "sherpa-amy": {
                                "type": "sherpaonnx",
                                "config": {
                                    "modelPath": "MODEL_PATH_PLACEHOLDER",
                                    "tokensPath": "TOKENS_PATH_PLACEHOLDER",
                                    "dataDir": "DATA_DIR_PLACEHOLDER",
                                    "noiseScale": 0.667,
                                    "noiseScaleW": 0.8,
                                    "lengthScale": 1.0,
                                    "numThreads": 1
                                }
                            }
                        },
                        "voices": {
                            "amy": "sherpa-amy",
                            "sherpa-amy": "sherpa-amy"
                        }
                    })";

                    // Replace placeholders with actual paths
                    size_t pos = 0;
                    while ((pos = amyConfig.find(L"MODEL_PATH_PLACEHOLDER", pos)) != std::wstring::npos) {
                        amyConfig.replace(pos, 21, modelPath);
                    }
                    pos = 0;
                    while ((pos = amyConfig.find(L"TOKENS_PATH_PLACEHOLDER", pos)) != std::wstring::npos) {
                        amyConfig.replace(pos, 22, tokensPath);
                    }
                    pos = 0;
                    while ((pos = amyConfig.find(L"DATA_DIR_PLACEHOLDER", pos)) != std::wstring::npos) {
                        amyConfig.replace(pos, 19, dataDir);
                    }

                    LogMessage((L"Using model path: " + modelPath).c_str());
                    hr = manager.ParseConfiguration(amyConfig);
                    if (SUCCEEDED(hr))
                    {
                        LogMessage(L"Loaded fallback Amy configuration");

                        // Query the engine for its actual sample rate
                        NativeTTS::ITTSEngine* engine = manager.GetEngine(m_currentEngineId);
                        if (engine)
                        {
                            int channels, bitsPerSample;
                            HRESULT hrQuery = engine->GetSupportedFormat(m_actualSampleRate, channels, bitsPerSample);
                            if (SUCCEEDED(hrQuery))
                            {
                                LogMessage((L"Engine sample rate: " + std::to_wstring(m_actualSampleRate) + L"Hz").c_str());
                            }
                            else
                            {
                                LogMessage(L"Could not query engine sample rate, using default 16000Hz");
                                m_actualSampleRate = 16000;
                            }
                        }

                        m_engineInitialized = true;
                        return S_OK;
                    }
                }

                LogMessage(L"Failed to load fallback configuration");
                return E_FAIL;
            }

            LogMessage((L"Configuration loaded, using engine: " + m_currentEngineId).c_str());

                // Query the engine for its actual sample rate
                NativeTTS::ITTSEngine* engine = manager.GetEngine(m_currentEngineId);
            if (engine && engine->IsInitialized())
            {
                int channels, bitsPerSample;
                HRESULT hr = engine->GetSupportedFormat(m_actualSampleRate, channels, bitsPerSample);
                if (SUCCEEDED(hr))
                {
                    LogMessage((L"Engine sample rate: " + std::to_wstring(m_actualSampleRate) + L"Hz").c_str());
                }
                else
                {
                    LogMessage(L"Could not query engine sample rate, using default 16000Hz");
                    m_actualSampleRate = 16000;
                }
            }
            else
            {
                LogMessage(L"Engine not initialized yet, will query sample rate later");
                m_actualSampleRate = 16000; // Default for SherpaOnnx models
            }

            m_engineInitialized = true;
            return S_OK;
        }

        LogMessage(L"No engine ID set in SetObjectToken");
        return E_FAIL;
    }
    catch (...)
    {
        LogMessage(L"Exception in InitializeEngineFromToken");
        return E_FAIL;
    }
}

HRESULT CNativeTTSWrapper::ConvertFloatSamplesToBytes(const std::vector<float>& samples, int sampleRate, std::vector<BYTE>& audioData)
{
    try
    {
        // Calculate sizes
        size_t numSamples = samples.size();
        size_t wavHeaderSize = 44;
        size_t audioDataSize = numSamples * sizeof(int16_t);
        size_t totalSize = wavHeaderSize + audioDataSize;

        // Prepare output buffer
        audioData.resize(totalSize);
        BYTE* pData = audioData.data();

        // Write WAV header
        memcpy(pData, "RIFF", 4);
        *reinterpret_cast<uint32_t*>(pData + 4) = static_cast<uint32_t>(totalSize - 8);
        memcpy(pData + 8, "WAVE", 4);
        memcpy(pData + 12, "fmt ", 4);
        *reinterpret_cast<uint32_t*>(pData + 16) = 16; // PCM format chunk size
        *reinterpret_cast<uint16_t*>(pData + 20) = 1;  // PCM format
        *reinterpret_cast<uint16_t*>(pData + 22) = 1;  // Mono
        *reinterpret_cast<uint32_t*>(pData + 24) = static_cast<uint32_t>(sampleRate);
        *reinterpret_cast<uint32_t*>(pData + 28) = static_cast<uint32_t>(sampleRate * 2); // Byte rate
        *reinterpret_cast<uint16_t*>(pData + 32) = 2;  // Block align
        *reinterpret_cast<uint16_t*>(pData + 34) = 16; // Bits per sample
        memcpy(pData + 36, "data", 4);
        *reinterpret_cast<uint32_t*>(pData + 40) = static_cast<uint32_t>(audioDataSize);

        // Convert float samples to 16-bit PCM
        int16_t* pcmData = reinterpret_cast<int16_t*>(pData + wavHeaderSize);
        for (size_t i = 0; i < numSamples; ++i)
        {
            // Clamp to [-1, 1] and convert to 16-bit
            float sample = samples[i];
            if (sample < -1.0f) sample = -1.0f;
            if (sample > 1.0f) sample = 1.0f;
            pcmData[i] = static_cast<int16_t>(sample * 32767.0f);
        }

        LogMessage(L"Successfully converted float samples to WAV format");
        return S_OK;
    }
    catch (...)
    {
        LogMessage(L"Exception in ConvertFloatSamplesToBytes");
        return E_FAIL;
    }
}

std::string CNativeTTSWrapper::WStringToUTF8(const std::wstring& wstr)
    {
        if (wstr.empty()) return std::string();
        int size_needed = WideCharToMultiByte(CP_UTF8, 0, &wstr[0], (int)wstr.size(), nullptr, 0, nullptr, nullptr);
        std::string strTo(size_needed, 0);
        WideCharToMultiByte(CP_UTF8, 0, &wstr[0], (int)wstr.size(), &strTo[0], size_needed, nullptr, nullptr);
        return strTo;
    }

std::wstring CNativeTTSWrapper::UTF8ToWString(const std::string& str)
    {
        if (str.empty()) return std::wstring();
        int size_needed = MultiByteToWideChar(CP_UTF8, 0, &str[0], (int)str.size(), nullptr, 0);
        std::wstring wstrTo(size_needed, 0);
        MultiByteToWideChar(CP_UTF8, 0, &str[0], (int)str.size(), &wstrTo[0], size_needed);
        return wstrTo;
    }

std::wstring CNativeTTSWrapper::GetModuleDirectory()
{
    wchar_t modulePath[MAX_PATH];
    GetModuleFileNameW((HMODULE)&__ImageBase, modulePath, MAX_PATH);

    // Get directory part
    std::wstring path(modulePath);
    size_t lastSlash = path.find_last_of(L"\\/");
    if (lastSlash != std::wstring::npos)
    {
        return path.substr(0, lastSlash);
    }
    return path;
}

// COM class factory and registration
OBJECT_ENTRY_AUTO(CLSID_CNativeTTSWrapper, CNativeTTSWrapper)
