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
#include "TTSEngineManager.h"

// Implementation of CNativeTTSWrapper

CNativeTTSWrapper::CNativeTTSWrapper() : m_engineInitialized(false)
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

            // Try native engine first, then direct SherpaOnnx, then ProcessBridge
            std::vector<BYTE> audioData;

            LogMessage(L"=== STARTING FALLBACK CHAIN ===");
            LogMessage(L"Step 1: Attempting native engine...");
            HRESULT hr = GenerateAudioViaNativeEngine(text, audioData);
            LogMessage((L"Native engine result: " + std::to_wstring(hr)).c_str());

            if (FAILED(hr))
            {
                LogMessage(L"Native engine failed, trying direct SherpaOnnx...");
                LogMessage(L"Step 2: Attempting direct SherpaOnnx...");
                hr = GenerateAudioViaDirectSherpaOnnx(text, audioData);
                LogMessage((L"Direct SherpaOnnx result: " + std::to_wstring(hr)).c_str());

                if (FAILED(hr))
                {
                    LogMessage(L"Direct SherpaOnnx failed, trying AACSpeakHelper pipe service...");
                    LogMessage(L"Step 3: Attempting AACSpeakHelper pipe service...");
                    hr = GenerateAudioViaPipeService(text, audioData);
                    LogMessage((L"AACSpeakHelper pipe service result: " + std::to_wstring(hr)).c_str());

                    if (FAILED(hr))
                    {
                        LogMessage(L"AACSpeakHelper pipe service failed, trying ProcessBridge fallback...");
                        LogMessage(L"Step 4: Attempting ProcessBridge fallback...");
                        if (!GenerateAudioViaProcessBridge(text, audioData))
                        {
                            LogMessage(L"All methods failed: native engine, direct SherpaOnnx, AACSpeakHelper pipe, and ProcessBridge");
                            return E_FAIL;
                        }
                        LogMessage(L"ProcessBridge fallback succeeded");
                    }
                    else
                    {
                        LogMessage(L"AACSpeakHelper pipe service succeeded");
                    }
                }
                else
                {
                    LogMessage(L"Direct SherpaOnnx succeeded");
                }
            }
            else
            {
                LogMessage(L"Native engine generation succeeded");
            }

            LogMessage(L"=== FALLBACK CHAIN COMPLETE ===");

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
        pFormat->nSamplesPerSec = 22050;
        pFormat->wBitsPerSample = 16;
        pFormat->nBlockAlign = pFormat->nChannels * pFormat->wBitsPerSample / 8;
        pFormat->nAvgBytesPerSec = pFormat->nSamplesPerSec * pFormat->nBlockAlign;
        pFormat->cbSize = 0;

        *ppCoMemOutputWaveFormatEx = pFormat;

        LogMessage(L"Returned PCM format: 22050Hz, 16-bit, mono");
        return S_OK;
    }

// ISpObjectWithToken implementation
STDMETHODIMP CNativeTTSWrapper::SetObjectToken(ISpObjectToken* pToken)
{
    LogMessage(L"*** NATIVE SET OBJECT TOKEN CALLED ***");

    m_pToken = pToken;

    LogMessage(L"SetObjectToken completed successfully");
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
        try
        {
            std::wofstream logFile(L"C:\\OpenSpeech\\native_tts_debug.log", std::ios::app);
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

        // Get voice name from token
        CSpDynamicString voiceName;
        HRESULT hr = pToken->GetStringValue(L"VoiceName", &voiceName);
        if (FAILED(hr) || !voiceName)
        {
            LogMessage(L"Failed to get voice name from token");
            return hr;
        }

        std::wstring voiceNameStr = static_cast<LPCWSTR>(voiceName);
        LogMessage((L"Voice name: " + voiceNameStr).c_str());

        // Get the engine manager and load configuration
        NativeTTS::TTSEngineManager& manager = NativeTTS::TTSEngineManagerSingleton::GetInstance();

        // Try to load configuration from the standard location
        std::wstring configPath = L"engines_config.json";
        hr = manager.LoadConfiguration(configPath);
        if (FAILED(hr))
        {
            LogMessage(L"Failed to load engine configuration, trying fallback...");

            // Try fallback configuration for Amy voice
            if (voiceNameStr == L"amy" || voiceNameStr == L"Amy" || voiceNameStr == L"piper-en-amy-medium")
            {
                std::wstring amyConfig = LR"({
                    "engines": {
                        "piper-en-amy-medium": {
                            "type": "sherpaonnx",
                            "config": {
                                "modelPath": "C:/Program Files/OpenAssistive/OpenSpeech/models/amy/model.onnx",
                                "tokensPath": "C:/Program Files/OpenAssistive/OpenSpeech/models/amy/tokens.txt",
                                "dataDir": "C:/Program Files/OpenAssistive/OpenSpeech/models/amy/espeak-ng-data",
                                "noiseScale": 0.667,
                                "noiseScaleW": 0.8,
                                "lengthScale": 1.0,
                                "numThreads": 1
                            }
                        }
                    },
                    "voices": {
                        "amy": "piper-en-amy-medium",
                        "piper-en-amy-medium": "piper-en-amy-medium"
                    }
                })";

                hr = manager.ParseConfiguration(amyConfig);
                if (FAILED(hr))
                {
                    LogMessage(L"Failed to parse fallback Amy configuration");
                    return hr;
                }

                m_currentEngineId = L"piper-en-amy-medium";
                LogMessage(L"Using fallback Amy configuration");
            }
            else
            {
                LogMessage(L"No fallback configuration available for this voice");
                return E_FAIL;
            }
        }
        else
        {
            // Get engine ID for this voice
            m_currentEngineId = manager.GetEngineIdForVoice(voiceNameStr);
            if (m_currentEngineId.empty())
            {
                LogMessage(L"No engine mapping found for voice");
                return E_FAIL;
            }
        }

        LogMessage((L"Using engine: " + m_currentEngineId).c_str());
        m_engineInitialized = true;

        return S_OK;
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

HRESULT CNativeTTSWrapper::GenerateAudioViaDirectSherpaOnnx(const std::wstring& text, std::vector<BYTE>& audioData)
{
    try
    {
        LogMessage(L"Starting direct SherpaOnnx audio generation...");

        // Create SherpaOnnx configuration
        SherpaOnnxOfflineTtsConfig config;
        memset(&config, 0, sizeof(config));

        // Set up the model configuration for Amy
        config.model.vits.model = "C:/Program Files/OpenAssistive/OpenSpeech/models/amy/model.onnx";
        config.model.vits.tokens = "C:/Program Files/OpenAssistive/OpenSpeech/models/amy/tokens.txt";
        config.model.vits.data_dir = "C:/Program Files/OpenAssistive/OpenSpeech/models/amy/espeak-ng-data";
        config.model.vits.noise_scale = 0.667f;
        config.model.vits.noise_scale_w = 0.8f;
        config.model.vits.length_scale = 1.0f;
        config.model.num_threads = 1;
        config.model.provider = "cpu";
        config.model.debug = 0;

        LogMessage(L"Creating SherpaOnnx TTS instance...");

        // Create the TTS instance
        const SherpaOnnxOfflineTts* tts = SherpaOnnxCreateOfflineTts(&config);
        if (!tts)
        {
            LogMessage(L"Failed to create SherpaOnnx TTS instance");
            return E_FAIL;
        }

        LogMessage(L"SherpaOnnx TTS instance created successfully");

        // Get sample rate
        int sampleRate = SherpaOnnxOfflineTtsSampleRate(tts);
        LogMessage((L"Sample rate: " + std::to_wstring(sampleRate) + L"Hz").c_str());

        // Convert text to UTF-8
        std::string utf8Text = WStringToUTF8(text);
        LogMessage((L"Generating audio for: " + text).c_str());

        // Generate audio
        const SherpaOnnxGeneratedAudio* audio = SherpaOnnxOfflineTtsGenerate(
            tts, utf8Text.c_str(), 0, 1.0f);

        if (!audio || !audio->samples || audio->n <= 0)
        {
            LogMessage(L"SherpaOnnx generation failed or returned empty audio");
            SherpaOnnxDestroyOfflineTts(tts);
            return E_FAIL;
        }

        LogMessage((L"Generated " + std::to_wstring(audio->n) + L" audio samples").c_str());

        // Convert float samples to vector
        std::vector<float> samples(audio->samples, audio->samples + audio->n);

        // Free the generated audio
        SherpaOnnxDestroyOfflineTtsGeneratedAudio(audio);

        // Destroy the TTS instance
        SherpaOnnxDestroyOfflineTts(tts);

        // Convert float samples to byte array (16-bit PCM WAV)
        HRESULT hr = ConvertFloatSamplesToBytes(samples, sampleRate, audioData);
        if (FAILED(hr))
        {
            LogMessage(L"Failed to convert samples to bytes");
            return hr;
        }

        LogMessage((L"Direct SherpaOnnx generated " + std::to_wstring(audioData.size()) + L" bytes of audio data").c_str());
        return S_OK;
    }
    catch (const std::exception& ex)
    {
        std::string error = "Exception in GenerateAudioViaDirectSherpaOnnx: ";
        error += ex.what();
        LogMessage(std::wstring(error.begin(), error.end()).c_str());
        return E_FAIL;
    }
    catch (...)
    {
        LogMessage(L"Unknown exception in GenerateAudioViaDirectSherpaOnnx");
        return E_FAIL;
    }
}

// NEW: AACSpeakHelper Pipe Service Implementation
HRESULT CNativeTTSWrapper::GenerateAudioViaPipeService(const std::wstring& text, std::vector<BYTE>& audioData)
{
    try
    {
        LogMessage(L"Starting AACSpeakHelper pipe service audio generation...");

        // Connect to AACSpeakHelper pipe
        HANDLE hPipe = INVALID_HANDLE_VALUE;
        if (!ConnectToAACSpeakHelper(hPipe))
        {
            LogMessage(L"Failed to connect to AACSpeakHelper pipe service");
            return E_FAIL;
        }

        LogMessage(L"Connected to AACSpeakHelper pipe service");

        // Create JSON message for AACSpeakHelper
        std::string jsonMessage = CreateAACSpeakHelperMessage(text);
        LogMessage((L"Sending message: " + UTF8ToWString(jsonMessage)).c_str());

        // Send message to pipe
        if (!SendPipeMessage(hPipe, jsonMessage))
        {
            LogMessage(L"Failed to send message to AACSpeakHelper");
            CloseHandle(hPipe);
            return E_FAIL;
        }

        LogMessage(L"Message sent successfully");

        // Receive audio response
        if (!ReceivePipeResponse(hPipe, audioData))
        {
            LogMessage(L"Failed to receive audio response from AACSpeakHelper");
            CloseHandle(hPipe);
            return E_FAIL;
        }

        LogMessage((L"Received " + std::to_wstring(audioData.size()) + L" bytes of audio data").c_str());

        CloseHandle(hPipe);
        return S_OK;
    }
    catch (const std::exception& ex)
    {
        std::string error = "Exception in GenerateAudioViaPipeService: ";
        error += ex.what();
        LogMessage(std::wstring(error.begin(), error.end()).c_str());
        return E_FAIL;
    }
    catch (...)
    {
        LogMessage(L"Unknown exception in GenerateAudioViaPipeService");
        return E_FAIL;
    }
}

bool CNativeTTSWrapper::ConnectToAACSpeakHelper(HANDLE& hPipe)
{
    const wchar_t* pipeName = L"\\\\.\\pipe\\AACSpeakHelper";
    const int maxRetries = 5;
    const int retryDelayMs = 1000;

    for (int retry = 0; retry < maxRetries; retry++)
    {
        LogMessage((L"Attempting to connect to pipe (attempt " + std::to_wstring(retry + 1) + L")").c_str());

        hPipe = CreateFileW(
            pipeName,
            GENERIC_READ | GENERIC_WRITE,
            0,
            nullptr,
            OPEN_EXISTING,
            0,
            nullptr
        );

        if (hPipe != INVALID_HANDLE_VALUE)
        {
            LogMessage(L"Successfully connected to AACSpeakHelper pipe");
            return true;
        }

        DWORD error = GetLastError();
        if (error == ERROR_PIPE_BUSY)
        {
            LogMessage(L"Pipe is busy, waiting...");
            if (!WaitNamedPipeW(pipeName, 30000)) // 30 second timeout
            {
                LogMessage(L"Timeout waiting for pipe to become available");
                continue;
            }
        }
        else
        {
            LogMessage((L"Failed to connect to pipe, error: " + std::to_wstring(error)).c_str());
        }

        if (retry < maxRetries - 1)
        {
            Sleep(retryDelayMs);
        }
    }

    LogMessage(L"Failed to connect to AACSpeakHelper pipe after all retries");
    return false;
}

bool CNativeTTSWrapper::SendPipeMessage(HANDLE hPipe, const std::string& jsonMessage)
{
    try
    {
        // Send message length first
        uint32_t messageLength = static_cast<uint32_t>(jsonMessage.length());
        DWORD bytesWritten = 0;

        if (!WriteFile(hPipe, &messageLength, sizeof(messageLength), &bytesWritten, nullptr) ||
            bytesWritten != sizeof(messageLength))
        {
            LogMessage(L"Failed to write message length to pipe");
            return false;
        }

        // Send message content
        if (!WriteFile(hPipe, jsonMessage.c_str(), messageLength, &bytesWritten, nullptr) ||
            bytesWritten != messageLength)
        {
            LogMessage(L"Failed to write message content to pipe");
            return false;
        }

        LogMessage(L"Message sent to pipe successfully");
        return true;
    }
    catch (...)
    {
        LogMessage(L"Exception in SendPipeMessage");
        return false;
    }
}

bool CNativeTTSWrapper::ReceivePipeResponse(HANDLE hPipe, std::vector<BYTE>& audioData)
{
    try
    {
        // Read response length first
        uint32_t responseLength = 0;
        DWORD bytesRead = 0;

        if (!ReadFile(hPipe, &responseLength, sizeof(responseLength), &bytesRead, nullptr) ||
            bytesRead != sizeof(responseLength))
        {
            LogMessage(L"Failed to read response length from pipe");
            return false;
        }

        LogMessage((L"Expecting " + std::to_wstring(responseLength) + L" bytes of response").c_str());

        if (responseLength == 0 || responseLength > 100 * 1024 * 1024) // 100MB max
        {
            LogMessage(L"Invalid response length");
            return false;
        }

        // Read response content
        audioData.resize(responseLength);
        if (!ReadFile(hPipe, audioData.data(), responseLength, &bytesRead, nullptr) ||
            bytesRead != responseLength)
        {
            LogMessage(L"Failed to read response content from pipe");
            return false;
        }

        LogMessage(L"Response received from pipe successfully");
        return true;
    }
    catch (...)
    {
        LogMessage(L"Exception in ReceivePipeResponse");
        return false;
    }
}

std::string CNativeTTSWrapper::CreateAACSpeakHelperMessage(const std::wstring& text)
{
    try
    {
        // Load voice configuration
        std::wstring voiceConfig = LoadVoiceConfiguration();

        // Convert text to UTF-8
        std::string utf8Text = WStringToUTF8(text);

        // Create JSON message in AACSpeakHelper format
        std::string jsonMessage = "{\n";
        jsonMessage += "  \"text\": \"";

        // Escape text for JSON
        for (char c : utf8Text)
        {
            if (c == '"') jsonMessage += "\\\"";
            else if (c == '\\') jsonMessage += "\\\\";
            else if (c == '\n') jsonMessage += "\\n";
            else if (c == '\r') jsonMessage += "\\r";
            else if (c == '\t') jsonMessage += "\\t";
            else jsonMessage += c;
        }

        jsonMessage += "\",\n";

        // Add voice configuration from loaded config or use defaults
        if (!voiceConfig.empty())
        {
            std::string utf8Config = WStringToUTF8(voiceConfig);
            jsonMessage += "  " + utf8Config + "\n";
        }
        else
        {
            // Default SherpaOnnx configuration
            jsonMessage += "  \"args\": {\n";
            jsonMessage += "    \"engine\": \"sherpaonnx\",\n";
            jsonMessage += "    \"voice\": \"en_GB-jenny_dioco-medium\",\n";
            jsonMessage += "    \"rate\": 0,\n";
            jsonMessage += "    \"volume\": 100\n";
            jsonMessage += "  }\n";
        }

        jsonMessage += "}";

        return jsonMessage;
    }
    catch (...)
    {
        LogMessage(L"Exception in CreateAACSpeakHelperMessage");
        // Return default message
        std::string utf8Text = WStringToUTF8(text);
        return "{\"text\":\"" + utf8Text + "\",\"args\":{\"engine\":\"sherpaonnx\",\"voice\":\"en_GB-jenny_dioco-medium\",\"rate\":0,\"volume\":100}}";
    }
}

std::wstring CNativeTTSWrapper::LoadVoiceConfiguration()
{
    try
    {
        if (!m_pToken)
        {
            LogMessage(L"No token available for voice configuration");
            return L"";
        }

        // Get voice name from token
        CSpDynamicString voiceName;
        HRESULT hr = m_pToken->GetStringValue(L"VoiceName", &voiceName);
        if (FAILED(hr) || !voiceName)
        {
            LogMessage(L"Failed to get voice name from token");
            return L"";
        }

        std::wstring voiceNameStr = static_cast<LPCWSTR>(voiceName);
        LogMessage((L"Loading configuration for voice: " + voiceNameStr).c_str());

        // Try to load configuration file
        std::wstring configPath = L"voice_configs\\" + voiceNameStr + L".json";

        std::ifstream configFile(configPath);
        if (!configFile.is_open())
        {
            LogMessage((L"Voice configuration file not found: " + configPath).c_str());
            return L"";
        }

        std::string configContent((std::istreambuf_iterator<char>(configFile)),
                                 std::istreambuf_iterator<char>());
        configFile.close();

        // Parse JSON to extract ttsConfig.args section
        size_t argsStart = configContent.find("\"args\":");
        if (argsStart == std::string::npos)
        {
            LogMessage(L"No args section found in voice configuration");
            return L"";
        }

        // Find the args object
        size_t braceStart = configContent.find("{", argsStart);
        if (braceStart == std::string::npos)
        {
            LogMessage(L"Invalid args section in voice configuration");
            return L"";
        }

        // Find matching closing brace
        int braceCount = 1;
        size_t pos = braceStart + 1;
        while (pos < configContent.length() && braceCount > 0)
        {
            if (configContent[pos] == '{') braceCount++;
            else if (configContent[pos] == '}') braceCount--;
            pos++;
        }

        if (braceCount != 0)
        {
            LogMessage(L"Malformed JSON in voice configuration");
            return L"";
        }

        std::string argsSection = "\"args\":" + configContent.substr(braceStart, pos - braceStart);
        LogMessage((L"Loaded voice configuration: " + UTF8ToWString(argsSection)).c_str());

        return UTF8ToWString(argsSection);
    }
    catch (...)
    {
        LogMessage(L"Exception in LoadVoiceConfiguration");
        return L"";
    }
}

bool CNativeTTSWrapper::GenerateAudioViaProcessBridge(const std::wstring& text, std::vector<BYTE>& audioData)
    {
        try
        {
            LogMessage(L"Starting ProcessBridge audio generation...");

            // Convert text to UTF-8 for JSON
            std::string utf8Text = WStringToUTF8(text);

            // Create temporary files
            wchar_t tempPath[MAX_PATH];
            GetTempPathW(MAX_PATH, tempPath);
            wcscat_s(tempPath, L"OpenSpeechTTS\\");
            CreateDirectoryW(tempPath, nullptr);

            // Generate unique ID
            GUID guid;
            CoCreateGuid(&guid);
            wchar_t guidStr[40];
            StringFromGUID2(guid, guidStr, 40);

            std::wstring requestPath = tempPath + std::wstring(L"native_request_") + guidStr + L".json";
            std::wstring responsePath = tempPath + std::wstring(L"native_request_") + guidStr + L".response.json";
            std::wstring audioPath = tempPath + std::wstring(L"native_audio_") + guidStr;

            // Create JSON request with proper escaping
            std::string escapedText = utf8Text;
            std::string escapedPath = WStringToUTF8(audioPath);

            // Escape backslashes and quotes for JSON
            auto escapeForJson = [](std::string& str) {
                std::string::size_type pos = 0;
                // Escape backslashes first
                while ((pos = str.find("\\", pos)) != std::string::npos) {
                    str.replace(pos, 1, "\\\\");
                    pos += 2;
                }
                // Escape quotes
                pos = 0;
                while ((pos = str.find("\"", pos)) != std::string::npos) {
                    str.replace(pos, 1, "\\\"");
                    pos += 2;
                }
            };

            escapeForJson(escapedText);
            escapeForJson(escapedPath);

            std::string jsonRequest = "{\n";
            jsonRequest += "  \"Text\": \"" + escapedText + "\",\n";
            jsonRequest += "  \"Speed\": 1.0,\n";
            jsonRequest += "  \"SpeakerId\": 0,\n";
            jsonRequest += "  \"OutputPath\": \"" + escapedPath + "\"\n";
            jsonRequest += "}";

            // Write request file
            std::ofstream requestFile(requestPath);
            if (!requestFile.is_open())
            {
                LogMessage(L"Failed to create request file");
                return false;
            }
            requestFile << jsonRequest;
            requestFile.close();

            LogMessage((L"Created request file: " + requestPath).c_str());

            // Launch SherpaWorker
            std::wstring sherpaWorkerPath = L"C:\\Program Files\\OpenAssistive\\OpenSpeech\\SherpaWorker.exe";
            std::wstring commandLine = L"\"" + sherpaWorkerPath + L"\" \"" + requestPath + L"\"";

            STARTUPINFOW si = { sizeof(si) };
            PROCESS_INFORMATION pi = { 0 };

            if (!CreateProcessW(nullptr, const_cast<wchar_t*>(commandLine.c_str()),
                               nullptr, nullptr, FALSE, CREATE_NO_WINDOW,
                               nullptr, nullptr, &si, &pi))
            {
                LogMessage(L"Failed to launch SherpaWorker");
                return false;
            }

            // Wait for completion (30 second timeout)
            DWORD waitResult = WaitForSingleObject(pi.hProcess, 30000);
            if (waitResult != WAIT_OBJECT_0)
            {
                LogMessage(L"SherpaWorker timed out");
                TerminateProcess(pi.hProcess, 1);
                CloseHandle(pi.hProcess);
                CloseHandle(pi.hThread);
                return false;
            }

            DWORD exitCode;
            GetExitCodeProcess(pi.hProcess, &exitCode);
            CloseHandle(pi.hProcess);
            CloseHandle(pi.hThread);

            if (exitCode != 0)
            {
                LogMessage((L"SherpaWorker failed with exit code: " + std::to_wstring(exitCode)).c_str());
                return false;
            }

            LogMessage(L"SherpaWorker completed successfully");

            // Read response
            std::ifstream responseFile(responsePath);
            if (!responseFile.is_open())
            {
                LogMessage(L"Response file not found");
                return false;
            }

            std::string responseJson((std::istreambuf_iterator<char>(responseFile)),
                                   std::istreambuf_iterator<char>());
            responseFile.close();

            // Parse response (simple parsing for AudioPath)
            size_t audioPathStart = responseJson.find("\"AudioPath\": \"");
            if (audioPathStart == std::string::npos)
            {
                LogMessage(L"AudioPath not found in response");
                return false;
            }

            audioPathStart += 14; // Length of "\"AudioPath\": \""
            size_t audioPathEnd = responseJson.find("\"", audioPathStart);
            if (audioPathEnd == std::string::npos)
            {
                LogMessage(L"Invalid AudioPath in response");
                return false;
            }

            std::string audioFilePath = responseJson.substr(audioPathStart, audioPathEnd - audioPathStart);
            std::wstring audioFilePathW = UTF8ToWString(audioFilePath);

            LogMessage((L"Loading audio file: " + audioFilePathW).c_str());

            // Read audio file and extract PCM data (skip WAV header)
            std::ifstream audioFile(audioFilePathW, std::ios::binary);
            if (!audioFile.is_open())
            {
                LogMessage(L"Failed to open audio file");
                return false;
            }

            // Read and validate WAV header
            char riff[4];
            audioFile.read(riff, 4);
            if (strncmp(riff, "RIFF", 4) != 0)
            {
                LogMessage(L"Invalid WAV file - missing RIFF header");
                audioFile.close();
                return false;
            }

            // Skip file size
            audioFile.seekg(4, std::ios::cur);

            // Read WAVE header
            char wave[4];
            audioFile.read(wave, 4);
            if (strncmp(wave, "WAVE", 4) != 0)
            {
                LogMessage(L"Invalid WAV file - missing WAVE header");
                audioFile.close();
                return false;
            }

            // Find data chunk
            bool foundData = false;
            uint32_t dataSize = 0;

            while (!audioFile.eof() && !foundData)
            {
                char chunkId[4];
                uint32_t chunkSize;

                audioFile.read(chunkId, 4);
                audioFile.read(reinterpret_cast<char*>(&chunkSize), 4);

                if (strncmp(chunkId, "data", 4) == 0)
                {
                    foundData = true;
                    dataSize = chunkSize;
                }
                else
                {
                    // Skip this chunk
                    audioFile.seekg(chunkSize, std::ios::cur);
                }
            }

            if (!foundData)
            {
                LogMessage(L"No data chunk found in WAV file");
                audioFile.close();
                return false;
            }

            // Read only the PCM audio data (skip WAV header)
            audioData.resize(dataSize);
            audioFile.read(reinterpret_cast<char*>(audioData.data()), dataSize);
            audioFile.close();

            LogMessage((L"Loaded " + std::to_wstring(audioData.size()) + L" bytes of audio").c_str());

            // Cleanup
            DeleteFileW(requestPath.c_str());
            DeleteFileW(responsePath.c_str());
            DeleteFileW(audioFilePathW.c_str());

            return true;
        }
        catch (...)
        {
            LogMessage(L"Exception in GenerateAudioViaProcessBridge");
            return false;
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

// COM class factory and registration
OBJECT_ENTRY_AUTO(CLSID_CNativeTTSWrapper, CNativeTTSWrapper)
