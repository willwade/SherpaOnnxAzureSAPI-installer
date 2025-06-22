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

            // Use ONLY AACSpeakHelper pipe service - no fallbacks
            std::vector<BYTE> audioData;

            LogMessage(L"=== USING AACSPEAKHELPER PIPE SERVICE ONLY - BUILD 2025-06-22 18:40 ===");
            HRESULT hr = GenerateAudioViaPipeService(text, audioData);
            LogMessage((L"AACSpeakHelper pipe service result: " + std::to_wstring(hr)).c_str());

            if (FAILED(hr))
            {
                LogMessage(L"AACSpeakHelper pipe service failed - no fallbacks, returning error");
                return hr;
            }

            LogMessage(L"AACSpeakHelper pipe service succeeded");

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

        // Now receive audio bytes back from AACSpeakHelper
        if (!ReceivePipeResponse(hPipe, audioData))
        {
            LogMessage(L"Failed to receive audio bytes from AACSpeakHelper");
            CloseHandle(hPipe);
            return E_FAIL;
        }

        CloseHandle(hPipe);

        LogMessage((L"Successfully received " + std::to_wstring(audioData.size()) + L" bytes of audio from AACSpeakHelper").c_str());
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
        // Send message content directly (AACSpeakHelper expects raw message, no length prefix)
        DWORD bytesWritten = 0;
        DWORD messageLength = static_cast<DWORD>(jsonMessage.length());

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
        LogMessage(L"Waiting to receive audio bytes from AACSpeakHelper...");

        // First, read the length prefix (4 bytes, little-endian uint32)
        BYTE lengthBuffer[4];
        DWORD bytesRead = 0;

        if (!ReadFile(hPipe, lengthBuffer, sizeof(lengthBuffer), &bytesRead, nullptr) ||
            bytesRead != sizeof(lengthBuffer))
        {
            LogMessage(L"Failed to read length prefix from pipe");
            return false;
        }

        // Extract audio data length from little-endian uint32
        DWORD audioLength =
            (static_cast<DWORD>(lengthBuffer[0])) |
            (static_cast<DWORD>(lengthBuffer[1]) << 8) |
            (static_cast<DWORD>(lengthBuffer[2]) << 16) |
            (static_cast<DWORD>(lengthBuffer[3]) << 24);

        LogMessage((L"Expecting " + std::to_wstring(audioLength) + L" bytes of audio data").c_str());

        if (audioLength == 0 || audioLength > 10 * 1024 * 1024) // Sanity check: max 10MB
        {
            LogMessage((L"Invalid audio length: " + std::to_wstring(audioLength)).c_str());
            return false;
        }

        // Resize buffer to hold the audio data
        audioData.resize(audioLength);

        // Read the audio data in chunks
        DWORD totalBytesRead = 0;
        const DWORD chunkSize = 64 * 1024; // 64KB chunks

        while (totalBytesRead < audioLength)
        {
            DWORD remainingBytes = audioLength - totalBytesRead;
            DWORD bytesToRead = min(chunkSize, remainingBytes);
            DWORD chunkBytesRead = 0;

            if (!ReadFile(hPipe, audioData.data() + totalBytesRead, bytesToRead, &chunkBytesRead, nullptr))
            {
                LogMessage((L"Failed to read audio chunk at offset " + std::to_wstring(totalBytesRead)).c_str());
                return false;
            }

            if (chunkBytesRead == 0)
            {
                LogMessage(L"Unexpected end of pipe data");
                return false;
            }

            totalBytesRead += chunkBytesRead;
            LogMessage((L"Read " + std::to_wstring(totalBytesRead) + L"/" + std::to_wstring(audioLength) + L" bytes").c_str());
        }

        LogMessage((L"Successfully received " + std::to_wstring(audioData.size()) + L" bytes of audio data").c_str());
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
        // AACSpeakHelper expects: {"args": {...}, "config": {...}, "clipboard_text": "..."}
        std::string jsonMessage = "{\n";

        // Add clipboard_text (the text to speak)
        jsonMessage += "  \"clipboard_text\": \"";
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

        // Add args section
        if (!voiceConfig.empty())
        {
            std::string utf8Config = WStringToUTF8(voiceConfig);
            jsonMessage += "  " + utf8Config + ",\n";
        }
        else
        {
            // Default SherpaOnnx configuration
            jsonMessage += "  \"args\": {\n";
            jsonMessage += "    \"engine\": \"sherpaonnx\",\n";
            jsonMessage += "    \"voice\": \"en_GB-jenny_dioco-medium\",\n";
            jsonMessage += "    \"rate\": 0,\n";
            jsonMessage += "    \"volume\": 100,\n";
            jsonMessage += "    \"listvoices\": false,\n";
            jsonMessage += "    \"return_audio_bytes\": true\n";
            jsonMessage += "  },\n";
        }

        // Add config section - load from voice configuration
        std::string configSection = CreateConfigSection();
        jsonMessage += "  \"config\": " + configSection + "\n";
        jsonMessage += "}";

        return jsonMessage;
    }
    catch (...)
    {
        LogMessage(L"Exception in CreateAACSpeakHelperMessage");
        // Return default message in AACSpeakHelper format
        std::string utf8Text = WStringToUTF8(text);
        return "{\n"
               "  \"clipboard_text\": \"" + utf8Text + "\",\n"
               "  \"args\": {\n"
               "    \"engine\": \"sherpaonnx\",\n"
               "    \"voice\": \"en_GB-jenny_dioco-medium\",\n"
               "    \"rate\": 0,\n"
               "    \"volume\": 100,\n"
               "    \"listvoices\": false,\n"
               "    \"return_audio_bytes\": true\n"
               "  },\n"
               "  \"config\": " + CreateDefaultConfig() + "\n"
               "}";
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

        // Get config path directly from token
        CSpDynamicString configPathStr;
        HRESULT hr = m_pToken->GetStringValue(L"ConfigPath", &configPathStr);
        if (FAILED(hr) || !configPathStr)
        {
            LogMessage(L"Failed to get ConfigPath from token");
            return L"";
        }

        std::wstring configPath = static_cast<LPCWSTR>(configPathStr);
        LogMessage((L"Loading configuration from: " + configPath).c_str());

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

        std::string argsContent = configContent.substr(braceStart, pos - braceStart);

        // Check if listvoices is already present
        if (argsContent.find("\"listvoices\"") == std::string::npos)
        {
            // Add listvoices parameter before the closing brace
            size_t closingBrace = argsContent.find_last_of('}');
            if (closingBrace != std::string::npos)
            {
                // Insert listvoices before the closing brace
                std::string beforeClosing = argsContent.substr(0, closingBrace);
                // Add comma if there are other parameters
                if (beforeClosing.find_last_not_of(" \t\n\r") != std::string::npos &&
                    beforeClosing.back() != '{')
                {
                    beforeClosing += ",\n      \"listvoices\": false\n    }";
                }
                else
                {
                    beforeClosing += "\n      \"listvoices\": false\n    }";
                }
                argsContent = beforeClosing;
            }
        }

        // Check if return_audio_bytes is already present
        if (argsContent.find("\"return_audio_bytes\"") == std::string::npos)
        {
            // Add return_audio_bytes parameter before the closing brace
            size_t closingBrace = argsContent.find_last_of('}');
            if (closingBrace != std::string::npos)
            {
                // Insert return_audio_bytes before the closing brace
                std::string beforeClosing = argsContent.substr(0, closingBrace);
                // Add comma if there are other parameters
                if (beforeClosing.find_last_not_of(" \t\n\r") != std::string::npos &&
                    beforeClosing.back() != '{')
                {
                    beforeClosing += ",\n      \"return_audio_bytes\": true\n    }";
                }
                else
                {
                    beforeClosing += "\n      \"return_audio_bytes\": true\n    }";
                }
                argsContent = beforeClosing;
            }
        }

        std::string argsSection = "\"args\":" + argsContent;
        LogMessage((L"Loaded voice configuration: " + UTF8ToWString(argsSection)).c_str());

        return UTF8ToWString(argsSection);
    }
    catch (...)
    {
        LogMessage(L"Exception in LoadVoiceConfiguration");
        return L"";
    }
}

std::string CNativeTTSWrapper::CreateConfigSection()
{
    try
    {
        // Load the full voice configuration to extract engine info
        if (!m_pToken)
        {
            LogMessage(L"No token available for config creation");
            return CreateDefaultConfig();
        }

        // Get config path directly from token
        CSpDynamicString configPathStr;
        HRESULT hr = m_pToken->GetStringValue(L"ConfigPath", &configPathStr);
        if (FAILED(hr) || !configPathStr)
        {
            LogMessage(L"Failed to get ConfigPath for config creation");
            return CreateDefaultConfig();
        }

        std::wstring configPath = static_cast<LPCWSTR>(configPathStr);

        std::ifstream configFile(configPath);
        if (!configFile.is_open())
        {
            LogMessage((L"Voice config file not found for config creation: " + configPath).c_str());
            return CreateDefaultConfig();
        }

        std::string configContent((std::istreambuf_iterator<char>(configFile)),
                                 std::istreambuf_iterator<char>());
        configFile.close();

        // Extract engine from ttsConfig.args.engine
        size_t engineStart = configContent.find("\"engine\":");
        if (engineStart == std::string::npos)
        {
            LogMessage(L"No engine found in voice config");
            return CreateDefaultConfig();
        }

        // Find the engine value
        size_t quoteStart = configContent.find("\"", engineStart + 9);
        size_t quoteEnd = configContent.find("\"", quoteStart + 1);
        if (quoteStart == std::string::npos || quoteEnd == std::string::npos)
        {
            LogMessage(L"Invalid engine format in voice config");
            return CreateDefaultConfig();
        }

        std::string engine = configContent.substr(quoteStart + 1, quoteEnd - quoteStart - 1);
        LogMessage((L"Extracted engine from config: " + UTF8ToWString(engine)).c_str());

        // Create config based on engine type
        if (engine == "sherpaonnx")
        {
            return CreateSherpaOnnxConfig();
        }
        else if (engine == "azure" || engine == "microsoft")
        {
            return CreateAzureConfig();
        }
        else
        {
            LogMessage((L"Unknown engine type: " + UTF8ToWString(engine)).c_str());
            return CreateDefaultConfig();
        }
    }
    catch (...)
    {
        LogMessage(L"Exception in CreateConfigSection");
        return CreateDefaultConfig();
    }
}

std::string CNativeTTSWrapper::CreateDefaultConfig()
{
    return "{\n"
           "  \"TTS\": {\n"
           "    \"engine\": \"SherpaOnnxTTS\",\n"
           "    \"bypass_tts\": \"False\",\n"
           "    \"save_audio_file\": \"True\",\n"
           "    \"rate\": \"0\",\n"
           "    \"volume\": \"100\"\n"
           "  },\n"
           "  \"translate\": {\n"
           "    \"no_translate\": \"True\",\n"
           "    \"start_lang\": \"en\",\n"
           "    \"end_lang\": \"en\",\n"
           "    \"replace_pb\": \"True\"\n"
           "  },\n"
           "  \"App\": {\n"
           "    \"collectstats\": \"True\"\n"
           "  }\n"
           "}";
}

std::string CNativeTTSWrapper::CreateSherpaOnnxConfig()
{
    return "{\n"
           "  \"TTS\": {\n"
           "    \"engine\": \"SherpaOnnxTTS\",\n"
           "    \"bypass_tts\": \"False\",\n"
           "    \"save_audio_file\": \"True\",\n"
           "    \"rate\": \"0\",\n"
           "    \"volume\": \"100\"\n"
           "  },\n"
           "  \"translate\": {\n"
           "    \"no_translate\": \"True\",\n"
           "    \"start_lang\": \"en\",\n"
           "    \"end_lang\": \"en\",\n"
           "    \"replace_pb\": \"True\"\n"
           "  },\n"
           "  \"App\": {\n"
           "    \"collectstats\": \"True\"\n"
           "  }\n"
           "}";
}

std::string CNativeTTSWrapper::CreateAzureConfig()
{
    return "{\n"
           "  \"TTS\": {\n"
           "    \"engine\": \"azureTTS\",\n"
           "    \"bypass_tts\": \"False\",\n"
           "    \"save_audio_file\": \"True\",\n"
           "    \"rate\": \"0\",\n"
           "    \"volume\": \"100\"\n"
           "  },\n"
           "  \"translate\": {\n"
           "    \"no_translate\": \"True\",\n"
           "    \"start_lang\": \"en\",\n"
           "    \"end_lang\": \"en\",\n"
           "    \"replace_pb\": \"True\"\n"
           "  },\n"
           "  \"App\": {\n"
           "    \"collectstats\": \"True\"\n"
           "  },\n"
           "  \"azureTTS\": {\n"
           "    \"key\": \"b14f8945b0f1459f9964bdd72c42c2cc\",\n"
           "    \"location\": \"uksouth\",\n"
           "    \"voice_id\": \"en-GB-LibbyNeural\"\n"
           "  }\n"
           "}";
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
