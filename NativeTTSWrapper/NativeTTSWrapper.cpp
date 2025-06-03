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

// Implementation of CNativeTTSWrapper

CNativeTTSWrapper::CNativeTTSWrapper()
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

            // Call ProcessBridge to generate audio
            std::vector<BYTE> audioData;
            if (!GenerateAudioViaProcessBridge(text, audioData))
            {
                LogMessage(L"ProcessBridge audio generation failed");
                return E_FAIL;
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
            HRESULT hr = pOutputSite->Write(audioData.data(), (ULONG)audioData.size(), &bytesWritten);
            if (FAILED(hr))
            {
                LogMessage((L"Failed to write audio data: " + std::to_wstring(hr)).c_str());
                return hr;
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

            // Create JSON request
            std::string jsonRequest = "{\n";
            jsonRequest += "  \"Text\": \"" + utf8Text + "\",\n";
            jsonRequest += "  \"Speed\": 1.0,\n";
            jsonRequest += "  \"SpeakerId\": 0,\n";
            jsonRequest += "  \"OutputPath\": \"" + WStringToUTF8(audioPath) + "\"\n";
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

            // Read audio file
            std::ifstream audioFile(audioFilePathW, std::ios::binary);
            if (!audioFile.is_open())
            {
                LogMessage(L"Failed to open audio file");
                return false;
            }

            audioFile.seekg(0, std::ios::end);
            size_t fileSize = audioFile.tellg();
            audioFile.seekg(0, std::ios::beg);

            audioData.resize(fileSize);
            audioFile.read(reinterpret_cast<char*>(audioData.data()), fileSize);
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
