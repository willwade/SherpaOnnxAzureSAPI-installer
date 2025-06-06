#pragma once

#include "resource.h"
#include "ITTSEngine.h"
#include "TTSEngineManager.h"
#include <windows.h>
#include <sapi.h>
#include <sphelper.h>
#include <atlbase.h>
#include <atlcom.h>
#include <string>
#include <vector>
#include <memory>

// Native COM wrapper for Multi-Engine TTS System
// This provides the native COM interface that SAPI expects
// while delegating to our high-performance engine system (SherpaOnnx, Azure, etc.)
//
// ARCHITECTURE CHANGE:
// OLD: SAPI → COM → ProcessBridge → SherpaWorker.exe → .NET → SherpaOnnx
// NEW: SAPI → COM → EngineManager → Direct C++ Engines (SherpaOnnx C API, Azure C++ SDK)

// Forward declaration - CLSID defined in generated file
class CNativeTTSWrapper;

class ATL_NO_VTABLE CNativeTTSWrapper :
    public CComObjectRootEx<CComMultiThreadModel>,
    public CComCoClass<CNativeTTSWrapper, &CLSID_CNativeTTSWrapper>,
    public ISpTTSEngine,
    public ISpObjectWithToken
{
public:
    CNativeTTSWrapper();
    virtual ~CNativeTTSWrapper();

    DECLARE_REGISTRY_RESOURCEID(IDR_NATIVETTSRAPPER)
    DECLARE_NOT_AGGREGATABLE(CNativeTTSWrapper)

    BEGIN_COM_MAP(CNativeTTSWrapper)
        COM_INTERFACE_ENTRY(ISpTTSEngine)
        COM_INTERFACE_ENTRY(ISpObjectWithToken)
    END_COM_MAP()

    // ISpTTSEngine implementation
    STDMETHOD(Speak)(
        DWORD dwSpeakFlags,
        REFGUID rguidFormatId,
        const WAVEFORMATEX* pWaveFormatEx,
        const SPVTEXTFRAG* pTextFragList,
        ISpTTSEngineSite* pOutputSite) override;

    STDMETHOD(GetOutputFormat)(
        const GUID* pTargetFormatId,
        const WAVEFORMATEX* pTargetWaveFormatEx,
        GUID* pOutputFormatId,
        WAVEFORMATEX** ppCoMemOutputWaveFormatEx) override;

    // ISpObjectWithToken implementation
    STDMETHOD(SetObjectToken)(ISpObjectToken* pToken) override;
    STDMETHOD(GetObjectToken)(ISpObjectToken** ppToken) override;

private:
    // Member variables
    CComPtr<ISpObjectToken> m_pToken;
    std::wstring m_currentEngineId;  // ID of the engine for this voice
    bool m_engineInitialized;

    // Helper methods
    void LogMessage(const wchar_t* message);
    std::wstring ExtractTextFromFragList(const SPVTEXTFRAG* pTextFragList);

    // NEW: Direct engine integration methods
    HRESULT GenerateAudioViaNativeEngine(const std::wstring& text, std::vector<BYTE>& audioData);
    HRESULT InitializeEngineFromToken(ISpObjectToken* pToken);
    HRESULT ConvertFloatSamplesToBytes(const std::vector<float>& samples, int sampleRate, std::vector<BYTE>& audioData);

    // DIRECT: SherpaOnnx C API fallback (when engine manager fails)
    HRESULT GenerateAudioViaDirectSherpaOnnx(const std::wstring& text, std::vector<BYTE>& audioData);

    // LEGACY: ProcessBridge fallback methods (for backward compatibility)
    bool GenerateAudioViaProcessBridge(const std::wstring& text, std::vector<BYTE>& audioData);
    bool CallSherpaWorker(const std::wstring& requestPath, std::wstring& responsePath);
    bool ReadWavFile(const std::wstring& filePath, std::vector<BYTE>& audioData);

    // Utility methods
    std::string WStringToUTF8(const std::wstring& wstr);
    std::wstring UTF8ToWString(const std::string& str);
    std::wstring GenerateGuid();
    bool CreateTempDirectory(std::wstring& tempDir);
    bool WriteJsonRequest(const std::wstring& filePath, const std::wstring& text);
    bool ParseJsonResponse(const std::wstring& filePath, std::wstring& audioPath);
};


