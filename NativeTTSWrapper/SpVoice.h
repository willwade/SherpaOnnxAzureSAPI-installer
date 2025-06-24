#pragma once

#include "resource.h"
#include <windows.h>
#include <sapi.h>
#include <sphelper.h>
#include <atlbase.h>
#include <atlcom.h>
#include <string>
#include <vector>
#include <memory>
#include "NativeTTSWrapper.h"

// CLSID for OpenSpeechSpVoice - This should match the UUID in the IDL file
// {F2E8B6A1-3C4D-4E5F-8A7B-9C1D2E3F4A5B}
DEFINE_GUID(CLSID_OpenSpeechSpVoice,
    0xF2E8B6A1, 0x3C4D, 0x4E5F, 0x8A, 0x7B, 0x9C, 0x1D, 0x2E, 0x3F, 0x4A, 0x5B);

// Forward declarations
class CNativeTTSWrapper;

// COpenSpeechSpVoice - Application-level SAPI interface that Grid 3 expects
// This bridges the gap between SAPI applications and our ISpTTSEngine implementation
class ATL_NO_VTABLE COpenSpeechSpVoice :
    public CComObjectRootEx<CComMultiThreadModel>,
    public CComCoClass<COpenSpeechSpVoice, &CLSID_OpenSpeechSpVoice>,
    public ISpVoice,
    public ISpObjectWithToken
{
public:
    COpenSpeechSpVoice();
    virtual ~COpenSpeechSpVoice();

    DECLARE_REGISTRY_RESOURCEID(IDR_SPVOICE)
    DECLARE_NOT_AGGREGATABLE(COpenSpeechSpVoice)

    BEGIN_COM_MAP(COpenSpeechSpVoice)
        COM_INTERFACE_ENTRY(ISpVoice)
        COM_INTERFACE_ENTRY(ISpEventSource)
        COM_INTERFACE_ENTRY(ISpNotifySource)
        COM_INTERFACE_ENTRY(ISpObjectWithToken)
    END_COM_MAP()

    // ISpVoice implementation - Core speech synthesis interface
    STDMETHOD(Speak)(
        LPCWSTR pwcs,
        DWORD dwFlags,
        ULONG* pulStreamNumber) override;

    STDMETHOD(SpeakStream)(
        IStream* pStream,
        DWORD dwFlags,
        ULONG* pulStreamNumber) override;

    STDMETHOD(GetStatus)(
        SPVOICESTATUS* pStatus,
        LPWSTR* ppszLastBookmark) override;

    STDMETHOD(Skip)(
        LPCWSTR pItemType,
        long lNumItems,
        ULONG* pulNumSkipped) override;

    STDMETHOD(SetPriority)(
        SPVPRIORITY ePriority) override;

    STDMETHOD(GetPriority)(
        SPVPRIORITY* pePriority) override;

    STDMETHOD(SetAlertBoundary)(
        SPEVENTENUM eBoundary) override;

    STDMETHOD(GetAlertBoundary)(
        SPEVENTENUM* peBoundary) override;

    STDMETHOD(SetRate)(
        long RateAdjust) override;

    STDMETHOD(GetRate)(
        long* pRateAdjust) override;

    STDMETHOD(SetVolume)(
        USHORT usVolume) override;

    STDMETHOD(GetVolume)(
        USHORT* pusVolume) override;

    STDMETHOD(WaitUntilDone)(
        ULONG msTimeout) override;

    STDMETHOD(SetSyncSpeakTimeout)(
        ULONG msTimeout) override;

    STDMETHOD(GetSyncSpeakTimeout)(
        ULONG* pmsTimeout) override;

    STDMETHOD_(HANDLE, SpeakCompleteEvent)(void) override;

    STDMETHOD(IsUISupported)(
        LPCWSTR pszTypeOfUI,
        void* pvExtraData,
        ULONG cbExtraData,
        BOOL* pfSupported) override;

    STDMETHOD(DisplayUI)(
        HWND hwndParent,
        LPCWSTR pszTitle,
        LPCWSTR pszTypeOfUI,
        void* pvExtraData,
        ULONG cbExtraData) override;

    // Voice selection and management
    STDMETHOD(SetVoice)(
        ISpObjectToken* pToken) override;

    STDMETHOD(GetVoice)(
        ISpObjectToken** ppToken) override;

    // Audio output control
    STDMETHOD(SetOutput)(
        IUnknown* pUnkOutput,
        BOOL fAllowFormatChanges) override;

    STDMETHOD(GetOutputStream)(
        ISpStreamFormat** ppStream) override;

    STDMETHOD(GetOutputObjectToken)(
        ISpObjectToken** ppObjectToken) override;

    // Flow control
    STDMETHOD(Pause)(void) override;
    STDMETHOD(Resume)(void) override;

    // ISpObjectWithToken implementation - Required for SAPI voice objects
    STDMETHOD(SetObjectToken)(ISpObjectToken* pToken) override;
    STDMETHOD(GetObjectToken)(ISpObjectToken** ppToken) override;

    // ISpEventSource implementation - Event management for speech synthesis
    // (These are inherited from ISpVoice which inherits from ISpEventSource)
    STDMETHOD(SetInterest)(
        ULONGLONG ullEventInterest,
        ULONGLONG ullQueuedInterest) override;

    STDMETHOD(GetEvents)(
        ULONG ulCount,
        SPEVENT* pEventArray,
        ULONG* pulFetched) override;

    STDMETHOD(GetInfo)(
        SPEVENTSOURCEINFO* pInfo) override;

    // ISpNotifySource implementation - Notification management
    // (These are inherited from ISpVoice which inherits from ISpNotifySource)
    STDMETHOD(SetNotifySink)(
        ISpNotifySink* pNotifySink) override;

    STDMETHOD(SetNotifyWindowMessage)(
        HWND hWnd,
        UINT Msg,
        WPARAM wParam,
        LPARAM lParam) override;

    STDMETHOD(SetNotifyCallbackFunction)(
        SPNOTIFYCALLBACK* pfnCallback,
        WPARAM wParam,
        LPARAM lParam) override;

    STDMETHOD(SetNotifyCallbackInterface)(
        ISpNotifyCallback* pSpCallback,
        WPARAM wParam,
        LPARAM lParam) override;

    STDMETHOD(SetNotifyWin32Event)(void) override;

    STDMETHOD(WaitForNotifyEvent)(
        DWORD dwMilliseconds) override;

    STDMETHOD_(HANDLE, GetNotifyEventHandle)(void) override;

private:
    // Member variables
    CComPtr<ISpObjectToken> m_pVoiceToken;          // Current voice token
    CComPtr<CNativeTTSWrapper> m_pTTSEngine;        // Our existing TTS engine
    
    // Voice state
    SPVOICESTATUS m_voiceStatus;                    // Current voice status
    long m_lRate;                                   // Current speech rate
    USHORT m_usVolume;                              // Current volume
    SPVPRIORITY m_ePriority;                        // Voice priority
    SPEVENTENUM m_eAlertBoundary;                   // Alert boundary
    ULONG m_ulSyncTimeout;                          // Synchronous timeout
    
    // Event management
    ULONGLONG m_ullEventInterest;                   // Events we're interested in
    std::vector<SPEVENT> m_eventQueue;             // Event queue
    CRITICAL_SECTION m_csEventQueue;               // Event queue protection
    
    // Notification
    CComPtr<ISpNotifySink> m_pNotifySink;          // Notification sink
    HWND m_hNotifyWnd;                             // Notification window
    UINT m_uNotifyMsg;                             // Notification message
    HANDLE m_hNotifyEvent;                         // Notification event
    
    // Helper methods
    void LogMessage(const wchar_t* message);
    HRESULT InitializeTTSEngine();
    HRESULT CreateTTSEngineForVoice(ISpObjectToken* pVoiceToken);
    HRESULT ConvertTextToFragments(LPCWSTR pwcs, SPVTEXTFRAG** ppFragments);
    void CleanupTextFragments(SPVTEXTFRAG* pFragments);
    HRESULT FireEvent(SPEVENTENUM eEventId, WPARAM wParam, LPARAM lParam);
    HRESULT ProcessNotification();

    // AACSpeakHelper pipe service methods
    HRESULT CallAACSpeakHelperPipeService(const std::wstring& text, std::vector<BYTE>& audioData);
    bool ConnectToAACSpeakHelper(HANDLE& hPipe);
    HRESULT SendTextToPipe(HANDLE hPipe, const std::wstring& text);
    HRESULT ReceiveAudioFromPipe(HANDLE hPipe, std::vector<BYTE>& audioData);
    
    // State management
    void InitializeDefaults();
    void UpdateVoiceStatus(SPRUNSTATE eRunState);
    
    // Thread safety
    CRITICAL_SECTION m_csVoiceState;
    bool m_bInitialized;
};
