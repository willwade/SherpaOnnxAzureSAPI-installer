#include "stdafx.h"
#include "SpVoice.h"
#include "NativeTTSWrapper.h"
#include <comdef.h>
#include <iostream>
#include <fstream>

// Constructor
COpenSpeechSpVoice::COpenSpeechSpVoice()
    : m_lRate(0)
    , m_usVolume(100)
    , m_ePriority(SPVPRI_NORMAL)
    , m_eAlertBoundary(SPEI_WORD_BOUNDARY)
    , m_ulSyncTimeout(INFINITE)
    , m_ullEventInterest(0)
    , m_hNotifyWnd(NULL)
    , m_uNotifyMsg(0)
    , m_hNotifyEvent(NULL)
    , m_bInitialized(false)
{
    LogMessage(L"COpenSpeechSpVoice::Constructor - Creating OpenSpeechSpVoice instance");

    // Initialize critical sections
    InitializeCriticalSection(&m_csEventQueue);
    InitializeCriticalSection(&m_csVoiceState);

    // Initialize default voice status
    InitializeDefaults();
}

// Destructor
COpenSpeechSpVoice::~COpenSpeechSpVoice()
{
    LogMessage(L"COpenSpeechSpVoice::Destructor - Cleaning up OpenSpeechSpVoice instance");

    // Cleanup critical sections
    DeleteCriticalSection(&m_csEventQueue);
    DeleteCriticalSection(&m_csVoiceState);

    // Close notification event if created
    if (m_hNotifyEvent)
    {
        CloseHandle(m_hNotifyEvent);
        m_hNotifyEvent = NULL;
    }
}

// Initialize default values
void COpenSpeechSpVoice::InitializeDefaults()
{
    ZeroMemory(&m_voiceStatus, sizeof(SPVOICESTATUS));
    m_voiceStatus.ulCurrentStream = 0;
    m_voiceStatus.ulLastStreamQueued = 0;
    m_voiceStatus.hrLastResult = S_OK;
    m_voiceStatus.dwRunningState = SPRS_DONE;
    m_voiceStatus.ulInputWordPos = 0;
    m_voiceStatus.ulInputSentPos = 0;
    m_voiceStatus.lBookmarkId = 0;
    m_voiceStatus.PhonemeId = 0;
    m_voiceStatus.VisemeId = static_cast<SPVISEMES>(0);
    m_voiceStatus.dwReserved1 = 0;
    m_voiceStatus.dwReserved2 = 0;
}

// Logging helper
void COpenSpeechSpVoice::LogMessage(const wchar_t* message)
{
    try
    {
        std::wofstream logFile(L"C:\\temp\\SpVoice.log", std::ios::app);
        if (logFile.is_open())
        {
            SYSTEMTIME st;
            GetLocalTime(&st);
            logFile << L"[" << st.wHour << L":" << st.wMinute << L":" << st.wSecond << L"] " << message << std::endl;
            logFile.close();
        }
    }
    catch (...)
    {
        // Ignore logging errors
    }
}

// CRITICAL: Main speech synthesis method that Grid 3 will call
STDMETHODIMP COpenSpeechSpVoice::Speak(LPCWSTR pwcs, DWORD dwFlags, ULONG* pulStreamNumber)
{
    LogMessage(L"COpenSpeechSpVoice::Speak - Called by application");

    if (!pwcs)
    {
        LogMessage(L"COpenSpeechSpVoice::Speak - NULL text pointer");
        return E_INVALIDARG;
    }

    try
    {
        std::wstring logMsg = L"COpenSpeechSpVoice::Speak - Text: ";
        logMsg += pwcs;
        LogMessage(logMsg.c_str());

        // Update voice status to indicate we're speaking
        EnterCriticalSection(&m_csVoiceState);
        UpdateVoiceStatus(SPRS_IS_SPEAKING);
        LeaveCriticalSection(&m_csVoiceState);

        // Initialize TTS engine if not already done
        HRESULT hr = InitializeTTSEngine();
        if (FAILED(hr))
        {
            LogMessage(L"COpenSpeechSpVoice::Speak - Failed to initialize TTS engine");
            return hr;
        }

        // Convert text to SPVTEXTFRAG format for ISpTTSEngine
        SPVTEXTFRAG* pFragments = nullptr;
        hr = ConvertTextToFragments(pwcs, &pFragments);
        if (FAILED(hr))
        {
            LogMessage(L"COpenSpeechSpVoice::Speak - Failed to convert text to fragments");
            return hr;
        }

        // Call our existing ISpTTSEngine implementation
        LogMessage(L"COpenSpeechSpVoice::Speak - Delegating to ISpTTSEngine");
        hr = m_pTTSEngine->Speak(
            dwFlags,
            SPDFID_WaveFormatEx,
            nullptr,  // Use default wave format
            pFragments,
            nullptr   // We'll implement ISpTTSEngineSite later
        );

        // Cleanup
        CleanupTextFragments(pFragments);

        // Update status based on result
        EnterCriticalSection(&m_csVoiceState);
        if (SUCCEEDED(hr))
        {
            UpdateVoiceStatus(SPRS_DONE);
            LogMessage(L"COpenSpeechSpVoice::Speak - Speech completed successfully");
        }
        else
        {
            UpdateVoiceStatus(SPRS_DONE);
            LogMessage(L"COpenSpeechSpVoice::Speak - Speech failed");
        }
        LeaveCriticalSection(&m_csVoiceState);

        // Set stream number if requested
        if (pulStreamNumber)
        {
            *pulStreamNumber = m_voiceStatus.ulCurrentStream;
        }

        return hr;
    }
    catch (...)
    {
        LogMessage(L"COpenSpeechSpVoice::Speak - Exception occurred");
        EnterCriticalSection(&m_csVoiceState);
        UpdateVoiceStatus(SPRS_DONE);
        LeaveCriticalSection(&m_csVoiceState);
        return E_FAIL;
    }
}

// Initialize TTS Engine
HRESULT COpenSpeechSpVoice::InitializeTTSEngine()
{
    if (m_pTTSEngine && m_bInitialized)
    {
        return S_OK;  // Already initialized
    }
    
    LogMessage(L"COpenSpeechSpVoice::InitializeTTSEngine - Creating TTS engine");
    
    try
    {
        // Create our existing CNativeTTSWrapper instance
        HRESULT hr = CoCreateInstance(CLSID_CNativeTTSWrapper, nullptr, CLSCTX_INPROC_SERVER, IID_ISpTTSEngine, (void**)&m_pTTSEngine);
        if (FAILED(hr))
        {
            LogMessage(L"COpenSpeechSpVoice::InitializeTTSEngine - Failed to create TTS engine");
            return hr;
        }
        
        // If we have a voice token, set it on the TTS engine
        if (m_pVoiceToken)
        {
            CComQIPtr<ISpObjectWithToken> pObjectWithToken(m_pTTSEngine);
            if (pObjectWithToken)
            {
                hr = pObjectWithToken->SetObjectToken(m_pVoiceToken);
                if (FAILED(hr))
                {
                    LogMessage(L"COpenSpeechSpVoice::InitializeTTSEngine - Failed to set voice token");
                    return hr;
                }
            }
        }
        
        m_bInitialized = true;
        LogMessage(L"COpenSpeechSpVoice::InitializeTTSEngine - TTS engine initialized successfully");
        return S_OK;
    }
    catch (...)
    {
        LogMessage(L"COpenSpeechSpVoice::InitializeTTSEngine - Exception occurred");
        return E_FAIL;
    }
}

// Convert text to SPVTEXTFRAG format
HRESULT COpenSpeechSpVoice::ConvertTextToFragments(LPCWSTR pwcs, SPVTEXTFRAG** ppFragments)
{
    if (!pwcs || !ppFragments)
        return E_INVALIDARG;
    
    try
    {
        // Allocate single fragment for the entire text
        SPVTEXTFRAG* pFrag = new SPVTEXTFRAG;
        ZeroMemory(pFrag, sizeof(SPVTEXTFRAG));
        
        pFrag->pNext = nullptr;
        pFrag->pTextStart = pwcs;
        pFrag->ulTextLen = static_cast<ULONG>(wcslen(pwcs));
        pFrag->ulTextSrcOffset = 0;
        
        // Initialize state with default values
        pFrag->State.eAction = SPVA_Speak;
        pFrag->State.LangID = MAKELANGID(LANG_ENGLISH, SUBLANG_ENGLISH_US);
        pFrag->State.wReserved = 0;
        pFrag->State.EmphAdj = 0;
        pFrag->State.RateAdj = m_lRate;
        pFrag->State.Volume = m_usVolume;
        pFrag->State.PitchAdj.MiddleAdj = 0;
        pFrag->State.PitchAdj.RangeAdj = 0;
        pFrag->State.SilenceMSecs = 0;
        pFrag->State.pPhoneIds = nullptr;
        pFrag->State.ePartOfSpeech = SPPS_Unknown;
        pFrag->State.Context.pCategory = nullptr;
        pFrag->State.Context.pBefore = nullptr;
        pFrag->State.Context.pAfter = nullptr;
        
        *ppFragments = pFrag;
        return S_OK;
    }
    catch (...)
    {
        return E_OUTOFMEMORY;
    }
}

// Cleanup text fragments
void COpenSpeechSpVoice::CleanupTextFragments(SPVTEXTFRAG* pFragments)
{
    if (pFragments)
    {
        delete pFragments;
    }
}

// Update voice status
void COpenSpeechSpVoice::UpdateVoiceStatus(SPRUNSTATE eRunState)
{
    m_voiceStatus.dwRunningState = eRunState;
    
    if (eRunState == SPRS_IS_SPEAKING)
    {
        m_voiceStatus.ulCurrentStream++;
        m_voiceStatus.ulLastStreamQueued = m_voiceStatus.ulCurrentStream;
    }
}

// Get current voice status
STDMETHODIMP COpenSpeechSpVoice::GetStatus(SPVOICESTATUS* pStatus, LPWSTR* ppszLastBookmark)
{
    if (!pStatus)
        return E_INVALIDARG;

    EnterCriticalSection(&m_csVoiceState);
    *pStatus = m_voiceStatus;
    LeaveCriticalSection(&m_csVoiceState);

    if (ppszLastBookmark)
    {
        *ppszLastBookmark = nullptr; // No bookmark support yet
    }

    return S_OK;
}

// Set voice token
STDMETHODIMP COpenSpeechSpVoice::SetVoice(ISpObjectToken* pToken)
{
    LogMessage(L"COpenSpeechSpVoice::SetVoice - Setting voice token");
    
    m_pVoiceToken = pToken;
    
    // If TTS engine is already initialized, update its token
    if (m_pTTSEngine && m_bInitialized)
    {
        CComQIPtr<ISpObjectWithToken> pObjectWithToken(m_pTTSEngine);
        if (pObjectWithToken)
        {
            return pObjectWithToken->SetObjectToken(pToken);
        }
    }
    
    return S_OK;
}

// Get voice token
STDMETHODIMP COpenSpeechSpVoice::GetVoice(ISpObjectToken** ppToken)
{
    if (!ppToken)
        return E_INVALIDARG;

    if (m_pVoiceToken)
    {
        return m_pVoiceToken.CopyTo(ppToken);
    }

    *ppToken = nullptr;
    return S_FALSE;
}

// Rate control
STDMETHODIMP COpenSpeechSpVoice::SetRate(long RateAdjust)
{
    LogMessage(L"COpenSpeechSpVoice::SetRate - Setting speech rate");
    m_lRate = RateAdjust;
    return S_OK;
}

STDMETHODIMP COpenSpeechSpVoice::GetRate(long* pRateAdjust)
{
    if (!pRateAdjust)
        return E_INVALIDARG;

    *pRateAdjust = m_lRate;
    return S_OK;
}

// Volume control
STDMETHODIMP COpenSpeechSpVoice::SetVolume(USHORT usVolume)
{
    LogMessage(L"COpenSpeechSpVoice::SetVolume - Setting volume");
    m_usVolume = usVolume;
    return S_OK;
}

STDMETHODIMP COpenSpeechSpVoice::GetVolume(USHORT* pusVolume)
{
    if (!pusVolume)
        return E_INVALIDARG;

    *pusVolume = m_usVolume;
    return S_OK;
}

// Flow control
STDMETHODIMP COpenSpeechSpVoice::Pause(void)
{
    LogMessage(L"COpenSpeechSpVoice::Pause - Pausing speech");
    EnterCriticalSection(&m_csVoiceState);
    UpdateVoiceStatus(SPRS_DONE); // Use DONE instead of PAUSED for now
    LeaveCriticalSection(&m_csVoiceState);
    return S_OK;
}

STDMETHODIMP COpenSpeechSpVoice::Resume(void)
{
    LogMessage(L"COpenSpeechSpVoice::Resume - Resuming speech");
    EnterCriticalSection(&m_csVoiceState);
    UpdateVoiceStatus(SPRS_IS_SPEAKING);
    LeaveCriticalSection(&m_csVoiceState);
    return S_OK;
}

// Priority management
STDMETHODIMP COpenSpeechSpVoice::SetPriority(SPVPRIORITY ePriority)
{
    m_ePriority = ePriority;
    return S_OK;
}

STDMETHODIMP COpenSpeechSpVoice::GetPriority(SPVPRIORITY* pePriority)
{
    if (!pePriority)
        return E_INVALIDARG;

    *pePriority = m_ePriority;
    return S_OK;
}

// Alert boundary
STDMETHODIMP COpenSpeechSpVoice::SetAlertBoundary(SPEVENTENUM eBoundary)
{
    m_eAlertBoundary = eBoundary;
    return S_OK;
}

STDMETHODIMP COpenSpeechSpVoice::GetAlertBoundary(SPEVENTENUM* peBoundary)
{
    if (!peBoundary)
        return E_INVALIDARG;

    *peBoundary = m_eAlertBoundary;
    return S_OK;
}

// Synchronization
STDMETHODIMP COpenSpeechSpVoice::WaitUntilDone(ULONG msTimeout)
{
    LogMessage(L"COpenSpeechSpVoice::WaitUntilDone - Waiting for speech completion");

    // Simple implementation - check status periodically
    DWORD startTime = GetTickCount();
    while (true)
    {
        EnterCriticalSection(&m_csVoiceState);
        SPRUNSTATE state = (SPRUNSTATE)m_voiceStatus.dwRunningState;
        LeaveCriticalSection(&m_csVoiceState);

        if (state == SPRS_DONE)
        {
            return S_OK;
        }

        if (msTimeout != INFINITE)
        {
            DWORD elapsed = GetTickCount() - startTime;
            if (elapsed >= msTimeout)
            {
                return S_FALSE; // Timeout
            }
        }

        Sleep(50); // Check every 50ms
    }
}

STDMETHODIMP COpenSpeechSpVoice::SetSyncSpeakTimeout(ULONG msTimeout)
{
    m_ulSyncTimeout = msTimeout;
    return S_OK;
}

STDMETHODIMP COpenSpeechSpVoice::GetSyncSpeakTimeout(ULONG* pmsTimeout)
{
    if (!pmsTimeout)
        return E_INVALIDARG;

    *pmsTimeout = m_ulSyncTimeout;
    return S_OK;
}

// Event handling
STDMETHODIMP_(HANDLE) COpenSpeechSpVoice::SpeakCompleteEvent(void)
{
    if (!m_hNotifyEvent)
    {
        m_hNotifyEvent = CreateEvent(NULL, TRUE, FALSE, NULL);
        if (!m_hNotifyEvent)
            return NULL;
    }

    return m_hNotifyEvent;
}

// Stream-based speech (basic implementation)
STDMETHODIMP COpenSpeechSpVoice::SpeakStream(IStream* pStream, DWORD dwFlags, ULONG* pulStreamNumber)
{
    LogMessage(L"COpenSpeechSpVoice::SpeakStream - Stream-based speech requested");

    if (!pStream)
        return E_INVALIDARG;

    // For now, we'll implement a basic version that reads the stream as text
    // In a full implementation, this would handle various stream formats
    try
    {
        // Read stream content
        STATSTG stat;
        HRESULT hr = pStream->Stat(&stat, STATFLAG_NONAME);
        if (FAILED(hr))
            return hr;

        ULONG streamSize = static_cast<ULONG>(stat.cbSize.QuadPart);
        if (streamSize == 0)
            return S_OK;

        // Allocate buffer for stream content
        std::vector<char> buffer(streamSize + 1);
        ULONG bytesRead = 0;
        hr = pStream->Read(buffer.data(), streamSize, &bytesRead);
        if (FAILED(hr))
            return hr;

        buffer[bytesRead] = '\0';

        // Convert to wide string (assuming UTF-8 input)
        int wideSize = MultiByteToWideChar(CP_UTF8, 0, buffer.data(), bytesRead, NULL, 0);
        if (wideSize == 0)
            return E_FAIL;

        std::vector<wchar_t> wideBuffer(wideSize + 1);
        MultiByteToWideChar(CP_UTF8, 0, buffer.data(), bytesRead, wideBuffer.data(), wideSize);
        wideBuffer[wideSize] = L'\0';

        // Delegate to regular Speak method
        return Speak(wideBuffer.data(), dwFlags, pulStreamNumber);
    }
    catch (...)
    {
        LogMessage(L"COpenSpeechSpVoice::SpeakStream - Exception occurred");
        return E_FAIL;
    }
}

// Skip functionality (basic implementation)
STDMETHODIMP COpenSpeechSpVoice::Skip(LPCWSTR pItemType, long lNumItems, ULONG* pulNumSkipped)
{
    LogMessage(L"COpenSpeechSpVoice::Skip - Skip requested");

    // Basic implementation - for now just report that we skipped the requested items
    if (pulNumSkipped)
    {
        *pulNumSkipped = abs(lNumItems);
    }

    return S_OK;
}

// UI Support
STDMETHODIMP COpenSpeechSpVoice::IsUISupported(LPCWSTR pszTypeOfUI, void* pvExtraData, ULONG cbExtraData, BOOL* pfSupported)
{
    if (!pfSupported)
        return E_INVALIDARG;

    // For now, we don't support any UI
    *pfSupported = FALSE;
    return S_OK;
}

STDMETHODIMP COpenSpeechSpVoice::DisplayUI(HWND hwndParent, LPCWSTR pszTitle, LPCWSTR pszTypeOfUI, void* pvExtraData, ULONG cbExtraData)
{
    // No UI supported yet
    return E_NOTIMPL;
}

// Audio output control (basic implementation)
STDMETHODIMP COpenSpeechSpVoice::SetOutput(IUnknown* pUnkOutput, BOOL fAllowFormatChanges)
{
    LogMessage(L"COpenSpeechSpVoice::SetOutput - Audio output change requested");
    // For now, we'll use the default output through AACSpeakHelper
    return S_OK;
}

STDMETHODIMP COpenSpeechSpVoice::GetOutputStream(ISpStreamFormat** ppStream)
{
    if (!ppStream)
        return E_INVALIDARG;

    *ppStream = nullptr;
    return S_FALSE; // No stream available
}

STDMETHODIMP COpenSpeechSpVoice::GetOutputObjectToken(ISpObjectToken** ppObjectToken)
{
    if (!ppObjectToken)
        return E_INVALIDARG;

    *ppObjectToken = nullptr;
    return S_FALSE; // No output token available
}

// ISpEventSource implementation
STDMETHODIMP COpenSpeechSpVoice::SetInterest(ULONGLONG ullEventInterest, ULONGLONG ullQueuedInterest)
{
    LogMessage(L"COpenSpeechSpVoice::SetInterest - Setting event interest");
    m_ullEventInterest = ullEventInterest;
    return S_OK;
}

STDMETHODIMP COpenSpeechSpVoice::GetEvents(ULONG ulCount, SPEVENT* pEventArray, ULONG* pulFetched)
{
    if (!pEventArray || !pulFetched)
        return E_INVALIDARG;

    EnterCriticalSection(&m_csEventQueue);

    ULONG eventsFetched = 0;
    ULONG eventsToFetch = min(ulCount, static_cast<ULONG>(m_eventQueue.size()));

    for (ULONG i = 0; i < eventsToFetch; i++)
    {
        pEventArray[i] = m_eventQueue[i];
        eventsFetched++;
    }

    // Remove fetched events from queue
    if (eventsFetched > 0)
    {
        m_eventQueue.erase(m_eventQueue.begin(), m_eventQueue.begin() + eventsFetched);
    }

    LeaveCriticalSection(&m_csEventQueue);

    *pulFetched = eventsFetched;
    return S_OK;
}

STDMETHODIMP COpenSpeechSpVoice::GetInfo(SPEVENTSOURCEINFO* pInfo)
{
    if (!pInfo)
        return E_INVALIDARG;

    EnterCriticalSection(&m_csEventQueue);
    pInfo->ulCount = static_cast<ULONG>(m_eventQueue.size());
    LeaveCriticalSection(&m_csEventQueue);

    pInfo->ullEventInterest = m_ullEventInterest;
    return S_OK;
}

// ISpNotifySource implementation
STDMETHODIMP COpenSpeechSpVoice::SetNotifySink(ISpNotifySink* pNotifySink)
{
    m_pNotifySink = pNotifySink;
    return S_OK;
}

STDMETHODIMP COpenSpeechSpVoice::SetNotifyWindowMessage(HWND hWnd, UINT Msg, WPARAM wParam, LPARAM lParam)
{
    m_hNotifyWnd = hWnd;
    m_uNotifyMsg = Msg;
    return S_OK;
}

STDMETHODIMP COpenSpeechSpVoice::SetNotifyCallbackFunction(SPNOTIFYCALLBACK* pfnCallback, WPARAM wParam, LPARAM lParam)
{
    // Not implemented yet
    return E_NOTIMPL;
}

STDMETHODIMP COpenSpeechSpVoice::SetNotifyCallbackInterface(ISpNotifyCallback* pSpCallback, WPARAM wParam, LPARAM lParam)
{
    // Not implemented yet
    return E_NOTIMPL;
}

STDMETHODIMP COpenSpeechSpVoice::SetNotifyWin32Event(void)
{
    if (!m_hNotifyEvent)
    {
        m_hNotifyEvent = CreateEvent(NULL, FALSE, FALSE, NULL);
        if (!m_hNotifyEvent)
            return E_FAIL;
    }
    return S_OK;
}

STDMETHODIMP COpenSpeechSpVoice::WaitForNotifyEvent(DWORD dwMilliseconds)
{
    if (!m_hNotifyEvent)
        return E_FAIL;

    DWORD result = WaitForSingleObject(m_hNotifyEvent, dwMilliseconds);
    return (result == WAIT_OBJECT_0) ? S_OK : S_FALSE;
}

STDMETHODIMP_(HANDLE) COpenSpeechSpVoice::GetNotifyEventHandle(void)
{
    return m_hNotifyEvent;
}

// Helper method to fire events
HRESULT COpenSpeechSpVoice::FireEvent(SPEVENTENUM eEventId, WPARAM wParam, LPARAM lParam)
{
    // Check if we're interested in this event
    if (!(m_ullEventInterest & (1ULL << eEventId)))
        return S_OK;

    // Create event
    SPEVENT event = {0};
    event.eEventId = eEventId;
    event.elParamType = SPET_LPARAM_IS_UNDEFINED;
    event.ulStreamNum = m_voiceStatus.ulCurrentStream;
    event.ullAudioStreamOffset = 0; // Would need actual audio position
    event.wParam = wParam;
    event.lParam = lParam;

    // Add to event queue
    EnterCriticalSection(&m_csEventQueue);
    m_eventQueue.push_back(event);
    LeaveCriticalSection(&m_csEventQueue);

    // Notify if we have notification set up
    ProcessNotification();

    return S_OK;
}

// Process notification
HRESULT COpenSpeechSpVoice::ProcessNotification()
{
    if (m_pNotifySink)
    {
        m_pNotifySink->Notify();
    }

    if (m_hNotifyWnd && m_uNotifyMsg)
    {
        PostMessage(m_hNotifyWnd, m_uNotifyMsg, 0, 0);
    }

    if (m_hNotifyEvent)
    {
        SetEvent(m_hNotifyEvent);
    }

    return S_OK;
}
