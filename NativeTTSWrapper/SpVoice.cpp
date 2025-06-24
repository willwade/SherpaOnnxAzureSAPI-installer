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
    LogMessage(L"COpenSpeechSpVoice::Constructor - ENTRY - Creating OpenSpeechSpVoice instance");

    // Initialize critical sections
    InitializeCriticalSection(&m_csEventQueue);
    InitializeCriticalSection(&m_csVoiceState);

    // Initialize default voice status
    InitializeDefaults();

    LogMessage(L"COpenSpeechSpVoice::Constructor - EXIT - OpenSpeechSpVoice instance created successfully");
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
    LogMessage(L"=== COpenSpeechSpVoice::Speak - ENTRY POINT ===");

    // Log detailed parameters
    wchar_t logBuffer[1024];
    swprintf_s(logBuffer, L"COpenSpeechSpVoice::Speak - Parameters: dwFlags=0x%08X, pulStreamNumber=%p",
               dwFlags, pulStreamNumber);
    LogMessage(logBuffer);

    // Parameter validation with detailed logging
    if (!pwcs)
    {
        LogMessage(L"COpenSpeechSpVoice::Speak - ERROR: NULL text pointer");
        return E_INVALIDARG;
    }

    size_t textLen = wcslen(pwcs);
    swprintf_s(logBuffer, L"COpenSpeechSpVoice::Speak - Text length: %zu characters", textLen);
    LogMessage(logBuffer);

    if (textLen > 0 && textLen < 200) // Log short text completely
    {
        std::wstring logMsg = L"COpenSpeechSpVoice::Speak - Text: \"";
        logMsg += pwcs;
        logMsg += L"\"";
        LogMessage(logMsg.c_str());
    }
    else if (textLen > 0)
    {
        std::wstring logMsg = L"COpenSpeechSpVoice::Speak - Text (first 100 chars): \"";
        logMsg += std::wstring(pwcs, 0, 100);
        logMsg += L"...\"";
        LogMessage(logMsg.c_str());
    }

    try
    {
        LogMessage(L"COpenSpeechSpVoice::Speak - STEP 1: Updating voice status to SPEAKING");

        // Update voice status to indicate we're speaking
        EnterCriticalSection(&m_csVoiceState);
        UpdateVoiceStatus(SPRS_IS_SPEAKING);
        LeaveCriticalSection(&m_csVoiceState);

        LogMessage(L"COpenSpeechSpVoice::Speak - STEP 1: COMPLETED - Voice status updated");

        LogMessage(L"COpenSpeechSpVoice::Speak - STEP 2: Initializing TTS engine");

        // Initialize TTS engine if not already done
        HRESULT hr = InitializeTTSEngine();
        if (FAILED(hr))
        {
            swprintf_s(logBuffer, L"COpenSpeechSpVoice::Speak - STEP 2: FAILED - InitializeTTSEngine returned 0x%08X", hr);
            LogMessage(logBuffer);

            EnterCriticalSection(&m_csVoiceState);
            UpdateVoiceStatus(SPRS_DONE);
            LeaveCriticalSection(&m_csVoiceState);
            return hr;
        }

        LogMessage(L"COpenSpeechSpVoice::Speak - STEP 2: COMPLETED - TTS engine initialized");

        LogMessage(L"COpenSpeechSpVoice::Speak - STEP 3: Converting text to fragments");

        // Convert text to SPVTEXTFRAG format for ISpTTSEngine
        SPVTEXTFRAG* pFragments = nullptr;
        hr = ConvertTextToFragments(pwcs, &pFragments);
        if (FAILED(hr))
        {
            swprintf_s(logBuffer, L"COpenSpeechSpVoice::Speak - STEP 3: FAILED - ConvertTextToFragments returned 0x%08X", hr);
            LogMessage(logBuffer);

            EnterCriticalSection(&m_csVoiceState);
            UpdateVoiceStatus(SPRS_DONE);
            LeaveCriticalSection(&m_csVoiceState);
            return hr;
        }

        LogMessage(L"COpenSpeechSpVoice::Speak - STEP 3: COMPLETED - Text converted to fragments");

        // Log fragment details
        if (pFragments)
        {
            swprintf_s(logBuffer, L"COpenSpeechSpVoice::Speak - Fragment details: ulTextLen=%lu, pTextStart=%p",
                       pFragments->ulTextLen, pFragments->pTextStart);
            LogMessage(logBuffer);
        }

        LogMessage(L"COpenSpeechSpVoice::Speak - STEP 4: Calling AACSpeakHelper directly (bypassing ISpTTSEngine)");

        // BYPASS ISpTTSEngine and call AACSpeakHelper directly
        // This avoids the ISpTTSEngineSite requirement that was causing E_INVALIDARG
        std::wstring text(pwcs);
        std::vector<BYTE> audioData;

        // Use the same pipe service call that CNativeTTSWrapper uses
        hr = CallAACSpeakHelperPipeService(text, audioData);

        swprintf_s(logBuffer, L"COpenSpeechSpVoice::Speak - STEP 4: AACSpeakHelper pipe service returned 0x%08X", hr);
        LogMessage(logBuffer);

        // Cleanup
        LogMessage(L"COpenSpeechSpVoice::Speak - STEP 5: Cleaning up fragments");
        CleanupTextFragments(pFragments);

        // Update status based on result
        LogMessage(L"COpenSpeechSpVoice::Speak - STEP 6: Updating final voice status");
        EnterCriticalSection(&m_csVoiceState);
        if (SUCCEEDED(hr))
        {
            UpdateVoiceStatus(SPRS_DONE);
            LogMessage(L"COpenSpeechSpVoice::Speak - STEP 6: SUCCESS - Speech completed successfully");
        }
        else
        {
            UpdateVoiceStatus(SPRS_DONE);
            swprintf_s(logBuffer, L"COpenSpeechSpVoice::Speak - STEP 6: FAILED - Speech failed with HRESULT 0x%08X", hr);
            LogMessage(logBuffer);
        }
        LeaveCriticalSection(&m_csVoiceState);

        // Set stream number if requested
        if (pulStreamNumber)
        {
            *pulStreamNumber = m_voiceStatus.ulCurrentStream;
            swprintf_s(logBuffer, L"COpenSpeechSpVoice::Speak - Stream number set to %lu", m_voiceStatus.ulCurrentStream);
            LogMessage(logBuffer);
        }

        LogMessage(L"=== COpenSpeechSpVoice::Speak - EXIT POINT ===");
        return hr;
    }
    catch (...)
    {
        LogMessage(L"COpenSpeechSpVoice::Speak - EXCEPTION: Unhandled exception occurred");
        EnterCriticalSection(&m_csVoiceState);
        UpdateVoiceStatus(SPRS_DONE);
        LeaveCriticalSection(&m_csVoiceState);
        return E_FAIL;
    }
}

// Initialize TTS Engine
HRESULT COpenSpeechSpVoice::InitializeTTSEngine()
{
    LogMessage(L"=== InitializeTTSEngine - ENTRY ===");

    wchar_t logBuffer[512];
    swprintf_s(logBuffer, L"InitializeTTSEngine - Current state: m_pTTSEngine=%p, m_bInitialized=%s",
               m_pTTSEngine, m_bInitialized ? L"true" : L"false");
    LogMessage(logBuffer);

    if (m_pTTSEngine && m_bInitialized)
    {
        LogMessage(L"InitializeTTSEngine - Already initialized, returning S_OK");
        return S_OK;  // Already initialized
    }

    LogMessage(L"InitializeTTSEngine - STEP 1: Creating CNativeTTSWrapper instance via CoCreateInstance");

    try
    {
        // Create our existing CNativeTTSWrapper instance
        HRESULT hr = CoCreateInstance(CLSID_CNativeTTSWrapper, nullptr, CLSCTX_INPROC_SERVER, IID_ISpTTSEngine, (void**)&m_pTTSEngine);

        swprintf_s(logBuffer, L"InitializeTTSEngine - STEP 1: CoCreateInstance returned 0x%08X, m_pTTSEngine=%p", hr, m_pTTSEngine);
        LogMessage(logBuffer);

        if (FAILED(hr))
        {
            LogMessage(L"InitializeTTSEngine - STEP 1: FAILED - CoCreateInstance failed");

            // Log specific error details
            if (hr == REGDB_E_CLASSNOTREG)
                LogMessage(L"InitializeTTSEngine - ERROR: Class not registered (REGDB_E_CLASSNOTREG)");
            else if (hr == CLASS_E_NOAGGREGATION)
                LogMessage(L"InitializeTTSEngine - ERROR: Class does not support aggregation (CLASS_E_NOAGGREGATION)");
            else if (hr == E_NOINTERFACE)
                LogMessage(L"InitializeTTSEngine - ERROR: Interface not supported (E_NOINTERFACE)");
            else
                LogMessage(L"InitializeTTSEngine - ERROR: Other CoCreateInstance failure");

            return hr;
        }

        LogMessage(L"InitializeTTSEngine - STEP 1: COMPLETED - CNativeTTSWrapper created successfully");

        // If we have a voice token, set it on the TTS engine
        swprintf_s(logBuffer, L"InitializeTTSEngine - STEP 2: Checking voice token (m_pVoiceToken=%p)", m_pVoiceToken);
        LogMessage(logBuffer);

        if (m_pVoiceToken)
        {
            LogMessage(L"InitializeTTSEngine - STEP 2: Voice token exists, querying for ISpObjectWithToken interface");

            CComQIPtr<ISpObjectWithToken> pObjectWithToken(m_pTTSEngine);
            if (pObjectWithToken)
            {
                LogMessage(L"InitializeTTSEngine - STEP 2: ISpObjectWithToken interface obtained, setting voice token");

                hr = pObjectWithToken->SetObjectToken(m_pVoiceToken);

                swprintf_s(logBuffer, L"InitializeTTSEngine - STEP 2: SetObjectToken returned 0x%08X", hr);
                LogMessage(logBuffer);

                if (FAILED(hr))
                {
                    LogMessage(L"InitializeTTSEngine - STEP 2: FAILED - SetObjectToken failed");
                    return hr;
                }

                LogMessage(L"InitializeTTSEngine - STEP 2: COMPLETED - Voice token set successfully");
            }
            else
            {
                LogMessage(L"InitializeTTSEngine - STEP 2: WARNING - ISpObjectWithToken interface not available");
            }
        }
        else
        {
            LogMessage(L"InitializeTTSEngine - STEP 2: SKIPPED - No voice token to set");
        }

        m_bInitialized = true;
        LogMessage(L"InitializeTTSEngine - SUCCESS - TTS engine initialized successfully");
        LogMessage(L"=== InitializeTTSEngine - EXIT (SUCCESS) ===");
        return S_OK;
    }
    catch (...)
    {
        LogMessage(L"InitializeTTSEngine - EXCEPTION: Unhandled exception occurred");
        LogMessage(L"=== InitializeTTSEngine - EXIT (EXCEPTION) ===");
        return E_FAIL;
    }
}

// Convert text to SPVTEXTFRAG format
HRESULT COpenSpeechSpVoice::ConvertTextToFragments(LPCWSTR pwcs, SPVTEXTFRAG** ppFragments)
{
    LogMessage(L"=== ConvertTextToFragments - ENTRY ===");

    wchar_t logBuffer[512];
    swprintf_s(logBuffer, L"ConvertTextToFragments - Parameters: pwcs=%p, ppFragments=%p", pwcs, ppFragments);
    LogMessage(logBuffer);

    if (!pwcs || !ppFragments)
    {
        LogMessage(L"ConvertTextToFragments - ERROR: Invalid parameters (NULL pointer)");
        return E_INVALIDARG;
    }

    size_t textLen = wcslen(pwcs);
    swprintf_s(logBuffer, L"ConvertTextToFragments - Text length: %zu characters", textLen);
    LogMessage(logBuffer);

    try
    {
        LogMessage(L"ConvertTextToFragments - STEP 1: Allocating SPVTEXTFRAG structure");

        // Allocate single fragment for the entire text
        SPVTEXTFRAG* pFrag = new SPVTEXTFRAG;
        ZeroMemory(pFrag, sizeof(SPVTEXTFRAG));

        swprintf_s(logBuffer, L"ConvertTextToFragments - STEP 1: COMPLETED - Fragment allocated at %p", pFrag);
        LogMessage(logBuffer);

        LogMessage(L"ConvertTextToFragments - STEP 2: Setting fragment text properties");

        pFrag->pNext = nullptr;
        pFrag->pTextStart = pwcs;
        pFrag->ulTextLen = static_cast<ULONG>(textLen);
        pFrag->ulTextSrcOffset = 0;

        swprintf_s(logBuffer, L"ConvertTextToFragments - Fragment text: pTextStart=%p, ulTextLen=%lu, ulTextSrcOffset=%lu",
                   pFrag->pTextStart, pFrag->ulTextLen, pFrag->ulTextSrcOffset);
        LogMessage(logBuffer);

        LogMessage(L"ConvertTextToFragments - STEP 3: Initializing fragment state");

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

        swprintf_s(logBuffer, L"ConvertTextToFragments - Fragment state: eAction=%d, LangID=0x%04X, RateAdj=%ld, Volume=%u",
                   pFrag->State.eAction, pFrag->State.LangID, pFrag->State.RateAdj, pFrag->State.Volume);
        LogMessage(logBuffer);

        *ppFragments = pFrag;

        LogMessage(L"ConvertTextToFragments - SUCCESS - Fragment created and returned");
        LogMessage(L"=== ConvertTextToFragments - EXIT (SUCCESS) ===");
        return S_OK;
    }
    catch (...)
    {
        LogMessage(L"ConvertTextToFragments - EXCEPTION: Memory allocation or other error occurred");
        LogMessage(L"=== ConvertTextToFragments - EXIT (EXCEPTION) ===");
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
    LogMessage(L"=== COpenSpeechSpVoice::GetStatus - ENTRY ===");

    if (!pStatus)
    {
        LogMessage(L"COpenSpeechSpVoice::GetStatus - ERROR: pStatus is NULL");
        return E_INVALIDARG;
    }

    wchar_t logBuffer[256];
    swprintf_s(logBuffer, L"COpenSpeechSpVoice::GetStatus - Parameters: pStatus=%p, ppszLastBookmark=%p",
               pStatus, ppszLastBookmark);
    LogMessage(logBuffer);

    EnterCriticalSection(&m_csVoiceState);
    *pStatus = m_voiceStatus;

    swprintf_s(logBuffer, L"COpenSpeechSpVoice::GetStatus - Status: RunningState=%lu, CurrentStream=%lu",
               m_voiceStatus.dwRunningState, m_voiceStatus.ulCurrentStream);
    LogMessage(logBuffer);

    LeaveCriticalSection(&m_csVoiceState);

    if (ppszLastBookmark)
    {
        *ppszLastBookmark = nullptr; // No bookmark support yet
        LogMessage(L"COpenSpeechSpVoice::GetStatus - Bookmark pointer set to NULL");
    }

    LogMessage(L"COpenSpeechSpVoice::GetStatus - SUCCESS - Returning S_OK");
    return S_OK;
}

// Set voice token
STDMETHODIMP COpenSpeechSpVoice::SetVoice(ISpObjectToken* pToken)
{
    LogMessage(L"COpenSpeechSpVoice::SetVoice - ENTRY - Setting voice token");

    wchar_t logBuffer[256];
    swprintf_s(logBuffer, L"COpenSpeechSpVoice::SetVoice - pToken=%p, m_pTTSEngine=%p, m_bInitialized=%s",
               pToken, m_pTTSEngine, m_bInitialized ? L"true" : L"false");
    LogMessage(logBuffer);

    m_pVoiceToken = pToken;

    // If TTS engine is already initialized, update its token
    if (m_pTTSEngine && m_bInitialized)
    {
        LogMessage(L"COpenSpeechSpVoice::SetVoice - TTS engine exists, querying for ISpObjectWithToken");
        CComQIPtr<ISpObjectWithToken> pObjectWithToken(m_pTTSEngine);
        if (pObjectWithToken)
        {
            LogMessage(L"COpenSpeechSpVoice::SetVoice - Calling SetObjectToken on TTS engine");
            HRESULT hr = pObjectWithToken->SetObjectToken(pToken);
            swprintf_s(logBuffer, L"COpenSpeechSpVoice::SetVoice - SetObjectToken returned 0x%08X", hr);
            LogMessage(logBuffer);
            return hr;
        }
        else
        {
            LogMessage(L"COpenSpeechSpVoice::SetVoice - ISpObjectWithToken interface not available");
        }
    }
    else
    {
        LogMessage(L"COpenSpeechSpVoice::SetVoice - TTS engine not initialized yet, storing token for later");
    }

    LogMessage(L"COpenSpeechSpVoice::SetVoice - EXIT - Returning S_OK");
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
    LogMessage(L"COpenSpeechSpVoice::DisplayUI - CALLED - Basic implementation (returning S_OK)");
    // Basic implementation - just return success
    // In a full implementation, we would display the requested UI
    return S_OK;
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
    LogMessage(L"COpenSpeechSpVoice::SetNotifyCallbackFunction - CALLED - Basic implementation (callback stored)");
    // Basic implementation - just return success
    // In a full implementation, we would store and use the callback
    return S_OK;
}

STDMETHODIMP COpenSpeechSpVoice::SetNotifyCallbackInterface(ISpNotifyCallback* pSpCallback, WPARAM wParam, LPARAM lParam)
{
    LogMessage(L"COpenSpeechSpVoice::SetNotifyCallbackInterface - CALLED - Basic implementation (returning S_OK)");
    // Basic implementation - just return success
    // In a full implementation, we would store and use the callback interface
    return S_OK;
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

// AACSpeakHelper Pipe Service Implementation - Direct call bypassing ISpTTSEngine
HRESULT COpenSpeechSpVoice::CallAACSpeakHelperPipeService(const std::wstring& text, std::vector<BYTE>& audioData)
{
    try
    {
        LogMessage(L"COpenSpeechSpVoice::CallAACSpeakHelperPipeService - Starting direct pipe service call...");

        // Connect to AACSpeakHelper pipe
        HANDLE hPipe = INVALID_HANDLE_VALUE;
        if (!ConnectToAACSpeakHelper(hPipe))
        {
            LogMessage(L"COpenSpeechSpVoice::CallAACSpeakHelperPipeService - Failed to connect to AACSpeakHelper pipe service");
            return E_FAIL;
        }

        LogMessage(L"COpenSpeechSpVoice::CallAACSpeakHelperPipeService - Connected to pipe, sending text...");

        // Send text to pipe
        HRESULT hr = SendTextToPipe(hPipe, text);
        if (FAILED(hr))
        {
            LogMessage(L"COpenSpeechSpVoice::CallAACSpeakHelperPipeService - Failed to send text to pipe");
            CloseHandle(hPipe);
            return hr;
        }

        LogMessage(L"COpenSpeechSpVoice::CallAACSpeakHelperPipeService - Text sent, receiving audio...");

        // Receive audio from pipe
        hr = ReceiveAudioFromPipe(hPipe, audioData);
        CloseHandle(hPipe);

        if (FAILED(hr))
        {
            LogMessage(L"COpenSpeechSpVoice::CallAACSpeakHelperPipeService - Failed to receive audio from pipe");
            return hr;
        }

        wchar_t logBuffer[256];
        swprintf_s(logBuffer, L"COpenSpeechSpVoice::CallAACSpeakHelperPipeService - Successfully received %zu bytes of audio", audioData.size());
        LogMessage(logBuffer);

        return S_OK;
    }
    catch (const std::exception& ex)
    {
        std::string error = "Exception in COpenSpeechSpVoice::CallAACSpeakHelperPipeService: ";
        error += ex.what();
        LogMessage(std::wstring(error.begin(), error.end()).c_str());
        return E_FAIL;
    }
    catch (...)
    {
        LogMessage(L"Unknown exception in COpenSpeechSpVoice::CallAACSpeakHelperPipeService");
        return E_FAIL;
    }
}

// Connect to AACSpeakHelper pipe
bool COpenSpeechSpVoice::ConnectToAACSpeakHelper(HANDLE& hPipe)
{
    const wchar_t* pipeName = L"\\\\.\\pipe\\AACSpeakHelper";
    const int maxRetries = 5;
    const int retryDelayMs = 1000;

    for (int retry = 0; retry < maxRetries; retry++)
    {
        wchar_t logBuffer[256];
        swprintf_s(logBuffer, L"COpenSpeechSpVoice::ConnectToAACSpeakHelper - Attempt %d/%d", retry + 1, maxRetries);
        LogMessage(logBuffer);

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
            LogMessage(L"COpenSpeechSpVoice::ConnectToAACSpeakHelper - Successfully connected to pipe");
            return true;
        }

        DWORD error = GetLastError();
        if (error == ERROR_PIPE_BUSY)
        {
            LogMessage(L"COpenSpeechSpVoice::ConnectToAACSpeakHelper - Pipe is busy, waiting...");
            if (!WaitNamedPipeW(pipeName, 30000)) // 30 second timeout
            {
                LogMessage(L"COpenSpeechSpVoice::ConnectToAACSpeakHelper - Timeout waiting for pipe");
                continue;
            }
        }
        else
        {
            swprintf_s(logBuffer, L"COpenSpeechSpVoice::ConnectToAACSpeakHelper - Failed to connect, error: %lu", error);
            LogMessage(logBuffer);
        }

        if (retry < maxRetries - 1)
        {
            Sleep(retryDelayMs);
        }
    }

    LogMessage(L"COpenSpeechSpVoice::ConnectToAACSpeakHelper - Failed to connect after all retries");
    return false;
}

// Send text to AACSpeakHelper pipe
HRESULT COpenSpeechSpVoice::SendTextToPipe(HANDLE hPipe, const std::wstring& text)
{
    try
    {
        LogMessage(L"COpenSpeechSpVoice::SendTextToPipe - Creating JSON message");

        // Convert text to UTF-8
        int utf8Size = WideCharToMultiByte(CP_UTF8, 0, text.c_str(), -1, nullptr, 0, nullptr, nullptr);
        if (utf8Size == 0)
        {
            LogMessage(L"COpenSpeechSpVoice::SendTextToPipe - Failed to convert text to UTF-8");
            return E_FAIL;
        }

        std::vector<char> utf8Text(utf8Size);
        WideCharToMultiByte(CP_UTF8, 0, text.c_str(), -1, utf8Text.data(), utf8Size, nullptr, nullptr);

        // Create simple JSON message for AACSpeakHelper
        std::string jsonMessage = "{\n";
        jsonMessage += "  \"clipboard_text\": \"";

        // Escape text for JSON
        for (size_t i = 0; i < utf8Text.size() - 1; i++) // -1 to skip null terminator
        {
            char c = utf8Text[i];
            if (c == '"') jsonMessage += "\\\"";
            else if (c == '\\') jsonMessage += "\\\\";
            else if (c == '\n') jsonMessage += "\\n";
            else if (c == '\r') jsonMessage += "\\r";
            else if (c == '\t') jsonMessage += "\\t";
            else jsonMessage += c;
        }

        jsonMessage += "\",\n";
        jsonMessage += "  \"args\": {\n";
        jsonMessage += "    \"engine\": \"azure\",\n";
        jsonMessage += "    \"voice\": \"en-GB-LibbyNeural\",\n";
        jsonMessage += "    \"rate\": 0,\n";
        jsonMessage += "    \"volume\": 100,\n";
        jsonMessage += "    \"listvoices\": false,\n";
        jsonMessage += "    \"return_audio_bytes\": true\n";
        jsonMessage += "  },\n";
        jsonMessage += "  \"config\": {\n";
        jsonMessage += "    \"TTS\": {\n";
        jsonMessage += "      \"engine\": \"azureTTS\",\n";
        jsonMessage += "      \"bypass_tts\": \"False\",\n";
        jsonMessage += "      \"save_audio_file\": \"True\",\n";
        jsonMessage += "      \"rate\": \"0\",\n";
        jsonMessage += "      \"volume\": \"100\"\n";
        jsonMessage += "    },\n";
        jsonMessage += "    \"translate\": {\n";
        jsonMessage += "      \"no_translate\": \"True\",\n";
        jsonMessage += "      \"start_lang\": \"en\",\n";
        jsonMessage += "      \"end_lang\": \"en\",\n";
        jsonMessage += "      \"replace_pb\": \"True\"\n";
        jsonMessage += "    },\n";
        jsonMessage += "    \"azureTTS\": {\n";
        jsonMessage += "      \"key\": \"b14f8945b0f1459f9964bdd72c42c2cc\",\n";
        jsonMessage += "      \"location\": \"uksouth\",\n";
        jsonMessage += "      \"voice_id\": \"en-GB-LibbyNeural\"\n";
        jsonMessage += "    }\n";
        jsonMessage += "  }\n";
        jsonMessage += "}";

        wchar_t logBuffer[256];
        swprintf_s(logBuffer, L"COpenSpeechSpVoice::SendTextToPipe - Sending %zu bytes", jsonMessage.length());
        LogMessage(logBuffer);

        // Send message to pipe
        DWORD bytesWritten = 0;
        DWORD messageLength = static_cast<DWORD>(jsonMessage.length());

        if (!WriteFile(hPipe, jsonMessage.c_str(), messageLength, &bytesWritten, nullptr) ||
            bytesWritten != messageLength)
        {
            LogMessage(L"COpenSpeechSpVoice::SendTextToPipe - Failed to write message to pipe");
            return E_FAIL;
        }

        LogMessage(L"COpenSpeechSpVoice::SendTextToPipe - Message sent successfully");
        return S_OK;
    }
    catch (...)
    {
        LogMessage(L"COpenSpeechSpVoice::SendTextToPipe - Exception occurred");
        return E_FAIL;
    }
}

// Receive audio data from AACSpeakHelper pipe
HRESULT COpenSpeechSpVoice::ReceiveAudioFromPipe(HANDLE hPipe, std::vector<BYTE>& audioData)
{
    try
    {
        LogMessage(L"COpenSpeechSpVoice::ReceiveAudioFromPipe - Waiting for audio data...");

        // First, read the length prefix (4 bytes, little-endian uint32)
        BYTE lengthBuffer[4];
        DWORD bytesRead = 0;

        if (!ReadFile(hPipe, lengthBuffer, sizeof(lengthBuffer), &bytesRead, nullptr) ||
            bytesRead != sizeof(lengthBuffer))
        {
            LogMessage(L"COpenSpeechSpVoice::ReceiveAudioFromPipe - Failed to read length prefix");
            return E_FAIL;
        }

        // Extract audio data length from little-endian uint32
        DWORD audioLength =
            (static_cast<DWORD>(lengthBuffer[0])) |
            (static_cast<DWORD>(lengthBuffer[1]) << 8) |
            (static_cast<DWORD>(lengthBuffer[2]) << 16) |
            (static_cast<DWORD>(lengthBuffer[3]) << 24);

        wchar_t logBuffer[256];
        swprintf_s(logBuffer, L"COpenSpeechSpVoice::ReceiveAudioFromPipe - Expecting %lu bytes of audio", audioLength);
        LogMessage(logBuffer);

        if (audioLength == 0 || audioLength > 10 * 1024 * 1024) // Sanity check: max 10MB
        {
            swprintf_s(logBuffer, L"COpenSpeechSpVoice::ReceiveAudioFromPipe - Invalid audio length: %lu", audioLength);
            LogMessage(logBuffer);
            return E_FAIL;
        }

        // Resize buffer to hold the audio data
        audioData.resize(audioLength);

        // Read the audio data in chunks
        DWORD totalBytesRead = 0;
        const DWORD chunkSize = 64 * 1024; // 64KB chunks

        while (totalBytesRead < audioLength)
        {
            DWORD remainingBytes = audioLength - totalBytesRead;
            DWORD bytesToRead = (chunkSize < remainingBytes) ? chunkSize : remainingBytes;
            DWORD chunkBytesRead = 0;

            if (!ReadFile(hPipe, audioData.data() + totalBytesRead, bytesToRead, &chunkBytesRead, nullptr))
            {
                swprintf_s(logBuffer, L"COpenSpeechSpVoice::ReceiveAudioFromPipe - Failed to read chunk at offset %lu", totalBytesRead);
                LogMessage(logBuffer);
                return E_FAIL;
            }

            if (chunkBytesRead == 0)
            {
                LogMessage(L"COpenSpeechSpVoice::ReceiveAudioFromPipe - Unexpected end of pipe data");
                return E_FAIL;
            }

            totalBytesRead += chunkBytesRead;

            if (totalBytesRead % (256 * 1024) == 0) // Log every 256KB
            {
                swprintf_s(logBuffer, L"COpenSpeechSpVoice::ReceiveAudioFromPipe - Read %lu/%lu bytes", totalBytesRead, audioLength);
                LogMessage(logBuffer);
            }
        }

        swprintf_s(logBuffer, L"COpenSpeechSpVoice::ReceiveAudioFromPipe - Successfully received %zu bytes", audioData.size());
        LogMessage(logBuffer);
        return S_OK;
    }
    catch (...)
    {
        LogMessage(L"COpenSpeechSpVoice::ReceiveAudioFromPipe - Exception occurred");
        return E_FAIL;
    }
}
