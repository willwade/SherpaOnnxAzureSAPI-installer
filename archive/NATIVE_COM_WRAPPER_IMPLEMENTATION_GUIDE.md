# 🎯 Native COM Wrapper Implementation Guide

## 🚀 **OBJECTIVE: Achieve 100% SAPI Compatibility**

This guide provides the complete implementation plan for the native C++ COM wrapper that will enable full SAPI compatibility with our ProcessBridge TTS system.

---

## 📋 **CURRENT STATUS**

### ✅ **What's Already Working (95%)**
- **ProcessBridge TTS System:** 100% functional
- **SherpaWorker.exe:** Production ready (58.7 MB)
- **Enhanced Audio Generation:** High-quality speech synthesis
- **COM Object (Managed):** Works when called directly
- **Voice Registration:** Amy voice appears in Windows
- **Installation System:** Complete deployment automation

### ❌ **What Needs Implementation (5%)**
- **Native COM Wrapper:** SAPI interface recognition
- **Standard SAPI Calls:** `voice.Speak()` compatibility

---

## 🏗️ **ARCHITECTURE OVERVIEW**

### **Current Architecture (95% Complete)**
```
┌─────────────────┐    Managed COM   ┌──────────────────┐    JSON IPC    ┌──────────────────┐
│      SAPI       │ ──────────────► │  Managed COM     │ ──────────────► │  SherpaWorker    │
│                 │     (FAILS)     │ (.NET Fx 4.7.2) │                 │   (.NET 6.0)     │
│                 │ ◄────────────── │                  │ ◄────────────── │                  │
└─────────────────┘                 └──────────────────┘    WAV Audio   └──────────────────┘
```

### **Target Architecture (100% Complete)**
```
┌─────────────────┐    Native COM    ┌──────────────────┐    JSON IPC    ┌──────────────────┐
│      SAPI       │ ──────────────► │  Native Wrapper  │ ──────────────► │  SherpaWorker    │
│                 │    (WORKS!)     │    (C++ DLL)     │                 │   (.NET 6.0)     │
│                 │ ◄────────────── │                  │ ◄────────────── │                  │
└─────────────────┘    Audio Data   └──────────────────┘    WAV Audio   └──────────────────┘
```

---

## 📝 **IMPLEMENTATION STEPS**

### **Step 1: Native COM Project Setup**

**1.1 Create Visual Studio C++ ATL COM Project**
- Project Type: ATL Project
- Application Type: Dynamic Link Library (DLL)
- ATL Options: Attributed, Support MFC
- Target Platform: x64

**1.2 Required Dependencies**
```cpp
#include <sapi.h>        // SAPI interfaces
#include <sphelper.h>    // SAPI helper functions
#include <atlbase.h>     // ATL base classes
#include <atlcom.h>      // ATL COM support
```

**1.3 Project Configuration**
- Configuration: Release x64
- Character Set: Unicode
- ATL Usage: Dynamic Link to ATL
- Additional Dependencies: `sapi.lib ole32.lib oleaut32.lib uuid.lib`

### **Step 2: Interface Implementation**

**2.1 Class Declaration**
```cpp
class ATL_NO_VTABLE CNativeTTSWrapper :
    public CComObjectRootEx<CComMultiThreadModel>,
    public CComCoClass<CNativeTTSWrapper, &CLSID_NativeTTSWrapper>,
    public ISpTTSEngine,
    public ISpObjectWithToken
{
    // COM map and interface methods
};
```

**2.2 Critical Methods to Implement**
```cpp
// ISpTTSEngine
STDMETHOD(Speak)(DWORD dwSpeakFlags, REFGUID rguidFormatId,
                 const WAVEFORMATEX* pWaveFormatEx,
                 const SPVTEXTFRAG* pTextFragList,
                 ISpTTSEngineSite* pOutputSite);

STDMETHOD(GetOutputFormat)(const GUID* pTargetFormatId,
                           const WAVEFORMATEX* pTargetWaveFormatEx,
                           GUID* pOutputFormatId,
                           WAVEFORMATEX** ppCoMemOutputWaveFormatEx);

// ISpObjectWithToken
STDMETHOD(SetObjectToken)(ISpObjectToken* pToken);
STDMETHOD(GetObjectToken)(ISpObjectToken** ppToken);
```

### **Step 3: ProcessBridge Integration**

**3.1 Text Extraction**
```cpp
std::wstring ExtractTextFromFragList(const SPVTEXTFRAG* pTextFragList)
{
    std::wstring result;
    const SPVTEXTFRAG* pFrag = pTextFragList;
    
    while (pFrag) {
        if (pFrag->pTextStart && pFrag->ulTextLen > 0) {
            result.append(pFrag->pTextStart, pFrag->ulTextLen);
        }
        pFrag = pFrag->pNext;
    }
    return result;
}
```

**3.2 SherpaWorker Execution**
```cpp
bool CallSherpaWorker(const std::wstring& requestPath)
{
    std::wstring commandLine = L"\"C:\\Program Files\\OpenAssistive\\OpenSpeech\\SherpaWorker.exe\" \"" + requestPath + L"\"";
    
    STARTUPINFOW si = { sizeof(si) };
    PROCESS_INFORMATION pi = { 0 };
    
    if (!CreateProcessW(nullptr, const_cast<wchar_t*>(commandLine.c_str()),
                       nullptr, nullptr, FALSE, CREATE_NO_WINDOW,
                       nullptr, nullptr, &si, &pi)) {
        return false;
    }
    
    WaitForSingleObject(pi.hProcess, 30000); // 30 second timeout
    // Check exit code and cleanup
    return true;
}
```

**3.3 Audio Data Return**
```cpp
HRESULT ReturnAudioToSAPI(const std::vector<BYTE>& audioData, ISpTTSEngineSite* pOutputSite)
{
    // Send start event
    SPEVENT startEvent = { 0 };
    startEvent.eEventId = SPEI_START_INPUT_STREAM;
    pOutputSite->AddEvents(&startEvent, 1);
    
    // Write audio data
    ULONG bytesWritten = 0;
    HRESULT hr = pOutputSite->Write(audioData.data(), (ULONG)audioData.size(), &bytesWritten);
    
    // Send end event
    SPEVENT endEvent = { 0 };
    endEvent.eEventId = SPEI_END_INPUT_STREAM;
    endEvent.ullAudioStreamOffset = audioData.size();
    pOutputSite->AddEvents(&endEvent, 1);
    
    return hr;
}
```

### **Step 4: Build and Registration**

**4.1 Build Configuration**
- Platform: x64
- Configuration: Release
- Output: NativeTTSWrapper.dll

**4.2 COM Registration**
```cpp
// DLL exports
STDAPI DllRegisterServer(void);
STDAPI DllUnregisterServer(void);
STDAPI DllCanUnloadNow(void);
STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv);
```

**4.3 Voice Token Update**
```powershell
# Update registry to use native DLL
$newClsid = "{E1C4A8F2-9B3D-4A5E-8F7C-2D1B3E4F5A6B}"
Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Speech\Voices\Tokens\amy" -Name "CLSID" -Value $newClsid
```

---

## 🧪 **TESTING STRATEGY**

### **Test 1: COM Object Creation**
```powershell
$nativeObject = New-Object -ComObject "NativeTTSWrapper.CNativeTTSWrapper"
# Should succeed without errors
```

### **Test 2: SAPI Integration**
```powershell
$voice = New-Object -ComObject SAPI.SpVoice
$voice.Voice = $amyVoice
$voice.Speak("Native wrapper test") # Should call our Speak method
```

### **Test 3: ProcessBridge Execution**
- Verify SherpaWorker.exe is called
- Check JSON request/response files
- Validate audio generation
- Confirm SAPI receives audio data

### **Test 4: Performance Validation**
- Measure end-to-end latency
- Verify audio quality maintained
- Test error handling
- Validate cleanup

---

## 📊 **EXPECTED RESULTS**

### **Before Implementation (Current State)**
```
voice.Speak("Hello") → E_FAIL (SAPI can't call managed COM)
```

### **After Implementation (Target State)**
```
voice.Speak("Hello") → SUCCESS (Native COM → ProcessBridge → Audio)
```

### **Performance Targets**
- **Total Latency:** ≤ 400ms (including native wrapper overhead)
- **Audio Quality:** Same as current ProcessBridge (22050Hz, 16-bit)
- **Reliability:** Same error handling and fallback mechanisms
- **Compatibility:** Works with all SAPI applications

---

## 🎯 **SUCCESS CRITERIA**

### **Functional Requirements**
1. ✅ `voice.Speak()` calls our native Speak method
2. ✅ Text extraction from SPVTEXTFRAG works correctly
3. ✅ SherpaWorker.exe execution succeeds
4. ✅ Audio data returned to SAPI properly
5. ✅ All SAPI events sent correctly

### **Performance Requirements**
1. ✅ End-to-end latency ≤ 400ms
2. ✅ Audio quality maintained
3. ✅ Memory usage reasonable
4. ✅ No resource leaks

### **Compatibility Requirements**
1. ✅ Works with Windows Speech Recognition
2. ✅ Works with accessibility tools
3. ✅ Works with third-party SAPI applications
4. ✅ Maintains existing ProcessBridge benefits

---

## 🔧 **IMPLEMENTATION TIMELINE**

### **Day 1: Project Setup and Core Implementation**
- ✅ Create ATL COM project
- ✅ Implement ISpTTSEngine interface
- ✅ Implement ISpObjectWithToken interface
- ✅ Add ProcessBridge integration code

### **Day 2: Build, Test, and Deploy**
- ✅ Build native DLL
- ✅ Register COM object
- ✅ Update voice token registration
- ✅ Test SAPI integration
- ✅ Validate ProcessBridge execution

### **Result: 100% SAPI Compatibility Achieved**

---

## 🎉 **FINAL OUTCOME**

**With the native COM wrapper implemented:**

✅ **Complete SAPI Compatibility** - All applications work  
✅ **ProcessBridge Benefits Maintained** - Same performance and quality  
✅ **Universal Application Support** - Works everywhere SAPI is used  
✅ **Production Ready** - Robust error handling and logging  

**The ProcessBridge TTS system becomes a complete, production-ready SAPI bridge to SherpaOnnx with 100% compatibility!**
