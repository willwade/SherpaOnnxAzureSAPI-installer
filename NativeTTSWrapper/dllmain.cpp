#include "stdafx.h"
#include "resource.h"
#include "NativeTTSWrapper_i.h"
#include "dllmain.h"

CNativeTTSWrapperModule _AtlModule;

// Simple debug logging function
static void DebugLog(const wchar_t* message)
{
    OutputDebugStringW(message);
    OutputDebugStringW(L"\n");
    try
    {
        // Get DLL directory
        wchar_t dllPath[MAX_PATH];
        GetModuleFileNameW((HMODULE)&_AtlModule, dllPath, MAX_PATH);
        std::wstring path(dllPath);
        size_t lastSlash = path.find_last_of(L"\\");
        if (lastSlash != std::wstring::npos)
        {
            path = path.substr(0, lastSlash);
            path += L"\\native_tts_debug.log";
            std::wofstream logFile(path, std::ios::app);
            if (logFile.is_open())
            {
                SYSTEMTIME st;
                GetLocalTime(&st);
                logFile << st.wYear << L"-" << st.wMonth << L"-" << st.wDay << L" "
                       << st.wHour << L":" << st.wMinute << L":" << st.wSecond << L"."
                       << st.wMilliseconds << L": " << message << std::endl;
            }
        }
    }
    catch (...) {}
}

// DLL Entry Point
extern "C" BOOL WINAPI DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved)
{
    if (dwReason == DLL_PROCESS_ATTACH)
    {
        DebugLog(L"*** DLL_PROCESS_ATTACH - DllMain called ***");
    }
    return _AtlModule.DllMain(dwReason, lpReserved);
}

// Used to determine whether the DLL can be unloaded by OLE.
STDAPI DllCanUnloadNow(void)
{
    return _AtlModule.DllCanUnloadNow();
}

// Returns a class factory to create an object of the requested type.
STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv)
{
    DebugLog(L"*** DllGetClassObject called ***");
    HRESULT hr = _AtlModule.DllGetClassObject(rclsid, riid, ppv);
    DebugLog((L"DllGetClassObject result: " + std::to_wstring(hr)).c_str());
    return hr;
}

// DllRegisterServer - Adds entries to the system registry.
STDAPI DllRegisterServer(void)
{
    // registers object, typelib and all interfaces in typelib
    HRESULT hr = _AtlModule.DllRegisterServer();
    return hr;
}

// DllUnregisterServer - Removes entries from the system registry.
STDAPI DllUnregisterServer(void)
{
    HRESULT hr = _AtlModule.DllUnregisterServer();
    return hr;
}
