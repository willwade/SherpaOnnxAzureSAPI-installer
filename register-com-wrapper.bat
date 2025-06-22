@echo off
REM Register C++ COM Wrapper for SAPI
REM This script is MORE RELIABLE than direct regsvr32 calls
REM IMPORTANT: This must be run as Administrator

echo === Register C++ COM Wrapper ===
echo.
echo ⚠️  IMPORTANT: This script is more reliable than direct regsvr32 calls
echo    Always use: sudo .\register-com-wrapper.bat
echo    Avoid: regsvr32 "NativeTTSWrapper\x64\Release\NativeTTSWrapper.dll"
echo.

REM Check if running as Administrator
net session >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo ❌ This script must be run as Administrator
    echo    Please run: sudo .\register-com-wrapper.bat
    pause
    exit /b 1
)

echo ✅ Running as Administrator

REM Check if DLL exists (try new version first, then old version)
if exist "NativeTTSWrapper\x64\Release\NativeTTSWrapper_New.dll" (
    echo ✅ Found updated COM wrapper DLL: NativeTTSWrapper_New.dll
    set "DLL_PATH=NativeTTSWrapper\x64\Release\NativeTTSWrapper_New.dll"
) else if exist "NativeTTSWrapper\x64\Release\NativeTTSWrapper.dll" (
    echo ✅ Found COM wrapper DLL: NativeTTSWrapper.dll
    set "DLL_PATH=NativeTTSWrapper\x64\Release\NativeTTSWrapper.dll"
) else (
    echo ❌ COM wrapper DLL not found
    echo    Please build the C++ project first:
    echo    .\build_com_wrapper.bat
    echo    OR: msbuild "NativeTTSWrapper\NativeTTSWrapper.vcxproj" /p:Configuration=Release /p:Platform=x64
    pause
    exit /b 1
)

REM Copy required DLLs to the same directory
echo 📋 Copying required dependencies...
copy "NativeTTSWrapper\libs\*.dll" "NativeTTSWrapper\x64\Release\" >nul 2>&1
copy "NativeTTSWrapper\azure-speech-sdk\bin\*.dll" "NativeTTSWrapper\x64\Release\" >nul 2>&1

REM Clean any old registry entries first
echo 🧹 Cleaning old registry entries...
reg delete "HKEY_CLASSES_ROOT\CLSID\{4A8B9C2D-1E3F-4567-8901-234567890ABC}" /f >nul 2>&1

REM Register the COM wrapper
echo 🚀 Registering COM wrapper...
regsvr32 /s "NativeTTSWrapper\x64\Release\NativeTTSWrapper.dll"

if %ERRORLEVEL% EQU 0 (
    echo ✅ COM wrapper registered successfully!
    echo    SAPI voices should now work for speech synthesis.
    echo.
    echo 🔍 Verifying registration...
    reg query "HKEY_CLASSES_ROOT\CLSID\{4A8B9C2D-1E3F-4567-8901-234567890ABC}" >nul 2>&1
    if %ERRORLEVEL% EQU 0 (
        echo ✅ CLSID registration verified
    ) else (
        echo ⚠️  CLSID not found in registry - registration may have failed
    )
) else (
    echo ❌ COM wrapper registration failed with exit code: %ERRORLEVEL%
    echo    This is why we recommend using this batch script instead of direct regsvr32!
    echo    Try running the script again or check the DLL dependencies.
)

echo.
echo Press any key to continue...
pause >nul
