@echo off
REM Register C++ COM Wrapper for SAPI
REM This must be run as Administrator

echo === Register C++ COM Wrapper ===

REM Check if DLL exists
if not exist "NativeTTSWrapper\x64\Release\NativeTTSWrapper.dll" (
    echo ❌ COM wrapper DLL not found
    echo    Please build the C++ project first:
    echo    msbuild "NativeTTSWrapper\NativeTTSWrapper.vcxproj" /p:Configuration=Release /p:Platform=x64
    pause
    exit /b 1
)

echo ✅ Found COM wrapper DLL

REM Copy required DLLs to the same directory
echo 📋 Copying required dependencies...
copy "NativeTTSWrapper\libs\*.dll" "NativeTTSWrapper\x64\Release\" >nul 2>&1
copy "NativeTTSWrapper\azure-speech-sdk\bin\*.dll" "NativeTTSWrapper\x64\Release\" >nul 2>&1

REM Register the COM wrapper
echo 🚀 Registering COM wrapper...
regsvr32 /s "NativeTTSWrapper\x64\Release\NativeTTSWrapper.dll"

if %ERRORLEVEL% EQU 0 (
    echo ✅ COM wrapper registered successfully!
    echo    SAPI voices should now work for speech synthesis.
) else (
    echo ❌ COM wrapper registration failed with exit code: %ERRORLEVEL%
    echo    Make sure you're running as Administrator.
)

echo.
echo Press any key to continue...
pause >nul
