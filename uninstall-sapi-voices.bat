@echo off
setlocal enabledelayedexpansion

REM ============================================================================
REM SherpaOnnx Azure SAPI Installer - Manual Uninstaller
REM ============================================================================
REM This script manually removes SAPI voices and unregisters COM components
REM Use this if the main uninstaller fails or for development cleanup
REM ============================================================================

echo.
echo ============================================================================
echo  SherpaOnnx Azure SAPI Installer - Manual Uninstaller
echo ============================================================================
echo.

REM Check if running as administrator
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo ❌ ERROR: This script must be run as Administrator
    echo    Right-click and select "Run as administrator"
    pause
    exit /b 1
)

echo ✅ Running as Administrator
echo.

echo ⚠️  WARNING: This will remove all SherpaOnnx and Azure SAPI voices
echo    and unregister the COM wrapper components.
echo.
set /p confirm="Are you sure you want to continue? (y/N): "
if /i not "%confirm%"=="y" (
    echo Operation cancelled.
    pause
    exit /b 0
)

echo.
echo 🗑️  Starting uninstallation process...
echo ============================================================================

REM Step 1: Stop any running AACSpeakHelper processes
echo.
echo 📝 Step 1: Stopping AACSpeakHelper processes...
echo ----------------------------------------
taskkill /f /im AACSpeakHelperServer.exe >nul 2>&1
taskkill /f /im AACSpeakHelper.exe >nul 2>&1
echo ✅ Stopped AACSpeakHelper processes

REM Step 2: Unregister COM wrapper
echo.
echo 📝 Step 2: Unregistering COM wrapper...
echo ----------------------------------------

if exist "NativeTTSWrapper\x64\Release\NativeTTSWrapper.dll" (
    echo Unregistering NativeTTSWrapper.dll...
    regsvr32 /u /s "NativeTTSWrapper\x64\Release\NativeTTSWrapper.dll"
    echo ✅ Unregistered local NativeTTSWrapper.dll
)

REM Try common installation locations
if exist "C:\Program Files\SherpaOnnx Azure SAPI Bridge\NativeTTSWrapper.dll" (
    echo Unregistering installed NativeTTSWrapper.dll...
    regsvr32 /u /s "C:\Program Files\SherpaOnnx Azure SAPI Bridge\NativeTTSWrapper.dll"
    echo ✅ Unregistered installed NativeTTSWrapper.dll
)

if exist "C:\Program Files (x86)\SherpaOnnx Azure SAPI Bridge\NativeTTSWrapper.dll" (
    echo Unregistering installed NativeTTSWrapper.dll (x86)...
    regsvr32 /u /s "C:\Program Files (x86)\SherpaOnnx Azure SAPI Bridge\NativeTTSWrapper.dll"
    echo ✅ Unregistered installed NativeTTSWrapper.dll (x86)
)

REM Step 3: Remove SAPI voice registrations
echo.
echo 📝 Step 3: Removing SAPI voice registrations...
echo ----------------------------------------

echo Removing voice registry entries...

REM Remove our CLSID entries
reg delete "HKEY_LOCAL_MACHINE\SOFTWARE\Classes\CLSID\{4A8B9C2D-1E3F-4567-8901-234567890ABC}" /f >nul 2>&1
reg delete "HKEY_LOCAL_MACHINE\SOFTWARE\Classes\CLSID\{4A8B9C2D-1E3F-4567-8901-234567890ABD}" /f >nul 2>&1

REM Remove voice tokens (common patterns)
for %%i in (
    "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices\Tokens\British-English-Azure-Libby"
    "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices\Tokens\British-English-sherpaonnx-001"
    "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices\Tokens\AACSpeakHelper*"
) do (
    reg delete "%%i" /f >nul 2>&1
)

echo ✅ Removed SAPI voice registrations

REM Step 4: Clean up files
echo.
echo 📝 Step 4: Cleaning up files...
echo ----------------------------------------

REM Remove log files
if exist "C:\OpenSpeech\native_tts_debug.log" (
    del "C:\OpenSpeech\native_tts_debug.log" >nul 2>&1
    echo ✅ Removed debug log file
)

REM Remove temporary files
if exist "AACSpeakHelper\*.log" (
    del "AACSpeakHelper\*.log" >nul 2>&1
    echo ✅ Removed AACSpeakHelper log files
)

if exist "dist\" (
    rmdir /s /q "dist\" >nul 2>&1
    echo ✅ Removed dist directory
)

if exist "build\" (
    rmdir /s /q "build\" >nul 2>&1
    echo ✅ Removed build directory
)

REM Step 5: Registry cleanup
echo.
echo 📝 Step 5: Additional registry cleanup...
echo ----------------------------------------

REM Remove any remaining entries
reg delete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\sapi_voice_installer.exe" /f >nul 2>&1
reg delete "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\SherpaOnnx Azure SAPI Bridge" /f >nul 2>&1

echo ✅ Registry cleanup completed

echo.
echo 🎉 UNINSTALLATION COMPLETED!
echo ============================================================================
echo.
echo The following components have been removed:
echo   ✅ COM wrapper unregistered
echo   ✅ SAPI voice registrations removed
echo   ✅ Registry entries cleaned
echo   ✅ Temporary files removed
echo.
echo Note: You may need to restart applications that use SAPI voices
echo       for the changes to take effect.
echo.
echo To reinstall, run: build-all.bat
echo.
pause
