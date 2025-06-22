@echo off
REM Install Voice with Admin Privileges
REM Usage: install-voice.bat English-SherpaOnnx-Jenny

if "%1"=="" (
    echo Usage: install-voice.bat VoiceName
    echo Example: install-voice.bat English-SherpaOnnx-Jenny
    pause
    exit /b 1
)

set VOICE_NAME=%1

echo === SAPI Voice Installation (Admin Mode) ===
echo Voice: %VOICE_NAME%

REM Check if installer exists
if not exist "Installer\bin\Release\net6.0\win-x64\SherpaOnnxSAPIInstaller.exe" (
    echo ❌ Installer not found
    echo    Please build the project first:
    echo    dotnet build Installer/Installer.csproj -c Release
    pause
    exit /b 1
)

REM Check if voice config exists
if not exist "voice_configs\%VOICE_NAME%.json" (
    echo ❌ Voice configuration not found: voice_configs\%VOICE_NAME%.json
    echo    Available configurations:
    dir /b voice_configs\*.json
    pause
    exit /b 1
)

echo ✅ Found installer and voice config
echo.
echo 🚀 Installing voice...

REM Install the voice
"Installer\bin\Release\net6.0\win-x64\SherpaOnnxSAPIInstaller.exe" install-pipe-voice %VOICE_NAME%

if %ERRORLEVEL% EQU 0 (
    echo ✅ Voice installed successfully!
    echo    The voice should now appear in Windows SAPI applications.
) else (
    echo ❌ Installation failed with exit code: %ERRORLEVEL%
)

echo.
echo Press any key to continue...
pause >nul
