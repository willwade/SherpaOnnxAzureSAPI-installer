@echo off
echo === Registering SAPI Voice Token ===

REM Check for admin privileges
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo ERROR: This script must be run as Administrator
    echo Please run as Administrator and try again
    pause
    exit /b 1
)

echo Running with Administrator privileges

REM Register our voice token in SAPI
echo.
echo Registering OpenSpeech Test Voice in SAPI...

REM Create the voice token registry key
reg add "HKLM\SOFTWARE\Microsoft\Speech\Voices\Tokens\OpenSpeechTestVoice" /f
if %errorLevel% neq 0 (
    echo FAILED: Could not create voice token registry key
    pause
    exit /b 1
)

REM Set the CLSID to point to our OpenSpeechSpVoice
reg add "HKLM\SOFTWARE\Microsoft\Speech\Voices\Tokens\OpenSpeechTestVoice" /v "CLSID" /t REG_SZ /d "{F2E8B6A1-3C4D-4E5F-8A7B-9C1D2E3F4A5B}" /f
if %errorLevel% neq 0 (
    echo FAILED: Could not set CLSID
    pause
    exit /b 1
)

REM Set the default value (voice name)
reg add "HKLM\SOFTWARE\Microsoft\Speech\Voices\Tokens\OpenSpeechTestVoice" /ve /t REG_SZ /d "OpenSpeech Test Voice" /f

REM Create attributes subkey
reg add "HKLM\SOFTWARE\Microsoft\Speech\Voices\Tokens\OpenSpeechTestVoice\Attributes" /f

REM Set voice attributes
reg add "HKLM\SOFTWARE\Microsoft\Speech\Voices\Tokens\OpenSpeechTestVoice\Attributes" /v "Name" /t REG_SZ /d "OpenSpeech Test Voice" /f
reg add "HKLM\SOFTWARE\Microsoft\Speech\Voices\Tokens\OpenSpeechTestVoice\Attributes" /v "Language" /t REG_SZ /d "409" /f
reg add "HKLM\SOFTWARE\Microsoft\Speech\Voices\Tokens\OpenSpeechTestVoice\Attributes" /v "Gender" /t REG_SZ /d "Female" /f
reg add "HKLM\SOFTWARE\Microsoft\Speech\Voices\Tokens\OpenSpeechTestVoice\Attributes" /v "Age" /t REG_SZ /d "Adult" /f
reg add "HKLM\SOFTWARE\Microsoft\Speech\Voices\Tokens\OpenSpeechTestVoice\Attributes" /v "Vendor" /t REG_SZ /d "OpenSpeech" /f
reg add "HKLM\SOFTWARE\Microsoft\Speech\Voices\Tokens\OpenSpeechTestVoice\Attributes" /v "Version" /t REG_SZ /d "1.0" /f

echo SUCCESS: Voice token registered successfully!

echo.
echo Testing voice registration...
cscript //NoLogo test-sapi-direct.vbs

echo.
echo === SAPI Voice Registration Complete ===
echo OpenSpeech voice should now be available in SAPI applications!
pause
