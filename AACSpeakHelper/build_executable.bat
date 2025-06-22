@echo off
echo Building AACSpeakHelper with TTS-Wrapper v0.10.11 PyInstaller utilities...

REM Kill any running AACSpeakHelper processes to avoid permission errors
echo Stopping any running AACSpeakHelper processes...
taskkill /f /im AACSpeakHelperServer.exe 2>nul
taskkill /f /im client.exe 2>nul

REM Wait a moment for processes to fully terminate
timeout /t 2 /nobreak >nul

REM Clean up any locked directories
echo Cleaning up previous build directories...
rmdir /s /q "dist\AACSpeakHelperServer" 2>nul
rmdir /s /q "dist\client" 2>nul

REM Get the TTS-Wrapper PyInstaller hooks directory
for /f "tokens=*" %%i in ('uv run python -c "import tts_wrapper; import os; print(os.path.join(os.path.dirname(tts_wrapper.__file__), '_pyinstaller'))"') do set tts_hooks_dir=%%i

REM Echo the paths for debugging
echo TTS-Wrapper hooks directory: %tts_hooks_dir%

REM Show what TTS-Wrapper will automatically include
echo.
echo TTS-Wrapper PyInstaller utilities will automatically handle:
echo - PyAudio and PortAudio DLLs
echo - SoundDevice audio binaries
echo - Azure Speech SDK DLLs
echo - ONNX Runtime DLLs
echo - All other TTS engine dependencies

REM Build Python executables with PyInstaller using TTS-Wrapper optimized configuration
echo.
echo Building AACSpeakHelperServer with TTS-Wrapper PyInstaller hooks...
uv run python -m PyInstaller AACSpeakHelperServer.py --noupx --onedir --noconsole --name "AACSpeakHelperServer" -i .\assets\translate.ico --clean --additional-hooks-dir="%tts_hooks_dir%" --collect-all language_data --collect-all language_tags --collect-all comtypes --collect-all pytz -y

echo Skipping GUI Configure AACSpeakHelper build (excluded from final build)...

uv run python -m PyInstaller client.py --noupx --console --onedir --clean -i .\assets\translate.ico -y

uv run python -m PyInstaller cli_config_creator.py --noupx --console --name "Configure AACSpeakHelper CLI" --onedir --clean -i .\assets\configure.ico -y

uv run python -m PyInstaller CreateGridset.py --noupx --noconsole --onedir --clean -y

REM Run Inno Setup (if available)
if exist "C:\Users\admin.will\AppData\Local\Programs\Inno Setup 6\ISCC.exe" (
    echo Running Inno Setup to create installer...
    "C:\Users\admin.will\AppData\Local\Programs\Inno Setup 6\ISCC.exe" buildscript.iss
) else if exist "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" (
    echo Running Inno Setup to create installer...
    "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" buildscript.iss
) else if exist "C:\Program Files\Inno Setup 6\ISCC.exe" (
    echo Running Inno Setup to create installer...
    "C:\Program Files\Inno Setup 6\ISCC.exe" buildscript.iss
) else (
    echo Inno Setup not found. Skipping installer creation.
    echo To create an installer, please install Inno Setup from: https://jrsoftware.org/isdl.php
    echo Then re-run this build script.
)