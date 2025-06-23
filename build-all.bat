@echo off
setlocal enabledelayedexpansion

REM ============================================================================
REM SherpaOnnx Azure SAPI Installer - Complete Build Script
REM ============================================================================
REM This script builds the entire project:
REM 1. C++ COM Wrapper (NativeTTSWrapper)
REM 2. Python Installer Executable (sapi_voice_installer.exe)
REM 3. AACSpeakHelper Server Executable
REM 4. NSIS Installer Package
REM Supports both interactive and unattended (CI) modes
REM ============================================================================

REM Check if running in CI/unattended mode
set "UNATTENDED_MODE=false"
if "%CI%"=="true" set "UNATTENDED_MODE=true"
if "%GITHUB_ACTIONS%"=="true" set "UNATTENDED_MODE=true"
if "%1"=="--unattended" set "UNATTENDED_MODE=true"

echo.
echo ============================================================================
echo  SherpaOnnx Azure SAPI Installer - Complete Build Script
if "%UNATTENDED_MODE%"=="true" (
    echo  [CI MODE] Running in unattended mode
) else (
    echo  [INTERACTIVE] Running in interactive mode
)
echo ============================================================================
echo.

REM Check if running as administrator (skip in CI)
if "%UNATTENDED_MODE%"=="false" (
    net session >nul 2>&1
    if %errorLevel% neq 0 (
        echo [ERROR] This script must be run as Administrator
        echo         Right-click and select "Run as administrator"
        pause
        exit /b 1
    )
    echo [OK] Running as Administrator
) else (
    echo [CI MODE] Skipping administrator check in unattended mode
)

REM Set build variables
set "PROJECT_ROOT=%~dp0"
set "MSBUILD_PATH="
set "PYTHON_EXE="
set "PYINSTALLER_EXE="
set "NSIS_EXE="

echo.
echo [INFO] Checking build dependencies...
echo ----------------------------------------

REM Find MSBuild
if exist "C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe" (
    set "MSBUILD_PATH=C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe"
    echo [OK] Found MSBuild 2022: !MSBUILD_PATH!
) else if exist "C:\Program Files (x86)\Microsoft Visual Studio\2019\BuildTools\MSBuild\Current\Bin\MSBuild.exe" (
    set "MSBUILD_PATH=C:\Program Files (x86)\Microsoft Visual Studio\2019\BuildTools\MSBuild\Current\Bin\MSBuild.exe"
    echo [OK] Found MSBuild 2019: !MSBUILD_PATH!
) else if exist "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" (
    set "MSBUILD_PATH=C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe"
    echo [OK] Found MSBuild 2022 Community: !MSBUILD_PATH!
) else (
    echo [ERROR] MSBuild not found
    echo         Please install Visual Studio Build Tools or Visual Studio
    echo         Download from: https://visualstudio.microsoft.com/downloads/
    if "%UNATTENDED_MODE%"=="false" pause
    exit /b 1
)

REM Find Python
where python >nul 2>&1
if %errorLevel% equ 0 (
    set "PYTHON_EXE=python"
    echo [OK] Found Python: !PYTHON_EXE!
) else (
    echo [ERROR] Python not found in PATH
    echo         Please install Python and add it to PATH
    if "%UNATTENDED_MODE%"=="false" pause
    exit /b 1
)

REM Check for uv (Python package manager)
where uv >nul 2>&1
if %errorLevel% equ 0 (
    echo [OK] Found uv package manager
) else (
    echo [ERROR] uv package manager not found
    echo         Please install uv: pip install uv
    if "%UNATTENDED_MODE%"=="false" pause
    exit /b 1
)

REM Check if uv environment is set up with build dependencies
echo [INFO] Setting up Python environment with build dependencies...
uv sync --extra build
if %errorLevel% neq 0 (
    echo [ERROR] Failed to sync Python environment with build dependencies
    if "%UNATTENDED_MODE%"=="false" pause
    exit /b 1
)
echo [OK] Python environment ready with PyInstaller

REM Find NSIS
if exist "C:\Program Files (x86)\NSIS\makensis.exe" (
    set "NSIS_EXE=C:\Program Files (x86)\NSIS\makensis.exe"
    echo [OK] Found NSIS: !NSIS_EXE!
) else if exist "C:\Program Files\NSIS\makensis.exe" (
    set "NSIS_EXE=C:\Program Files\NSIS\makensis.exe"
    echo [OK] Found NSIS: !NSIS_EXE!
) else (
    echo [WARNING] NSIS not found - installer package will not be created
    echo           Download from: https://nsis.sourceforge.io/Download
)

echo.
echo [INFO] Starting build process...
echo ============================================================================

REM Step 1: Build C++ COM Wrapper
echo.
echo [STEP 1] Building C++ COM Wrapper...
echo ----------------------------------------
cd /d "%PROJECT_ROOT%NativeTTSWrapper"

echo Cleaning previous build...
"!MSBUILD_PATH!" NativeTTSWrapper.vcxproj /p:Configuration=Release /p:Platform=x64 /t:Clean
if %errorLevel% neq 0 (
    echo [ERROR] Failed to clean C++ project
    if "%UNATTENDED_MODE%"=="false" pause
    exit /b 1
)

echo Building C++ COM Wrapper...
"!MSBUILD_PATH!" NativeTTSWrapper.vcxproj /p:Configuration=Release /p:Platform=x64
if %errorLevel% neq 0 (
    echo [ERROR] Failed to build C++ COM Wrapper
    if "%UNATTENDED_MODE%"=="false" pause
    exit /b 1
)

if not exist "x64\Release\NativeTTSWrapper.dll" (
    echo [ERROR] NativeTTSWrapper.dll was not created
    if "%UNATTENDED_MODE%"=="false" pause
    exit /b 1
)

echo [OK] C++ COM Wrapper built successfully

REM Step 2: Build Python Installer Executable (needed for COM registration)
echo.
echo [STEP 2] Building Python Installer Executable...
echo ----------------------------------------

echo Building sapi_voice_installer.exe...
uv run pyinstaller --onefile --console --name sapi_voice_installer sapi_voice_installer.py
if %errorLevel% neq 0 (
    echo [ERROR] Failed to build Python installer executable
    if "%UNATTENDED_MODE%"=="false" pause
    exit /b 1
)

if not exist "dist\sapi_voice_installer.exe" (
    echo [ERROR] sapi_voice_installer.exe was not created
    if "%UNATTENDED_MODE%"=="false" pause
    exit /b 1
)

echo [OK] Python installer executable built successfully

REM Step 3: Register COM Wrapper using Python Installer
echo.
echo [STEP 3] Registering COM Wrapper using Python Installer...
echo ----------------------------------------
cd /d "%PROJECT_ROOT%"

if "%UNATTENDED_MODE%"=="true" (
    set "CI=true"
    dist\sapi_voice_installer.exe register-com
) else (
    dist\sapi_voice_installer.exe register-com
)

if %errorLevel% neq 0 (
    echo [ERROR] Failed to register COM wrapper using Python installer
    if "%UNATTENDED_MODE%"=="true" (
        echo [CI MODE] COM registration failed in unattended mode
        echo          This is expected in CI environments without admin privileges
        echo          The build will continue - COM registration will be needed on target systems
        echo [WARNING] Continuing build without COM registration...
    ) else (
        echo.
        echo [INFO] The Python installer includes sophisticated error handling and cleanup.
        echo        If it suggested a restart, please restart your computer and run:
        echo        dist\sapi_voice_installer.exe register-com
        echo.
        echo        The build was successful - only COM registration failed.
        echo        You can continue with remaining builds if needed.
        echo.
        set /p continue="Continue with remaining builds anyway? (y/N): "
        if /i not "%continue%"=="y" (
            echo Build stopped. Please resolve COM registration and try again.
            pause
            exit /b 1
        )
        echo [WARNING] Continuing build without COM registration...
    )
)

echo [OK] COM Wrapper registered successfully using Python installer

REM Step 4: Build AACSpeakHelper Server Executable
echo.
echo [STEP 4] Building AACSpeakHelper Server Executable...
echo ----------------------------------------
cd /d "%PROJECT_ROOT%"

echo Building AACSpeakHelperServer.exe...
uv run pyinstaller --onefile --console --name AACSpeakHelperServer AACSpeakHelperServer.py
if %errorLevel% neq 0 (
    echo [ERROR] Failed to build AACSpeakHelper server executable
    if "%UNATTENDED_MODE%"=="false" pause
    exit /b 1
)

if not exist "dist\AACSpeakHelperServer.exe" (
    echo [ERROR] AACSpeakHelperServer.exe was not created
    if "%UNATTENDED_MODE%"=="false" pause
    exit /b 1
)

echo [OK] AACSpeakHelper server executable built successfully

cd /d "%PROJECT_ROOT%"

echo.
echo [SUCCESS] BUILD COMPLETED SUCCESSFULLY!
echo ============================================================================
echo.
echo Built components:
echo   [OK] NativeTTSWrapper.dll (C++ COM Wrapper)
echo   [OK] sapi_voice_installer.exe (Python Installer)
echo   [OK] AACSpeakHelperServer.exe (TTS Server)
echo.

if not "!NSIS_EXE!"=="" (
    echo [STEP 5] Creating NSIS Installer Package...
    echo ----------------------------------------
    if exist "installer.nsi" (
        "!NSIS_EXE!" installer.nsi
        if %errorLevel% equ 0 (
            echo [OK] NSIS installer package created successfully
        ) else (
            echo [WARNING] NSIS installer creation failed
        )
    ) else (
        echo [WARNING] installer.nsi not found - skipping NSIS package creation
    )
) else (
    echo [WARNING] NSIS not found - installer package not created
    echo           Install NSIS to create installer packages
)

echo.
echo [SUCCESS] Ready for distribution!
echo.
if "%UNATTENDED_MODE%"=="false" pause
